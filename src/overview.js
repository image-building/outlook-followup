/* Online overview: one list of follow-ups across all your Microsoft 365 mailboxes. */
(function () {
  "use strict";
  var cfg = window.FU_CONFIG;
  var $ = function (id) { return document.getElementById(id); };
  var COLORS = ["#0f6cbd", "#c239b3", "#107c10", "#ca5010", "#4f6bed", "#986f0b"];
  var AUTO_MS = 5 * 60 * 1000;
  var pca, engines = [], filter = "all", lastSync = 0, signInNeeded = {};

  var redirectUri = cfg.baseUrl + "/";
  // Local mode: the page is served by the helper on your Mac (http://localhost:47200) and gets its sign-in from it,
  // so Safari never has to keep a Microsoft session.
  var LOCAL = /^(localhost|127\.0\.0\.1)$/.test(location.hostname) && location.port === "47200";
  var localAccounts = [], localTokens = {};

  init().catch(function (e) { status(e.message, true); });

  async function api(path, body) {
    var res = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json", "X-Follow-Up": "1" }, body: JSON.stringify(body || {}) });
    var json = await res.json().catch(function () { return {}; });
    if (!res.ok) { var err = new Error(json.error || ("Helper error " + res.status)); err.status = res.status; throw err; }
    return json;
  }
  async function loadLocalAccounts() {
    var r = await api("/api/accounts");
    localAccounts = (r.accounts || []).filter(Boolean).map(function (u) { return { homeAccountId: "local:" + u, username: u }; });
    return r;
  }

  async function init() {
    if (LOCAL) {
      document.title = "Follow-ups";
      $("signinBtn").onclick = addAccount;
      $("addAccountBtn").onclick = addAccount;
      $("syncBtn").onclick = function () { syncAll(); };
      try { await loadLocalAccounts(); } catch (e) { status("The helper on your Mac isn't answering: " + e.message, true); }
      setup();
      if (engines.length) syncAll();
      setInterval(function () { if (!document.hidden) syncAll(); }, AUTO_MS);
      document.addEventListener("visibilitychange", function () {
        if (!document.hidden && Date.now() - lastSync > 2 * 60 * 1000) syncAll();
      });
      return;
    }
    pca = new msal.PublicClientApplication({
      auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: redirectUri },
      cache: { cacheLocation: "localStorage" }
    });
    await pca.initialize();
    try {
      await pca.handleRedirectPromise();
      renewDone(true);
    } catch (e) {
      // a quiet renewal (prompt=none) that needs you to sign in: remember it and show "Sign in again"
      renewDone(false);
      console.warn("quiet sign-in renewal not possible", e && e.errorCode);
    }
    $("signinBtn").onclick = addAccount;
    $("addAccountBtn").onclick = addAccount;
    $("syncBtn").onclick = function () { syncAll(); };
    setup();
    if (engines.length) syncAll();
    setInterval(function () { if (!document.hidden) syncAll(); }, AUTO_MS);
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden && Date.now() - lastSync > 2 * 60 * 1000) syncAll();
    });
  }

  function accounts() {
    if (LOCAL) return localAccounts.slice();
    // One entry per mailbox (MSAL can list the same person once per organisation).
    var seen = {};
    return pca.getAllAccounts().filter(function (a) {
      var k = (a.username || "").toLowerCase();
      if (seen[k]) return false; seen[k] = true; return true;
    }).sort(function (a, b) { return a.username.localeCompare(b.username); });
  }

  function setup() {
    engines = accounts().map(function (account, i) {
      var prefix = "fu:" + account.homeAccountId + ":";
      return {
        account: account,
        color: COLORS[i % COLORS.length],
        engine: FUEngine({
          myAddress: account.username,
          getToken: function (interactive) { return token(account, interactive); },
          store: {
            get: function (k) { try { return localStorage.getItem(prefix + k); } catch (e) { return null; } },
            set: function (k, v) { try { localStorage.setItem(prefix + k, v); } catch (e) { /* private mode: fine, mailbox is the truth */ } }
          }
        })
      };
    });
    var any = engines.length > 0;
    $("welcome").classList.toggle("hidden", any);
    $("accounts").classList.toggle("hidden", !any);
    $("lists").classList.toggle("hidden", !any);
    renderAccounts();
    render();
  }

  async function token(account, interactive) {
    if (LOCAL) {
      var c = localTokens[account.username];
      if (c && c.exp > Date.now() + 60000) return c.token;
      try {
        var r = await api("/api/token", { username: account.username });
        localTokens[account.username] = { token: r.accessToken, exp: r.expiresOn };
        return r.accessToken;
      } catch (e) {
        var se = new Error("Sign-in required: " + e.message); se.code = "SIGNIN_REQUIRED"; throw se;
      }
    }
    var req = { scopes: cfg.scopes, account: account };
    try {
      return (await pca.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (interactive) { await pca.acquireTokenRedirect({ scopes: cfg.scopes, account: account, loginHint: account.username }); }
      var err = new Error("Sign-in required"); err.code = "SIGNIN_REQUIRED"; throw err;
    }
  }

  async function connectGmail() {
    var link = prompt("Paste the connection link from your Gmail script (it starts with https://script.google.com/…):");
    if (!link) return;
    link = link.trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(link) || link.indexOf("key=") < 0) { status("That doesn't look like the connection link from the script.", true); return; }
    var host = engines.find(function (x) { return /@imagebuilding\.com$/i.test(x.account.username); }) || engines[0];
    status("Connecting Gmail… reminders will go to the calendar of " + host.account.username);
    try { await host.engine.setGmail(link, true); } catch (e) { status(e.message, true); return; }
    await syncAll();
  }

  async function disconnectGmail() {
    var host = gmailHost();
    if (!host || !confirm("Disconnect Gmail? Its reminders are removed from your calendar. (Labels in Gmail stay.)")) return;
    status("Disconnecting Gmail…");
    try { await host.engine.setGmail(null, true); status(""); } catch (e) { status(e.message, true); }
    renderAccounts(); render();
  }

  async function addAccountLocal() {
    var before = localAccounts.map(function (a) { return a.username; });
    try { await api("/api/signin"); } catch (e) { status(e.message, true); return; }
    status("A browser window opened on your Mac — sign in there with the account to add, then come back here.");
    var until = Date.now() + 5 * 60 * 1000;
    while (Date.now() < until) {
      await new Promise(function (r) { setTimeout(r, 3000); });
      var r = await loadLocalAccounts().catch(function () { return null; });
      if (!r) continue;
      if (r.signin && r.signin.error) { status("Sign-in failed: " + r.signin.error, true); return; }
      var added = localAccounts.some(function (a) { return before.indexOf(a.username) < 0; });
      if (added || (r.signin && !r.signin.running && r.signin.last)) { status(""); setup(); syncAll(); return; }
    }
    status("Sign-in timed out — try + Add account again.", true);
  }

  function addAccount() {
    if (LOCAL) return addAccountLocal();
    pca.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" });
  }

  async function removeAccount(acc) {
    if (!confirm("Remove " + acc.username + " from this page? (Nothing changes in the mailbox.)")) return;
    try {
      Object.keys(localStorage).forEach(function (k) { if (k.indexOf("fu:" + acc.homeAccountId + ":") === 0) localStorage.removeItem(k); });
    } catch (e) { /* ignore */ }
    if (LOCAL) {
      try { await api("/api/remove", { username: acc.username }); await loadLocalAccounts(); } catch (e) { status(e.message, true); return; }
    } else {
      await pca.clearCache({ account: acc });
    }
    if (filter === acc.username) filter = "all";
    setup();
  }

  /* ---------- sync ---------- */
  var syncing = false;
  /* ---------- quiet sign-in renewal ----------
   * Microsoft lets a web page keep its sign-in for 24 hours. After that, Safari blocks the hidden renewal, so we
   * renew with a quick redirect to Microsoft and back (prompt=none: no login screen, no MFA, if your Microsoft
   * session is remembered). Tried at most once per account per hour, so it can never loop. */
  function ssGet(k, d) { try { return JSON.parse(sessionStorage.getItem(k)) || d; } catch (e) { return d; } }
  function ssSet(k, v) { try { sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* ignore */ } }
  function renewDone(ok) {
    var who = ssGet("fu:renewing", null);
    if (!who) return;
    ssSet("fu:renewing", null);
    var failed = ssGet("fu:renewFailed", {});
    if (ok) delete failed[who]; else failed[who] = Date.now();
    ssSet("fu:renewFailed", failed);
  }
  function tryQuietRenewal() {
    if (LOCAL) return false;
    var tried = ssGet("fu:renewTried", {}), failed = ssGet("fu:renewFailed", {});
    var hour = 60 * 60 * 1000;
    var x = engines.find(function (e) {
      var u = e.account.username;
      return signInNeeded[u] && !(failed[u] && Date.now() - failed[u] < hour) && !(tried[u] && Date.now() - tried[u] < hour);
    });
    if (!x) return false;
    tried[x.account.username] = Date.now();
    ssSet("fu:renewTried", tried);
    ssSet("fu:renewing", x.account.username);
    status("Refreshing your sign-in for " + x.account.username + "…");
    pca.acquireTokenRedirect({ scopes: cfg.scopes, account: x.account, loginHint: x.account.username, prompt: "none" });
    return true;
  }

  async function syncAll() {
    if (syncing || !engines.length) return;
    syncing = true;
    lastSync = Date.now();
    $("syncBtn").disabled = true;
    status("Checking " + engines.length + (engines.length === 1 ? " mailbox" : " mailboxes") + " for replies…");
    var errors = [];
    await Promise.all(engines.map(async function (x) {
      var res = await x.engine.syncAll(false);
      signInNeeded[x.account.username] = !!res.needsSignIn;
      res.errors.forEach(function (e) { errors.push(x.account.username + " — " + e); });
    }));
    syncing = false;
    $("syncBtn").disabled = false;
    renderAccounts();
    render();
    status(errors.join(" · "), errors.length > 0);
    if (!document.hidden) tryQuietRenewal();
  }

  /* ---------- render ---------- */
  function renderAccounts() {
    var box = $("accountChips");
    box.innerHTML = "";
    engines.forEach(function (x) {
      var chip = document.createElement("span");
      var needs = signInNeeded[x.account.username];
      chip.className = "chip" + (needs ? " err" : "");
      chip.innerHTML = '<span class="dot"></span><span class="n"></span>' + (needs ? '<button class="re">Sign in again</button>' : "") + '<button class="rm" title="Remove from this page">✕</button>';
      chip.querySelector(".dot").style.background = x.color;
      chip.querySelector(".n").textContent = x.account.username;
      if (needs) chip.querySelector(".re").onclick = function () {
        if (LOCAL) { addAccount(); return; } // signs that account in again through the helper
        token(x.account, true);
      };
      chip.querySelector(".rm").onclick = function () { removeAccount(x.account); };
      box.appendChild(chip);
    });
    var g = gmailSource();
    if (g) {
      var gc = document.createElement("span");
      gc.className = "chip" + (g.error ? " err" : "");
      gc.innerHTML = '<span class="dot"></span><span class="n"></span><button class="rm" title="Disconnect Gmail">✕</button>';
      gc.querySelector(".dot").style.background = g.color;
      gc.querySelector(".n").textContent = g.account.username + " · Gmail" + (g.error ? " — " + g.error : "");
      gc.querySelector(".rm").onclick = disconnectGmail;
      box.appendChild(gc);
    }
    var im = imapSource();
    if (im) {
      var ic = document.createElement("span");
      ic.className = "chip" + (im.warn ? " err" : "");
      ic.innerHTML = '<span class="dot"></span><span class="n"></span>';
      ic.querySelector(".dot").style.background = im.color;
      ic.querySelector(".n").textContent = im.account.username + " · IMAP" + (im.warn ? " — " + im.warn : "");
      ic.title = "Checked by the helper on your Mac every 15 minutes";
      box.appendChild(ic);
    }
    var f = $("filters");
    f.innerHTML = "";
    var srcs = sources();
    if (srcs.length < 2) return;
    [{ k: "all", label: "All" }].concat(srcs.map(function (x) { return { k: x.account.username, label: x.account.username.split("@")[1] || x.account.username }; }))
      .forEach(function (o) {
        var b = document.createElement("button");
        b.textContent = o.label;
        b.className = filter === o.k ? "on" : "";
        b.onclick = function () { filter = o.k; renderAccounts(); render(); };
        f.appendChild(b);
      });
  }

  function fmtDate(iso) { return new Date(iso).toLocaleDateString(undefined, { day: "numeric", month: "short" }); }
  function describe(e) {
    if (e.st === "waiting") {
      var overdue = e.d && new Date(e.d) <= new Date();
      if (e.ro) return (overdue ? "Reopened — reminder was due " : "Waiting (reopened) · reminder ") + (e.d ? fmtDate(e.d) : "pending");
      return overdue ? "No reply — reminder was due " + fmtDate(e.d) : "Waiting · reminder " + (e.d ? fmtDate(e.d) : "pending");
    }
    if (e.st === "nomine") return "No email from you found in this thread" + (e.x && e.x.length ? " (from: " + e.x.join(", ") + ")" : "");
    if (e.st === "answered" && e.m) return "Marked done " + fmtDate(e.m);
    if (e.st === "answered") return "Answered " + (e.t ? fmtDate(e.t) : "");
    return "Checking…";
  }

  /** The engine that holds the Gmail link (its calendar gets the Gmail reminders). */
  function gmailHost() {
    return engines.find(function (x) { return !!x.engine.load().settings.gmail; }) || null;
  }
  function gmailSource() {
    var host = gmailHost();
    var g = host && host.engine.gmail();
    if (!host || !g) return null;
    return {
      account: { username: g.account || "Gmail" }, color: "#d93025", gmail: true, host: host,
      items: (g.items || []).map(function (it) { return Object.assign({ c: it.id }, it); }),
      error: g.error,
      setDone: function (id, done) { return host.engine.gmailAction(id, done ? "done" : "undone", true); },
      untrack: function (id) { return host.engine.gmailAction(id, "untrack", true); }
    };
  }
  function imapSource() {
    var host = engines.find(function (x) { return !!x.engine.imap(); });
    var d = host && host.engine.imap();
    if (!d) return null;
    var stale = d.lastRun && (Date.now() - new Date(d.lastRun).getTime() > 60 * 60 * 1000);
    var warn = d.error ? d.error : (stale ? "helper last ran " + fmtWhen(d.lastRun) + " (is your Mac on?)" : "");
    return {
      account: { username: d.account || "IMAP" }, color: "#0b8a8a", host: host, warn: warn, queued: d.queued,
      items: (d.items || []).map(function (it) { return Object.assign({ c: it.id }, it); }),
      setDone: function (id, done) { return host.engine.imapAction(id, done ? "done" : "undone", true); },
      untrack: function (id) { return host.engine.imapAction(id, "untrack", true); }
    };
  }
  function fmtWhen(iso) {
    var d = new Date(iso);
    return d.toLocaleDateString(undefined, { day: "numeric", month: "short" }) + " " + d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
  }

  function sources() {
    var list = engines.map(function (x) {
      return { account: x.account, color: x.color, items: x.engine.load().items,
        setDone: function (c, done) { return x.engine.setDone(c, done, true); },
        untrack: function (c) { return x.engine.untrack(c, true); } };
    });
    var g = gmailSource();
    if (g) list.push(g);
    var im = imapSource();
    if (im) list.push(im);
    return list;
  }

  function render() {
    var all = [];
    sources().forEach(function (x) {
      if (filter !== "all" && filter !== x.account.username) return;
      x.items.forEach(function (e) { all.push({ e: e, x: x }); });
    });
    var waiting = all.filter(function (r) { return r.e.st !== "answered"; })
      .sort(function (a, b) { return new Date(a.e.d || 8.64e15) - new Date(b.e.d || 8.64e15); });
    var answered = all.filter(function (r) { return r.e.st === "answered"; })
      .sort(function (a, b) { return new Date(b.e.t || 0) - new Date(a.e.t || 0); });
    fill("waitingList", waiting, "Nothing waiting for a reply.");
    fill("answeredList", answered, "No answered emails yet.");
    $("waitingCount").textContent = waiting.length ? "(" + waiting.length + ")" : "";
    $("answeredCount").textContent = answered.length ? "(" + answered.length + ")" : "";
  }

  function fill(id, rows, emptyText) {
    var ul = $(id);
    ul.innerHTML = "";
    if (!rows.length) { var li0 = document.createElement("li"); li0.className = "empty"; li0.textContent = emptyText; ul.appendChild(li0); return; }
    rows.forEach(function (r) {
      var e = r.e, x = r.x;
      var done = e.st === "answered";
      var overdue = e.st === "waiting" && e.d && new Date(e.d) <= new Date();
      var li = document.createElement("li");
      li.className = (done ? "answered" : "") + (overdue ? " overdue" : "");
      li.innerHTML = '<button class="box" role="checkbox" aria-checked="' + done + '" title="' + (done ? "Mark as waiting again" : "Mark as done") + '">' + (done ? "✓" : "") + '</button>' +
        '<span class="t"><a target="_blank" rel="noopener"></a><span class="meta"></span></span><button class="x" title="Stop tracking">✕</button>';
      var a = li.querySelector("a");
      a.textContent = e.s || "(no subject)";
      if (e.w) a.href = e.w; else a.removeAttribute("href");
      var meta = li.querySelector(".meta");
      if (sources().length > 1) {
        var tag = document.createElement("span"); tag.className = "tag"; tag.style.background = x.color;
        tag.textContent = x.account.username.split("@")[1] || x.account.username;
        meta.appendChild(tag);
      }
      meta.appendChild(document.createTextNode(describe(e) + (e.pending ? " · updating within 15 min" : "")));
      if (e.lp) {
        var lp = document.createElement("span");
        lp.className = "latest";
        lp.textContent = "↳ " + e.lp.f + " · " + fmtDate(e.lp.d) + (e.lp.p ? " — “" + e.lp.p + "”" : "");
        li.querySelector(".t").appendChild(lp);
      }
      li.querySelector(".box").onclick = async function (ev) {
        var btn = ev.currentTarget; btn.disabled = true; btn.textContent = done ? "" : "✓";
        status(done ? "Moving back to waiting…" : "Marking as done…");
        try {
          await x.setDone(e.c, !done);
          status("");
        } catch (err) { status(err.message, true); }
        render();
      };
      li.querySelector(".x").onclick = async function () {
        if (!confirm("Stop tracking “" + (e.s || "this email") + "”?")) return;
        status("Removing…");
        try { await x.untrack(e.c); status(""); } catch (err) { status(err.message, true); }
        render();
      };
      ul.appendChild(li);
    });
  }

  function status(text, isError) {
    var f = $("footer");
    f.innerHTML = "";
    var span = document.createElement("span");
    if (isError) span.className = "err";
    span.textContent = text || (lastSync ? "Last checked " + new Date(lastSync).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }) + " · checks every 5 minutes while this page is open" : "");
    f.appendChild(span);
  }
})();
