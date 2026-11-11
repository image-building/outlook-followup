/* Online overview: one list of follow-ups across all your Microsoft 365 mailboxes. */
(function () {
  "use strict";
  var cfg = window.FU_CONFIG;
  var $ = function (id) { return document.getElementById(id); };
  var COLORS = ["#0f6cbd", "#c239b3", "#107c10", "#ca5010", "#4f6bed", "#986f0b"];
  var AUTO_MS = 5 * 60 * 1000;
  var pca, engines = [], filter = "all", lastSync = 0, signInNeeded = {};

  var redirectUri = cfg.baseUrl + "/";

  init().catch(function (e) { status(e.message, true); });

  async function init() {
    pca = new msal.PublicClientApplication({
      auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: redirectUri },
      cache: { cacheLocation: "localStorage" }
    });
    await pca.initialize();
    await pca.handleRedirectPromise();
    $("signinBtn").onclick = addAccount;
    $("addAccountBtn").onclick = addAccount;
    $("gmailBtn").onclick = connectGmail;
    $("syncBtn").onclick = function () { syncAll(); };
    setup();
    if (engines.length) syncAll();
    setInterval(function () { if (!document.hidden) syncAll(); }, AUTO_MS);
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden && Date.now() - lastSync > 2 * 60 * 1000) syncAll();
    });
  }

  function accounts() {
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

  function addAccount() {
    pca.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" });
  }

  async function removeAccount(acc) {
    if (!confirm("Remove " + acc.username + " from this page? (Nothing changes in the mailbox.)")) return;
    try {
      Object.keys(localStorage).forEach(function (k) { if (k.indexOf("fu:" + acc.homeAccountId + ":") === 0) localStorage.removeItem(k); });
    } catch (e) { /* ignore */ }
    await pca.clearCache({ account: acc });
    if (filter === acc.username) filter = "all";
    setup();
  }

  /* ---------- sync ---------- */
  var syncing = false;
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
      if (needs) chip.querySelector(".re").onclick = function () { token(x.account, true); };
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
    $("gmailBtn").classList.toggle("hidden", !!gmailHost() || !engines.length);
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
  function sources() {
    var list = engines.map(function (x) {
      return { account: x.account, color: x.color, items: x.engine.load().items,
        setDone: function (c, done) { return x.engine.setDone(c, done, true); },
        untrack: function (c) { return x.engine.untrack(c, true); } };
    });
    var g = gmailSource();
    if (g) list.push(g);
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
      meta.appendChild(document.createTextNode(describe(e)));
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
