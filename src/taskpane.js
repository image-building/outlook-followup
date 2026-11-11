/* Task pane UI. */
(function () {
  "use strict";
  var cfg = window.FU_CONFIG;
  var $ = function (id) { return document.getElementById(id); };
  var AUTO_SYNC_MS = 3 * 60 * 1000;
  var lastAuto = 0;

  Office.onReady(async function () {
    var rs = Office.context.roamingSettings;
    window.FUSync = FUEngine({
      getToken: FUAuth.getToken,
      myAddress: Office.context.mailbox.userProfile.emailAddress,
      store: {
        get: function (k) { return rs.get(k); },
        set: function (k, v) {
          rs.set(k, v);
          return new Promise(function (resolve, reject) {
            rs.saveAsync(function (r) { r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message)); });
          });
        }
      }
    });
    wireUi();
    render();
    await refreshCurrent();
    // When the pane is pinned it stays open while you move between emails — refresh and sync on each switch.
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { refreshCurrent(); autoSync(); });
    }
    autoSync();
  });

  function wireUi() {
    $("syncBtn").onclick = function () { doSync(true); };
    $("signinBtn").onclick = function () { doSync(true); };
    $("trackBtn").onclick = onTrackClick;
    $("composeToggle").onchange = onComposeToggle;
    $("saveSettings").onclick = async function () {
      await FUSync.saveSettings({ days: Math.max(1, parseInt($("days").value, 10) || 14), reminderTime: $("reminderTime").value || "09:00" });
      status("Settings saved — updating reminders…");
      doSync(true);
    };
    var s = FUSync.load().settings;
    $("days").value = s.days;
    $("reminderTime").value = s.reminderTime;
  }

  /* ---------- sync ---------- */
  function autoSync() {
    if (Date.now() - lastAuto < AUTO_SYNC_MS) return;
    lastAuto = Date.now();
    doSync(false);
  }

  async function doSync(interactive) {
    $("syncBtn").disabled = true;
    status("Checking for replies…");
    var res = await FUSync.syncAll(interactive);
    $("syncBtn").disabled = false;
    $("signin").classList.toggle("hidden", !res.needsSignIn);
    if (res.needsSignIn) status("");
    else if (res.errors.length) status(res.errors.join(" · "), true);
    else status("");
    render();
    refreshCurrent();
    var st = FUSync.load().settings;
    if (document.activeElement !== $("days")) $("days").value = st.days;
    if (document.activeElement !== $("reminderTime")) $("reminderTime").value = st.reminderTime;
  }

  /* ---------- current email ---------- */
  function item() { return Office.context.mailbox.item; }
  function isCompose() { var it = item(); return it && typeof it.subject !== "string"; }

  async function refreshCurrent() {
    var it = item();
    $("current").classList.add("hidden");
    $("compose").classList.add("hidden");
    if (!it) return;
    if (isCompose()) {
      if (!it.categories || !Office.context.requirements.isSetSupported("Mailbox", "1.8")) return; // e.g. older builds
      $("compose").classList.remove("hidden");
      if (it.categories) {
        it.categories.getAsync(function (r) {
          var cats = (r.status === Office.AsyncResultStatus.Succeeded && r.value) || [];
          $("composeToggle").checked = cats.some(function (c) { return c.displayName === cfg.category; });
        });
      }
      return;
    }
    $("current").classList.remove("hidden");
    $("currentSubject").textContent = it.subject || "(no subject)";
    var entry = FUSync.load().items.find(function (e) { return e.c === it.conversationId; });
    $("currentStatus").textContent = entry ? describe(entry) : "Not tracked.";
    $("trackBtn").textContent = entry ? "Stop tracking" : "Track for follow-up";
    $("trackBtn").classList.toggle("primary", !entry);
  }

  async function onTrackClick() {
    var it = item();
    if (!it || !it.conversationId) return;
    var btn = $("trackBtn");
    btn.disabled = true;
    try {
      var tracked = FUSync.load().items.some(function (e) { return e.c === it.conversationId; });
      if (tracked) { status("Removing…"); await FUSync.untrack(it.conversationId, true); }
      else { status("Tracking…"); await FUSync.track(it.conversationId, it.subject, true); }
      status("");
    } catch (e) {
      status(e.message, true);
    }
    btn.disabled = false;
    render();
    refreshCurrent();
  }

  async function onComposeToggle(ev) {
    var on = ev.target.checked;
    var it = item();
    try {
      if (on) await FUSync.ensureCategory(true);
      await new Promise(function (resolve, reject) {
        var fn = on ? it.categories.addAsync : it.categories.removeAsync;
        fn.call(it.categories, [cfg.category], function (r) {
          r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message));
        });
      });
      status(on ? "Will be tracked once sent." : "");
    } catch (e) {
      ev.target.checked = !on;
      status(e.message, true);
    }
  }

  /* ---------- list ---------- */
  function fmtDate(iso) {
    return new Date(iso).toLocaleDateString(undefined, { day: "numeric", month: "short" });
  }
  function describe(e) {
    if (e.st === "waiting") {
      var overdue = e.d && new Date(e.d) <= new Date();
      return overdue ? "No reply — reminder was due " + fmtDate(e.d) : "Waiting for reply · reminder " + (e.d ? fmtDate(e.d) : "pending");
    }
    if (e.st === "answered" && e.m) return "Marked done " + fmtDate(e.m);
    if (e.st === "answered") return "Answered " + (e.t ? fmtDate(e.t) : "");
    return "Tracked — checking…";
  }

  function render() {
    var data = FUSync.load();
    var waiting = data.items.filter(function (e) { return e.st !== "answered"; })
      .sort(function (a, b) { return new Date(a.d || 0) - new Date(b.d || 0); });
    var answered = data.items.filter(function (e) { return e.st === "answered"; })
      .sort(function (a, b) { return new Date(b.t || 0) - new Date(a.t || 0); });
    fill("waitingList", waiting, "Nothing waiting. Tag a sent email to track it.");
    fill("answeredList", answered, "No answered emails yet.");
    $("waitingCount").textContent = waiting.length ? "(" + waiting.length + ")" : "";
    $("answeredCount").textContent = answered.length ? "(" + answered.length + ")" : "";
    if (data.lastSync && !$("footer").querySelector(".err")) {
      $("footer").textContent = "Last checked " + new Date(data.lastSync).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
    }
  }

  function fill(id, entries, emptyText) {
    var ul = $(id);
    ul.innerHTML = "";
    if (!entries.length) { var li = document.createElement("li"); li.className = "empty"; li.textContent = emptyText; ul.appendChild(li); return; }
    entries.forEach(function (e) {
      var li = document.createElement("li");
      var overdue = e.st === "waiting" && e.d && new Date(e.d) <= new Date();
      li.className = (e.st === "answered" ? "answered" : "") + (overdue ? " overdue" : "");
      var done = e.st === "answered";
      li.innerHTML = '<button class="box" role="checkbox" aria-checked="' + done + '" title="' + (done ? "Mark as waiting again" : "Mark as done") + '">' + (done ? "✓" : "") + '</button><span class="t"></span><button class="x" title="Stop tracking">✕</button>';
      var t = li.querySelector(".t");
      t.textContent = e.s || "(no subject)";
      var meta = document.createElement("span"); meta.className = "meta"; meta.textContent = describe(e); t.appendChild(meta);
      t.onclick = function () { openEntry(e); };
      li.querySelector(".box").onclick = async function (ev) {
        ev.stopPropagation();
        var box = ev.currentTarget;
        box.disabled = true;
        box.textContent = done ? "" : "✓"; // instant feedback
        status(done ? "Moving back to waiting…" : "Marking as done…");
        try { await FUSync.setDone(e.c, !done, true); status(""); } catch (err) { status(err.message, true); }
        render(); refreshCurrent();
      };
      li.querySelector(".x").onclick = async function (ev) {
        ev.stopPropagation();
        status("Removing…");
        try { await FUSync.untrack(e.c, true); status(""); } catch (err) { status(err.message, true); }
        render(); refreshCurrent();
      };
      ul.appendChild(li);
    });
  }

  function isMobile() {
    var host = Office.context.mailbox.diagnostics && Office.context.mailbox.diagnostics.hostName;
    return host === "OutlookIOS" || host === "OutlookAndroid";
  }

  function openEntry(e) {
    if (isMobile()) { copySubject(e.s || ""); return; } // Outlook for iOS/Android can't open another email from an add-in
    try {
      if (!isMobile() && e.fm && Office.context.mailbox.convertToEwsId) {
        var ews = Office.context.mailbox.convertToEwsId(e.fm, Office.MailboxEnums.RestVersion.v2_0);
        Office.context.mailbox.displayMessageForm(ews);
        return;
      }
    } catch (err) { console.warn(err); }
    if (e.w) Office.context.ui.openBrowserWindow ? Office.context.ui.openBrowserWindow(e.w) : window.open(e.w, "_blank");
  }

  function copySubject(text) {
    var done = function () { status("Subject copied — paste it in Outlook's search to find the email."); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else fallback();
    function fallback() {
      var ta = document.createElement("textarea");
      ta.value = text; ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select(); ta.setSelectionRange(0, text.length);
      var ok = false; try { ok = document.execCommand("copy"); } catch (x) {}
      document.body.removeChild(ta);
      ok ? done() : status("Search Outlook for: " + text);
    }
  }

  function status(text, isError) {
    var f = $("footer");
    f.innerHTML = "";
    if (!text) { render(); return; }
    var span = document.createElement("span");
    if (isError) span.className = "err";
    span.textContent = text;
    f.appendChild(span);
  }
})();
