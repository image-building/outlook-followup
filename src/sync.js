/* Tracking store (Outlook roaming settings, synced across your devices) + Microsoft Graph sync. */
window.FUSync = (function () {
  "use strict";
  var cfg = window.FU_CONFIG;
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var KEY = "fu.v1";
  var PRUNE_ANSWERED_DAYS = 90;

  /* ---------- store ---------- */
  // Entry: { c: conversationId, s: subject, st: 'waiting'|'answered'|'new', fm: flagged message id,
  //          e: reminder event id, d: reminder ISO, t: last activity ISO, w: weblink of anchor }
  function load() {
    var raw = Office.context.roamingSettings.get(KEY);
    var data = raw ? JSON.parse(raw) : {};
    data.items = data.items || [];
    data.ignored = data.ignored || [];
    data.settings = Object.assign({}, cfg.defaults, data.settings || {});
    return data;
  }
  function save(data) {
    var json = JSON.stringify(data);
    if (json.length > 30000) throw new Error("Tracking list is full — stop tracking some old answered emails.");
    Office.context.roamingSettings.set(KEY, json);
    return new Promise(function (resolve, reject) {
      Office.context.roamingSettings.saveAsync(function (r) {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message));
      });
    });
  }

  /* ---------- graph ---------- */
  async function graph(method, path, body, interactive) {
    var token = await FUAuth.getToken(interactive);
    for (var attempt = 0; attempt < 3; attempt++) {
      var res = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, {
        method: method,
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : undefined
      });
      if (res.status === 429 || res.status === 503) {
        var wait = (parseInt(res.headers.get("Retry-After"), 10) || 2) * 1000;
        await new Promise(function (r) { setTimeout(r, wait); });
        continue;
      }
      if (res.status === 204) return null;
      if (res.status === 404 && method !== "GET") return null; // already gone
      var json = await res.json().catch(function () { return null; });
      if (!res.ok) {
        var err = new Error((json && json.error && json.error.message) || ("Graph " + res.status));
        err.status = res.status;
        throw err;
      }
      return json;
    }
    throw new Error("Microsoft Graph is busy, try again in a minute.");
  }

  var ctxCache = null;
  async function context(interactive) {
    if (ctxCache) return ctxCache;
    var me = await graph("GET", "/me?$select=mail,userPrincipalName", null, interactive);
    var sent = await graph("GET", "/me/mailFolders/sentitems?$select=id", null, interactive);
    var addrs = [me.mail, me.userPrincipalName, Office.context.mailbox.userProfile.emailAddress]
      .filter(Boolean).map(function (a) { return a.toLowerCase(); });
    ctxCache = { addresses: addrs, sentFolderId: sent.id };
    return ctxCache;
  }

  async function ensureCategory(interactive) {
    var list = await graph("GET", "/me/outlook/masterCategories", null, interactive);
    var exists = (list.value || []).some(function (c) { return c.displayName.toLowerCase() === cfg.category.toLowerCase(); });
    if (!exists) await graph("POST", "/me/outlook/masterCategories", { displayName: cfg.category, color: "preset7" }, interactive);
  }

  async function conversationMessages(conversationId, ctx, interactive) {
    var filter = encodeURIComponent("conversationId eq '" + conversationId.replace(/'/g, "''") + "'");
    var url = "/me/messages?$filter=" + filter + "&$select=id,subject,from,sentDateTime,isDraft,parentFolderId,webLink,categories&$top=100";
    var out = [];
    while (url) {
      var page = await graph("GET", url, null, interactive);
      out = out.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    out.forEach(function (m) {
      var from = ((m.from && m.from.emailAddress && m.from.emailAddress.address) || "").toLowerCase();
      m.isMine = m.parentFolderId === ctx.sentFolderId || ctx.addresses.indexOf(from) >= 0;
      m.isAuto = !m.isMine && FUCore.isAutoReply(m);
    });
    // Incoming messages newer than my latest one: check their headers for auto-replies (out-of-office etc.)
    var mine = out.filter(function (m) { return m.isMine && !m.isDraft; })
      .map(function (m) { return new Date(m.sentDateTime).getTime(); });
    var lastMine = mine.length ? Math.max.apply(null, mine) : 0;
    var candidates = out.filter(function (m) { return !m.isMine && !m.isAuto && !m.isDraft && new Date(m.sentDateTime).getTime() > lastMine; });
    for (var i = 0; i < candidates.length && i < 5; i++) {
      var h = await graph("GET", "/me/messages/" + candidates[i].id + "?$select=internetMessageHeaders", null, interactive);
      candidates[i].internetMessageHeaders = h.internetMessageHeaders;
      candidates[i].isAuto = FUCore.isAutoReply(candidates[i]);
    }
    return out;
  }

  /* ---------- actions ---------- */
  function utc(iso) { return { dateTime: iso.replace(/Z$/, "").replace(/\.\d+$/, ""), timeZone: "UTC" }; }

  function reminderEvent(entry, due) {
    var end = new Date(new Date(due).getTime() + 15 * 60000).toISOString();
    var link = entry.w ? '<p><a href="' + entry.w + '">Open the email</a></p>' : "";
    return {
      subject: "Follow up: " + (entry.s || "(no subject)"),
      body: { contentType: "HTML", content: "<p>No reply yet to your email <b>" + escapeHtml(entry.s || "") + "</b>.</p>" + link +
        "<p><i>Created by the Follow-up add-in. It is removed automatically when a reply arrives.</i></p>" },
      start: utc(due), end: utc(end),
      isReminderOn: true, reminderMinutesBeforeStart: 0,
      showAs: "free", categories: [cfg.category]
    };
  }

  async function run(action, entry, interactive) {
    switch (action.type) {
      case "flag":
        await graph("PATCH", "/me/messages/" + action.messageId, {
          flag: { flagStatus: "flagged", startDateTime: utc(action.start), dueDateTime: utc(action.due) }
        }, interactive);
        entry.fm = action.messageId;
        break;
      case "unflag":
        await graph("PATCH", "/me/messages/" + action.messageId, { flag: { flagStatus: "notFlagged" } }, interactive);
        break;
      case "complete":
        await graph("PATCH", "/me/messages/" + action.messageId, {
          flag: { flagStatus: "complete", completedDateTime: utc(action.at) }
        }, interactive);
        entry.fm = action.messageId;
        break;
      case "createReminder":
        var ev = await graph("POST", "/me/events", reminderEvent(entry, action.due), interactive);
        entry.e = ev.id;
        break;
      case "moveReminder":
        var moved = await graph("PATCH", "/me/events/" + action.eventId, reminderEvent(entry, action.due), interactive);
        if (!moved) entry.e = (await graph("POST", "/me/events", reminderEvent(entry, action.due), interactive)).id;
        break;
      case "deleteReminder":
        await graph("DELETE", "/me/events/" + action.eventId, null, interactive);
        entry.e = null;
        break;
    }
  }

  async function syncEntry(entry, data, ctx, interactive) {
    var msgs = await conversationMessages(entry.c, ctx, interactive);
    var result = FUCore.computeState(msgs);
    if (result.state === "empty") return;
    var opts = { days: Number(data.settings.days) || 14, reminderTime: data.settings.reminderTime };
    if (result.anchor) entry.w = result.anchor.webLink;
    var newest = msgs.filter(function (m) { return !m.isDraft; }).map(function (m) { return m.sentDateTime; }).sort().pop();
    if (!entry.s) entry.s = (msgs[0] && msgs[0].subject || "").slice(0, 80);
    var actions = FUCore.planActions(entry, result, opts);
    for (var i = 0; i < actions.length; i++) await run(actions[i], entry, interactive);
    entry.st = result.state;
    entry.d = result.state === "waiting" ? FUCore.reminderMoment(result.anchor.sentDateTime, opts.days, opts.reminderTime) : null;
    entry.t = newest || entry.t;
  }

  /** Picks up emails you tagged while composing (they carry the Follow-up category in Sent Items). */
  async function adoptTaggedSent(data, interactive) {
    var since = new Date(Date.now() - 60 * 86400000).toISOString();
    var filter = encodeURIComponent("categories/any(c:c eq '" + cfg.category + "') and sentDateTime ge " + since);
    var res = await graph("GET", "/me/mailFolders/sentitems/messages?$filter=" + filter + "&$select=id,conversationId,subject,sentDateTime&$top=100", null, interactive);
    (res.value || []).forEach(function (m) {
      if (!m.conversationId) return;
      if (data.ignored.indexOf(m.conversationId) >= 0) return;
      if (data.items.some(function (e) { return e.c === m.conversationId; })) return;
      data.items.push({ c: m.conversationId, s: (m.subject || "").slice(0, 80), st: "new", t: m.sentDateTime });
    });
  }

  var running = null;
  /** Full sync. Returns { ok, errors, needsSignIn }. */
  function syncAll(interactive) {
    if (running) return running;
    running = (async function () {
      var data = load();
      var errors = [];
      try {
        var ctx = await context(interactive);
        await adoptTaggedSent(data, interactive);
        for (var i = 0; i < data.items.length; i++) {
          try { await syncEntry(data.items[i], data, ctx, interactive); }
          catch (e) { if (e.code === "SIGNIN_REQUIRED") throw e; errors.push((data.items[i].s || "email") + ": " + e.message); }
        }
        var cutoff = Date.now() - PRUNE_ANSWERED_DAYS * 86400000;
        data.items = data.items.filter(function (e) { return !(e.st === "answered" && new Date(e.t).getTime() < cutoff); });
        data.lastSync = new Date().toISOString();
        await save(data);
        return { ok: true, errors: errors };
      } catch (e) {
        if (e.code === "SIGNIN_REQUIRED") return { ok: false, needsSignIn: true, errors: [] };
        return { ok: false, errors: [e.message] };
      }
    })().finally(function () { running = null; });
    return running;
  }

  async function track(conversationId, subject, interactive) {
    var data = load();
    data.ignored = data.ignored.filter(function (c) { return c !== conversationId; });
    var entry = data.items.find(function (e) { return e.c === conversationId; });
    if (!entry) { entry = { c: conversationId, s: (subject || "").slice(0, 80), st: "new" }; data.items.push(entry); }
    var ctx = await context(interactive);
    await syncEntry(entry, data, ctx, interactive);
    await save(data);
    return entry;
  }

  async function untrack(conversationId, interactive) {
    var data = load();
    var entry = data.items.find(function (e) { return e.c === conversationId; });
    if (entry) {
      if (entry.e) await graph("DELETE", "/me/events/" + entry.e, null, interactive);
      if (entry.fm && entry.st === "waiting") await graph("PATCH", "/me/messages/" + entry.fm, { flag: { flagStatus: "notFlagged" } }, interactive);
      try { // remove the Follow-up category so it isn't picked up again
        var ctx = await context(interactive);
        var msgs = await conversationMessages(conversationId, ctx, interactive);
        for (var i = 0; i < msgs.length; i++) {
          var cats = msgs[i].categories || [];
          if (cats.indexOf(cfg.category) >= 0) {
            await graph("PATCH", "/me/messages/" + msgs[i].id, { categories: cats.filter(function (c) { return c !== cfg.category; }) }, interactive);
          }
        }
      } catch (e) { console.warn(e); }
    }
    data.items = data.items.filter(function (e) { return e.c !== conversationId; });
    data.ignored.push(conversationId);
    data.ignored = data.ignored.slice(-50);
    await save(data);
  }

  async function saveSettings(settings) {
    var data = load();
    data.settings = Object.assign(data.settings, settings);
    data.items.forEach(function (e) { if (e.st === "waiting") e.d = null; }); // forces reminders to be re-timed on next sync
    await save(data);
  }

  function escapeHtml(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  return { load: load, syncAll: syncAll, track: track, untrack: untrack, saveSettings: saveSettings, ensureCategory: ensureCategory, graph: graph };
})();
