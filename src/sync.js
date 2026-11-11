/* Follow-up engine for one mailbox: Microsoft Graph sync + a small local cache.
 * The mailbox is the source of truth (Follow-up category, flags, reminder events);
 * the cache only remembers ids so fewer calls are needed.
 *
 * opts: { getToken(interactive) -> Promise<string>,
 *         store: { get() -> string|null, set(string) -> Promise },
 *         myAddress: optional extra address that counts as "me" } */
window.FUEngine = function (opts) {
  "use strict";
  var cfg = window.FU_CONFIG;
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var KEY = "fu.v1";
  var PRUNE_ANSWERED_DAYS = 90;
  var STORE_VERSION = 3;
  // Hidden property on each reminder event that links it to its conversation.
  var CONV_PROP = "String {00020329-0000-0000-C000-000000000046} Name FollowUpConversationId";
  // Hidden property on the mailbox's root folder holding your settings, so every device and the web page agree.
  var SETTINGS_PROP = "String {00020329-0000-0000-C000-000000000046} Name FollowUpSettings";

  /* ---------- store ---------- */
  // Entry: { c: conversationId, s: subject, st: 'waiting'|'answered'|'new', fm: flagged message id,
  //          e: reminder event id, d: reminder ISO, t: last activity ISO, w: weblink of anchor }
  function load() {
    var raw = null;
    try { raw = opts.store.get(KEY); } catch (e) { raw = null; }
    var data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (e) { data = {}; }
    data.items = data.items || [];
    data.ignored = data.ignored || [];
    data.settings = Object.assign({}, cfg.defaults, data.settings || {});
    return data;
  }
  function save(data) {
    var json = JSON.stringify(data);
    if (json.length > 30000) throw new Error("Tracking list is full — stop tracking some old answered emails.");
    return Promise.resolve(opts.store.set(KEY, json));
  }

  /* ---------- graph ---------- */
  async function graph(method, path, body, interactive) {
    var token = await opts.getToken(interactive);
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
  var SENT_NAMES = /^(sent|sent items|sent messages|sent mail|verzonden|verzonden items|verzonden berichten|elementos enviados|enviados|correo enviado|gesendet|gesendete elemente|gesendete objekte|éléments envoyés|envoyés|posta inviata|elementi inviati)$/i;

  async function context(interactive) {
    if (ctxCache) return ctxCache;
    var me;
    try { me = await graph("GET", "/me?$select=mail,userPrincipalName,proxyAddresses,otherMails", null, interactive); }
    catch (e) { if (e.code === "SIGNIN_REQUIRED") throw e; me = await graph("GET", "/me?$select=mail,userPrincipalName", null, interactive); }
    var sent = await graph("GET", "/me/mailFolders/sentitems?$select=id", null, interactive);
    var sentIds = [sent.id];
    try { // other "sent" folders, e.g. created by a phone or IMAP app
      var folders = await graph("GET", "/me/mailFolders?$select=id,displayName&$top=200", null, interactive);
      (folders.value || []).forEach(function (f) { if (SENT_NAMES.test((f.displayName || "").trim()) && sentIds.indexOf(f.id) < 0) sentIds.push(f.id); });
    } catch (e) { if (e.code === "SIGNIN_REQUIRED") throw e; }
    var addrs = [me.mail, me.userPrincipalName, opts.myAddress]
      .concat((me.proxyAddresses || []).map(function (p) { return String(p).replace(/^smtp:/i, ""); }))
      .concat(me.otherMails || [])
      .filter(Boolean).map(function (a) { return a.toLowerCase(); });
    ctxCache = { addresses: addrs, sentFolderIds: sentIds, sentFolderId: sent.id, email: (me.mail || me.userPrincipalName || "").toLowerCase() };
    return ctxCache;
  }

  async function readMailboxSettings(interactive) {
    var f = await graph("GET", "/me/mailFolders/msgfolderroot?$select=id&$expand=singleValueExtendedProperties($filter=" +
      encodeURIComponent("id eq '" + SETTINGS_PROP + "'") + ")", null, interactive);
    var p = f && (f.singleValueExtendedProperties || [])[0];
    try { return p ? JSON.parse(p.value) : null; } catch (e) { return null; }
  }
  async function writeMailboxSettings(settings, interactive) {
    await graph("PATCH", "/me/mailFolders/msgfolderroot", {
      singleValueExtendedProperties: [{ id: SETTINGS_PROP, value: JSON.stringify({ days: settings.days, reminderTime: settings.reminderTime }) }]
    }, interactive);
  }

  async function ensureCategory(interactive) {
    var list = await graph("GET", "/me/outlook/masterCategories", null, interactive);
    var exists = (list.value || []).some(function (c) { return c.displayName.toLowerCase() === cfg.category.toLowerCase(); });
    if (!exists) await graph("POST", "/me/outlook/masterCategories", { displayName: cfg.category, color: "preset7" }, interactive);
  }

  async function conversationMessages(conversationId, ctx, interactive) {
    var filter = encodeURIComponent("conversationId eq '" + conversationId.replace(/'/g, "''") + "'");
    var url = "/me/messages?$filter=" + filter + "&$select=id,subject,from,sentDateTime,isDraft,parentFolderId,webLink,categories,flag&$top=100";
    var out = [];
    while (url) {
      var page = await graph("GET", url, null, interactive);
      out = out.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    out.forEach(function (m) {
      var from = ((m.from && m.from.emailAddress && m.from.emailAddress.address) || "").toLowerCase();
      m.isMine = ctx.sentFolderIds.indexOf(m.parentFolderId) >= 0 || ctx.addresses.indexOf(from) >= 0;
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
      showAs: "free", categories: [cfg.category],
      singleValueExtendedProperties: [{ id: CONV_PROP, value: entry.c }]
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

  function hasCategory(msgs) {
    return msgs.some(function (m) { return (m.categories || []).indexOf(cfg.category) >= 0; });
  }

  /** Puts the Follow-up category on one message of the conversation, so the mailbox itself records that it's tracked. */
  async function markConversation(msgs, interactive) {
    if (hasCategory(msgs)) return;
    var real = msgs.filter(function (m) { return !m.isDraft; })
      .sort(function (a, b) { return new Date(b.sentDateTime) - new Date(a.sentDateTime); });
    var target = real.find(function (m) { return m.isMine; }) || real[0];
    if (!target) return;
    var cats = (target.categories || []).concat([cfg.category]);
    await graph("PATCH", "/me/messages/" + target.id, { categories: cats }, interactive);
    target.categories = cats;
  }

  async function syncEntry(entry, data, ctx, interactive, msgs) {
    msgs = msgs || await conversationMessages(entry.c, ctx, interactive);
    var result = FUCore.computeState(msgs);
    if (result.state === "empty") return;
    if (result.state === "nomine") {
      // Nothing from you in this thread: show it, but don't pretend it's answered. A ✓ on the latest email = done by hand.
      var fl = result.anchor.flag;
      entry.x = msgs.map(function (m) { return ((m.from && m.from.emailAddress && m.from.emailAddress.address) || "").toLowerCase(); })
        .filter(function (a, i, arr) { return a && arr.indexOf(a) === i; }).slice(0, 3);
      entry.w = result.anchor.webLink;
      if (!entry.s) entry.s = (result.anchor.subject || "").slice(0, 80);
      if (entry.e) { await run({ type: "deleteReminder", eventId: entry.e }, entry, interactive); }
      if (fl && fl.flagStatus === "complete") {
        entry.st = "answered";
        entry.m = (fl.completedDateTime && fl.completedDateTime.dateTime ? fl.completedDateTime.dateTime.replace(/(\.\d{3})\d*/, "$1").replace(/Z?$/, "Z") : null) || entry.m || new Date().toISOString();
      } else {
        entry.st = "nomine";
        delete entry.m;
      }
      entry.d = null;
      entry.t = result.anchor.sentDateTime;
      return;
    }
    delete entry.x;
    // Outlook's own flag is the truth for "done by hand": a ✓ on my latest email while no reply has come in.
    var f = result.state === "waiting" && result.anchor.flag;
    if (f && f.flagStatus === "complete") {
      entry.m = (f.completedDateTime && f.completedDateTime.dateTime ? f.completedDateTime.dateTime.replace(/(\.\d{3})\d*/, "$1").replace(/Z?$/, "Z") : null) || entry.m || new Date().toISOString();
      if (new Date(entry.m) < new Date(result.anchor.sentDateTime)) entry.m = result.anchor.sentDateTime;
    } else {
      delete entry.m;
    }
    result = FUCore.applyManual(entry, result);
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

  function addEntry(data, conversationId, subject, when) {
    if (!conversationId || data.ignored.indexOf(conversationId) >= 0) return null;
    var found = data.items.find(function (e) { return e.c === conversationId; });
    if (found) return found;
    var entry = { c: conversationId, s: (subject || "").slice(0, 80), st: "new", t: when };
    data.items.push(entry);
    return entry;
  }

  /** Every conversation with a Follow-up category (tagged while composing, or tracked on any device). */
  async function adoptTagged(data, interactive) {
    var since = new Date(Date.now() - 120 * 86400000).toISOString();
    var filter = encodeURIComponent("categories/any(c:c eq '" + cfg.category + "') and receivedDateTime ge " + since);
    var url = "/me/messages?$filter=" + filter + "&$select=id,conversationId,subject,sentDateTime&$top=100";
    var pages = 0;
    while (url && pages++ < 5) {
      var res = await graph("GET", url, null, interactive);
      (res.value || []).forEach(function (m) { addEntry(data, m.conversationId, m.subject, m.sentDateTime); });
      url = res["@odata.nextLink"] || null;
    }
  }

  /** Finds the add-in's reminder events. Returns { conversationId: [eventId, ...] } and adopts conversations that only have a reminder. */
  async function reminderEvents(data, interactive) {
    var start = new Date(Date.now() - 60 * 86400000).toISOString();
    var end = new Date(Date.now() + 130 * 86400000).toISOString();
    var url = "/me/calendarView?startDateTime=" + start + "&endDateTime=" + end +
      "&$select=id,subject,categories,body&$top=200&$expand=singleValueExtendedProperties($filter=" + encodeURIComponent("id eq '" + CONV_PROP + "'") + ")";
    var map = {};
    var pages = 0;
    while (url && pages++ < 10) {
      var res = await graph("GET", url, null, interactive);
      for (var i = 0; i < (res.value || []).length; i++) {
        var ev = res.value[i];
        if ((ev.categories || []).indexOf(cfg.category) < 0 || !/^Follow up: /.test(ev.subject || "")) continue;
        var prop = (ev.singleValueExtendedProperties || [])[0];
        var conv = prop && prop.value;
        if (!conv) conv = await conversationFromEventBody(ev, interactive); // reminders made by version 1.0–1.2
        if (!conv) continue;
        (map[conv] = map[conv] || []).push(ev.id);
        addEntry(data, conv, (ev.subject || "").replace(/^Follow up: /, ""), null);
      }
      url = res["@odata.nextLink"] || null;
    }
    return map;
  }

  async function conversationFromEventBody(ev, interactive) {
    var html = (ev.body && ev.body.content) || "";
    var m = html.match(/ItemID=([^&"'\s<]+)/i);
    if (!m) return null;
    var raw = decodeURIComponent(m[1]);
    var candidates = [raw, raw.replace(/\//g, "-").replace(/\+/g, "_")];
    for (var i = 0; i < candidates.length; i++) {
      try {
        var msg = await graph("GET", "/me/messages/" + encodeURIComponent(candidates[i]) + "?$select=conversationId", null, interactive);
        if (msg && msg.conversationId) {
          await graph("PATCH", "/me/events/" + ev.id, { singleValueExtendedProperties: [{ id: CONV_PROP, value: msg.conversationId }] }, interactive).catch(function () {});
          return msg.conversationId;
        }
      } catch (e) { /* try next form */ }
    }
    try {
      var tr = await graph("POST", "/me/translateExchangeIds", { inputIds: [raw], sourceIdType: "ewsId", targetIdType: "restId" }, interactive);
      var rid = tr && tr.value && tr.value[0] && tr.value[0].targetId;
      if (rid) {
        var msg2 = await graph("GET", "/me/messages/" + encodeURIComponent(rid) + "?$select=conversationId", null, interactive);
        return msg2 && msg2.conversationId;
      }
    } catch (e) { /* give up */ }
    return null;
  }

  var running = null;
  /** Full sync. Returns { ok, errors, needsSignIn }. */
  function syncAll(interactive) {
    if (running) return running;
    running = (async function () {
      var data = load();
      var migrating = (data.v || 1) < STORE_VERSION;
      var errors = [];
      try {
        var ctx = await context(interactive);
        try {
          var ms = await readMailboxSettings(interactive);
          if (ms) data.settings = Object.assign(data.settings, ms);
        } catch (e) { console.warn("settings", e); }
        await adoptTagged(data, interactive);
        var events = await reminderEvents(data, interactive);
        var keep = [];
        for (var i = 0; i < data.items.length; i++) {
          var entry = data.items[i];
          try {
            // One reminder per conversation: reuse the one that exists, delete duplicates.
            var evs = events[entry.c] || [];
            if (evs.length) {
              if (evs.indexOf(entry.e) < 0) entry.e = evs[0];
              for (var j = 0; j < evs.length; j++) if (evs[j] !== entry.e) await graph("DELETE", "/me/events/" + evs[j], null, interactive);
            }
            var msgs = await conversationMessages(entry.c, ctx, interactive);
            if (!hasCategory(msgs)) {
              if (!migrating && entry.st !== "new") { // stopped tracking on another device
                if (entry.e) await graph("DELETE", "/me/events/" + entry.e, null, interactive);
                continue;
              }
              await markConversation(msgs, interactive);
            }
            await syncEntry(entry, data, ctx, interactive, msgs);
            keep.push(entry);
          } catch (e) {
            if (e.code === "SIGNIN_REQUIRED") throw e;
            keep.push(entry);
            errors.push((entry.s || "email") + ": " + e.message);
          }
        }
        var cutoff = Date.now() - PRUNE_ANSWERED_DAYS * 86400000;
        data.items = keep.filter(function (e) { return !(e.st === "answered" && new Date(e.t).getTime() < cutoff); });
        data.v = STORE_VERSION;
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
    var entry = addEntry(data, conversationId, subject, null);
    var ctx = await context(interactive);
    var msgs = await conversationMessages(conversationId, ctx, interactive);
    try { await ensureCategory(interactive); } catch (e) { console.warn(e); }
    await markConversation(msgs, interactive);
    await syncEntry(entry, data, ctx, interactive, msgs);
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

  /** done=true: you tick it off yourself (✓ on Outlook's flag). done=false: back to waiting. */
  async function setDone(conversationId, done, interactive) {
    var data = load();
    var entry = data.items.find(function (e) { return e.c === conversationId; });
    if (!entry) return;
    var ctx = await context(interactive);
    var msgs = await conversationMessages(conversationId, ctx, interactive);
    var anchor = FUCore.computeState(msgs).anchor;
    if (!anchor) throw new Error("Couldn't find the emails of this conversation in your mailbox.");
    if (anchor) {
      var now = new Date().toISOString();
      var flag = done ? { flagStatus: "complete", completedDateTime: { dateTime: now.replace(/Z$/, ""), timeZone: "UTC" } } : { flagStatus: "notFlagged" };
      await graph("PATCH", "/me/messages/" + anchor.id, { flag: flag }, interactive);
      anchor.flag = done ? { flagStatus: "complete", completedDateTime: { dateTime: now } } : flag;
    }
    if (!done) { entry.st = "new"; entry.fm = null; } // fresh flag + reminder
    await syncEntry(entry, data, ctx, interactive, msgs);
    await save(data);
  }

  async function saveSettings(settings) {
    var data = load();
    data.settings = Object.assign(data.settings, settings);
    await writeMailboxSettings(data.settings, true);
    data.items.forEach(function (e) { if (e.st === "waiting") e.d = null; }); // forces reminders to be re-timed on next sync
    await save(data);
  }

  async function settings(interactive) {
    var data = load();
    try { var ms = await readMailboxSettings(interactive); if (ms) data.settings = Object.assign(data.settings, ms); } catch (e) { /* keep cached */ }
    return data.settings;
  }

  async function whoAmI(interactive) { return (await context(interactive)).email; }

  function escapeHtml(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  return { load: load, setDone: setDone, syncAll: syncAll, track: track, untrack: untrack, saveSettings: saveSettings,
    settings: settings, whoAmI: whoAmI, ensureCategory: ensureCategory, graph: graph };
};
