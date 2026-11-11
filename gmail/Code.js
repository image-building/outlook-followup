/* ============================================================================
 * Follow-up Tracker — Gmail part (Google Apps Script)
 *
 * Give an email the Gmail label "Follow-up" to track it. Every 15 minutes this script
 * checks those threads: a reply -> answered (label "Follow-up/Answered"), a newer email
 * from you -> waiting again. The overview page and the Outlook add-in read the result
 * through this script's web address and put the reminders in your Outlook calendar.
 *
 * Setup: see the README ("Gmail").
 * ========================================================================== */

var LABEL = "Follow-up";
var ANSWERED_LABEL = "Follow-up/Answered";
var LOOKBACK_DAYS = 120;

/** Run this once after deploying (Run ▸ setup). It creates the labels, the 15-minute timer and your link. */
function setup() {
  getLabel_(LABEL);
  getLabel_(ANSWERED_LABEL);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "checkFollowUps") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkFollowUps").timeBased().everyMinutes(15).create();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("KEY")) props.setProperty("KEY", Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, ""));
  checkFollowUps();
  showConnectionLink();
}

/** Prints the link you paste into the overview page (Connect Gmail). Keep it private. */
function showConnectionLink() {
  var url = ScriptApp.getService().getUrl();
  if (!url) { Logger.log("Deploy the script as a web app first (Deploy ▸ New deployment ▸ Web app), then run this again."); return; }
  Logger.log("Your connection link (paste it in the Follow-ups page → + Connect Gmail):\n\n" +
    url + "?key=" + PropertiesService.getScriptProperties().getProperty("KEY") + "\n");
}

/** Makes a new link and invalidates the old one. */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty("KEY");
  setup();
}

/* ---------- web address used by the overview page / Outlook add-in ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var key = PropertiesService.getScriptProperties().getProperty("KEY");
  if (!key || p.key !== key) return out_({ ok: false, error: "Wrong or missing key — run showConnectionLink in the script for the right link." });
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (p.action && p.id) act_(p.action, p.id);
    var items = checkFollowUps();
    return out_({ ok: true, account: Session.getEffectiveUser().getEmail(), lastRun: new Date().toISOString(), items: items });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- the check ---------- */
function checkFollowUps() {
  var me = myAddresses_();
  var store = loadStore_();
  var answeredLabel = getLabel_(ANSWERED_LABEL);
  var threads = GmailApp.search("label:" + LABEL.replace(/ /g, "-") + " newer_than:" + LOOKBACK_DAYS + "d", 0, 200);
  var items = [];
  var seen = {};
  threads.forEach(function (thread) {
    var id = thread.getId();
    seen[id] = true;
    var item = evaluate_(thread, me, store[id] || {});
    store[id] = item.manual;
    var isAnswered = item.st === "answered";
    var hasLabel = thread.getLabels().some(function (l) { return l.getName() === ANSWERED_LABEL; });
    if (isAnswered && !hasLabel) thread.addLabel(answeredLabel);
    if (!isAnswered && hasLabel) thread.removeLabel(answeredLabel);
    delete item.manual;
    items.push(item);
  });
  Object.keys(store).forEach(function (id) { if (!seen[id]) delete store[id]; });
  saveStore_(store);
  return items;
}

function evaluate_(thread, me, manual) {
  var account = Session.getEffectiveUser().getEmail();
  var msgs = thread.getMessages().filter(function (m) { return !m.isDraft(); }).map(function (m) {
    var fromRaw = m.getFrom() || "";
    var addr = (fromRaw.match(/<([^>]+)>/) || [null, fromRaw])[1].trim().toLowerCase();
    var name = fromRaw.replace(/<[^>]+>/, "").replace(/"/g, "").trim() || addr;
    var isMine = me.indexOf(addr) >= 0;
    var headers = [];
    if (!isMine) ["Auto-Submitted", "X-Autoreply", "X-Autorespond", "Precedence"].forEach(function (h) {
      var v = m.getHeader(h);
      if (v) headers.push({ name: h, value: v });
    });
    var msg = { id: m.getId(), subject: m.getSubject(), from: { emailAddress: { address: addr, name: name } },
      sentDateTime: m.getDate().toISOString(), isMine: isMine, internetMessageHeaders: headers, gm: m };
    msg.isAuto = !isMine && FUCore.isAutoReply(msg);
    return msg;
  });
  var r = FUCore.computeState(msgs);
  var item = { id: thread.getId(), s: (thread.getFirstMessageSubject() || "").slice(0, 80), st: r.state, a: null, t: null, m: null, ro: false,
    w: "https://mail.google.com/mail/u/?authuser=" + encodeURIComponent(account) + "#all/" + thread.getId(), lp: preview_(msgs), x: null, manual: {} };
  var real = msgs.filter(function (m) { return !m.isAuto; });
  var latest = real.length ? real[real.length - 1] : null;
  item.t = latest ? latest.sentDateTime : null;
  manual = manual || {};

  if (r.state === "waiting") {
    if (manual.m && new Date(r.anchor.sentDateTime) <= new Date(manual.m)) { item.st = "answered"; item.m = manual.m; item.manual.m = manual.m; }
    else item.a = r.anchor.sentDateTime;
  } else if (r.state === "answered") {
    if (manual.ro && new Date(r.lastReply.sentDateTime) <= new Date(manual.ro)) { item.st = "waiting"; item.ro = true; item.a = manual.ro; item.manual.ro = manual.ro; }
  } else if (r.state === "nomine") {
    item.x = msgs.map(function (m) { return m.from.emailAddress.address; }).filter(function (a, i, arr) { return arr.indexOf(a) === i; }).slice(0, 3);
    if (manual.m && latest && new Date(latest.sentDateTime) <= new Date(manual.m)) { item.st = "answered"; item.m = manual.m; item.manual.m = manual.m; }
  }
  return item;
}

function act_(action, id) {
  var thread = GmailApp.getThreadById(id);
  if (!thread) throw new Error("Thread not found");
  var store = loadStore_();
  var cur = evaluate_(thread, myAddresses_(), store[id] || {});
  var now = new Date().toISOString();
  var manual = store[id] || {};
  if (action === "done") {
    if (cur.ro) delete manual.ro; else manual.m = now;
  } else if (action === "undone") {
    if (cur.m) delete manual.m; else if (cur.st === "answered") manual.ro = now;
  } else if (action === "untrack") {
    thread.removeLabel(getLabel_(LABEL));
    thread.removeLabel(getLabel_(ANSWERED_LABEL));
    delete store[id];
    saveStore_(store);
    return;
  }
  store[id] = manual;
  saveStore_(store);
}

/* ---------- helpers ---------- */
function preview_(msgs) {
  var real = msgs.filter(function (m) { return !m.isAuto; });
  var m = real.length ? real[real.length - 1] : msgs[msgs.length - 1];
  if (!m) return null;
  var body = String(m.gm ? m.gm.getPlainBody() : "");
  // cut at the start of the quoted earlier email ("On … wrote:", "Op … schreef:", "-----Original Message-----", "From:/Van:")
  var cut = body.search(/^\s*((On|Op|El|Am|Le) .{3,160}(wrote|schreef|escribió|schrieb|a écrit).*:|-{2,}\s*(Original Message|Oorspronkelijk bericht|Mensaje original)|(From|Van|De|Von): .+)\s*$/im);
  var text = (cut >= 0 ? body.slice(0, cut) : body).replace(/^>.*$/gm, "").replace(/\s+/g, " ").trim();
  if (text.length > 110) text = text.slice(0, 107).replace(/\s+\S*$/, "") + "…";
  return { f: m.isMine ? "You" : m.from.emailAddress.name.slice(0, 40), d: m.sentDateTime, p: text };
}

function myAddresses_() {
  return [Session.getEffectiveUser().getEmail()].concat(GmailApp.getAliases()).map(function (a) { return a.toLowerCase(); });
}

function getLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function loadStore_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("STORE") || "{}"); } catch (e) { return {}; }
}
function saveStore_(store) {
  PropertiesService.getScriptProperties().setProperty("STORE", JSON.stringify(store));
}
