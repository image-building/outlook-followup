/* Pure follow-up logic. No Office or network calls here, so it can be unit-tested in Node. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.FUCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var AUTO_SUBJECT = /^\s*(automatic reply|auto(matic)?[- ]?reply|autoreply|out of (the )?office|automatisch antwoord|afwezig|abwesenheitsnotiz|automatische antwort|réponse automatique|undeliverable|onbestelbaar|delivery status notification|mail delivery failed|unzustellbar)\b/i;
  var AUTO_SENDER = /^(postmaster|mailer-daemon|microsoftexchange[0-9a-f]*)@/i;

  /** True when a message is an out-of-office, bounce or other automatic reply (these should not count as an answer). */
  function isAutoReply(msg) {
    var subject = msg.subject || "";
    var from = (msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || "";
    if (AUTO_SUBJECT.test(subject)) return true;
    if (AUTO_SENDER.test(from)) return true;
    var headers = msg.internetMessageHeaders || [];
    for (var i = 0; i < headers.length; i++) {
      var name = (headers[i].name || "").toLowerCase();
      var value = (headers[i].value || "").toLowerCase().trim();
      if (name === "auto-submitted" && value && value !== "no") return true;
      if (name === "x-autoreply" || name === "x-autorespond") return true;
      if (name === "precedence" && /auto_reply|bulk|junk/.test(value)) return true;
    }
    return false;
  }

  function addDays(iso, days) {
    return new Date(new Date(iso).getTime() + days * 86400000).toISOString();
  }

  /** The reminder moment: `days` after `iso`, at local time hh:mm. */
  function reminderMoment(iso, days, hhmm) {
    var d = new Date(new Date(iso).getTime() + days * 86400000);
    var parts = String(hhmm || "09:00").split(":");
    d.setHours(parseInt(parts[0], 10) || 9, parseInt(parts[1], 10) || 0, 0, 0);
    return d.toISOString();
  }

  /**
   * Decide the state of one conversation.
   * messages: [{ id, sentDateTime, isMine, isAuto, isDraft }]
   * Returns { state: 'waiting' | 'answered' | 'empty', anchor, lastReply }
   *   waiting  -> my message is the latest real one; anchor = that message
   *   answered -> someone else replied after my latest message; anchor = my latest message (may be null)
   */
  function computeState(messages) {
    var real = (messages || [])
      .filter(function (m) { return !m.isDraft && !m.isAuto && m.sentDateTime; })
      .sort(function (a, b) { return new Date(a.sentDateTime) - new Date(b.sentDateTime); });
    if (!real.length) return { state: "empty", anchor: null, lastReply: null };
    var latest = real[real.length - 1];
    var mine = real.filter(function (m) { return m.isMine; });
    var lastMine = mine.length ? mine[mine.length - 1] : null;
    if (latest.isMine) return { state: "waiting", anchor: latest, lastReply: null };
    return { state: "answered", anchor: lastMine, lastReply: latest };
  }

  /**
   * Work out which mailbox changes are needed to bring a tracked entry in line with the computed state.
   * entry: stored tracking record { c, s, st, fm, e, d }
   * Returns a list of actions; the caller performs them against Microsoft Graph.
   */
  function planActions(entry, result, opts) {
    var actions = [];
    var days = opts.days, hhmm = opts.reminderTime;
    if (result.state === "waiting") {
      var anchor = result.anchor;
      var due = reminderMoment(anchor.sentDateTime, days, hhmm);
      if (entry.fm && entry.fm !== anchor.id && entry.st === "waiting") {
        actions.push({ type: "unflag", messageId: entry.fm }); // superseded by a newer follow-up of mine
      }
      if (entry.fm !== anchor.id || entry.st !== "waiting" || entry.d !== due) {
        actions.push({ type: "flag", messageId: anchor.id, start: anchor.sentDateTime, due: due });
      }
      if (entry.e) {
        if (entry.d !== due) actions.push({ type: "moveReminder", eventId: entry.e, due: due });
      } else {
        actions.push({ type: "createReminder", due: due });
      }
      return actions;
    }
    if (result.state === "answered") {
      var target = (result.anchor && result.anchor.id) || entry.fm;
      if (target && entry.st !== "answered") {
        actions.push({ type: "complete", messageId: target, at: result.lastReply.sentDateTime });
      }
      if (target && entry.fm && entry.fm !== target && entry.st === "waiting") {
        actions.push({ type: "unflag", messageId: entry.fm });
      }
      if (entry.e) actions.push({ type: "deleteReminder", eventId: entry.e });
      return actions;
    }
    return actions;
  }

  return { isAutoReply: isAutoReply, computeState: computeState, planActions: planActions, addDays: addDays, reminderMoment: reminderMoment };
});
