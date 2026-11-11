const assert = require("assert");
const C = require("../src/core.js");

const me = (id, t) => ({ id, sentDateTime: t, isMine: true });
const them = (id, t, extra = {}) => ({ id, sentDateTime: t, isMine: false, ...extra });
const opts = { days: 14, reminderTime: "09:00" };
let n = 0;
const test = (name, fn) => { fn(); n++; console.log("ok -", name); };

test("my email is latest -> waiting", () => {
  const r = C.computeState([me("a", "2026-11-01T10:00:00Z")]);
  assert.equal(r.state, "waiting"); assert.equal(r.anchor.id, "a");
});

test("reply after my email -> answered (checked off)", () => {
  const r = C.computeState([me("a", "2026-11-01T10:00:00Z"), them("b", "2026-11-02T10:00:00Z")]);
  assert.equal(r.state, "answered"); assert.equal(r.anchor.id, "a");
});

test("I send again after the reply -> waiting again on the new message", () => {
  const r = C.computeState([me("a", "2026-11-01T10:00:00Z"), them("b", "2026-11-02T10:00:00Z"), me("c", "2026-11-03T10:00:00Z")]);
  assert.equal(r.state, "waiting"); assert.equal(r.anchor.id, "c");
});

test("out-of-office reply does not count", () => {
  const ooo = them("b", "2026-11-02T10:00:00Z", { subject: "Automatisch antwoord: Offerte" });
  ooo.isAuto = C.isAutoReply(ooo);
  assert.equal(C.computeState([me("a", "2026-11-01T10:00:00Z"), ooo]).state, "waiting");
});

test("auto-reply detected via headers", () => {
  assert.ok(C.isAutoReply({ subject: "Re: hi", internetMessageHeaders: [{ name: "Auto-Submitted", value: "auto-replied" }] }));
  assert.ok(!C.isAutoReply({ subject: "Re: hi", internetMessageHeaders: [{ name: "Auto-Submitted", value: "no" }] }));
  assert.ok(C.isAutoReply({ subject: "Undeliverable: hi", from: { emailAddress: { address: "x@y.com" } } }));
  assert.ok(!C.isAutoReply({ subject: "Re: Out of scope items", from: { emailAddress: { address: "bob@client.com" } } }));
});

test("drafts are ignored", () => {
  const r = C.computeState([me("a", "2026-11-01T10:00:00Z"), them("b", "2026-11-02T10:00:00Z"), { ...me("d", "2026-11-04T10:00:00Z"), isDraft: true }]);
  assert.equal(r.state, "answered");
});

test("reminder is 14 days later at 09:00 local", () => {
  const due = new Date(C.reminderMoment("2026-11-01T15:30:00Z", 14, "09:00"));
  assert.equal(due.getHours(), 9); assert.equal(due.getMinutes(), 0);
  const days = Math.round((due - new Date("2026-11-01T00:00:00Z")) / 86400000);
  assert.ok(days === 14 || days === 15);
});

test("new tracking -> flag + create reminder", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z")]);
  const acts = C.planActions({ c: "x", st: "new" }, res, opts).map(a => a.type);
  assert.deepEqual(acts, ["flag", "createReminder"]);
});

test("waiting, nothing changed -> no actions", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z")]);
  const d = C.reminderMoment("2026-11-01T10:00:00Z", 14, "09:00");
  assert.deepEqual(C.planActions({ c: "x", st: "waiting", fm: "a", e: "ev", d }, res, opts), []);
});

test("reply arrives -> complete flag (checkmark) + delete reminder", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z"), them("b", "2026-11-02T10:00:00Z")]);
  const acts = C.planActions({ c: "x", st: "waiting", fm: "a", e: "ev", d: "z" }, res, opts).map(a => a.type);
  assert.deepEqual(acts, ["complete", "deleteReminder"]);
});

test("send again after answered -> flag new message + new reminder, old checkmark kept", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z"), them("b", "2026-11-02T10:00:00Z"), me("c", "2026-11-05T10:00:00Z")]);
  const acts = C.planActions({ c: "x", st: "answered", fm: "a", e: null }, res, opts);
  assert.deepEqual(acts.map(a => a.type), ["flag", "createReminder"]);
  assert.equal(acts[0].messageId, "c");
});

test("I nudge while still waiting -> old flag removed, reminder moved", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z"), me("c", "2026-11-08T10:00:00Z")]);
  const d = C.reminderMoment("2026-11-01T10:00:00Z", 14, "09:00");
  const acts = C.planActions({ c: "x", st: "waiting", fm: "a", e: "ev", d }, res, opts).map(a => a.type);
  assert.deepEqual(acts, ["unflag", "flag", "moveReminder"]);
});

test("nudge then reply before a sync -> newest checked, older unflagged", () => {
  const res = C.computeState([me("a", "2026-11-01T10:00:00Z"), me("c", "2026-11-08T10:00:00Z"), them("b", "2026-11-09T10:00:00Z")]);
  const acts = C.planActions({ c: "x", st: "waiting", fm: "a", e: "ev", d: "z" }, res, opts);
  assert.deepEqual(acts.map(a => a.type), ["complete", "unflag", "deleteReminder"]);
  assert.equal(acts[0].messageId, "c");
});

console.log(`\n${n} tests passed`);
