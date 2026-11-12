// Tests voor ms365_contact_days: tellingen per adres per dag, zonder inhoud.
import { test } from "node:test";
import assert from "node:assert/strict";
import { amsterdamDay, contactWindow, countMail, countChat, countEvents, otherMember, formatChannel, DayCounts, isAutomated, normalizeAddress } from "../dist/shared/contact-days.js";
import { graphGetPages } from "../dist/auth/graph-client.js";

const SELF = new Set(["test@example.org", "alias@example.org"]);
const NOW = new Date("2026-11-12T15:00:00Z");
const W = contactWindow(NOW, 4);
const adr = (address, name = "") => ({ emailAddress: { address, name } });

test("dagen tellen in Europe/Amsterdam, ook rond middernacht en de wintertijd", () => {
  assert.equal(amsterdamDay("2026-11-11T23:30:00Z"), "2026-11-12");
  assert.equal(amsterdamDay("2026-07-01T21:59:00Z"), "2026-07-01");
  assert.equal(amsterdamDay("2026-07-01T22:00:00Z"), "2026-07-02");
  assert.deepEqual(W, { from: "2026-11-09", to: "2026-11-12", since: "2026-11-08T21:00:00Z" });
  assert.equal(contactWindow(NOW, 0).from, "2026-11-12");
  assert.equal(contactWindow(NOW, 9999).from, "2026-08-15", "begrensd op 90 dagen");
  assert.equal(contactWindow(new Date("2026-03-01T00:30:00Z"), 3).from, "2026-02-27");
});

test("adressen: klein geschreven, geen interne Exchange-paden, automatische afzenders herkend", () => {
  assert.equal(normalizeAddress(" Piet@Example.COM "), "piet@example.com");
  assert.equal(normalizeAddress("/O=EXCHANGELABS/OU=X/CN=RECIPIENTS/CN=ABC"), "");
  assert.equal(normalizeAddress(undefined), "");
  for (const a of ["noreply@example.com", "no-reply@example.com", "do_not_reply@example.com", "mailer-daemon@example.com", "bounce+123@example.com"]) assert.equal(isAutomated(a), true, a);
  for (const a of ["piet@example.com", "norbert@example.com", "info@example.com"]) assert.equal(isAutomated(a), false, a);
});

test("mail: verzonden per geadresseerde, ontvangen alleen van wie je zelf schreef en alleen als je in Aan staat", () => {
  const sent = [
    { sentDateTime: "2026-11-12T09:00:00Z", toRecipients: [adr("Piet@Example.com", "Jansen, Piet"), adr("roos@example.com", "Roos Visser")] },
    { sentDateTime: "2026-11-12T10:00:00Z", toRecipients: [adr("piet@example.com", "Piet oud")] },
    // Buiten het venster, maar wel iemand aan wie je schreef: zijn antwoord telt.
    { sentDateTime: "2026-10-20T10:00:00Z", toRecipients: [adr("kees@example.com", "Kees de Boer")] },
    { sentDateTime: "2026-11-11T10:00:00Z", toRecipients: Array.from({ length: 11 }, (_, i) => adr(`p${i}@example.com`)) },
    { sentDateTime: "2026-11-11T10:00:00Z", toRecipients: [adr("noreply@example.com"), adr("test@example.org")] },
    { sentDateTime: "2026-11-11T10:00:00Z", isDraft: true, toRecipients: [adr("concept@example.com")] },
  ];
  const received = [
    { receivedDateTime: "2026-11-11T23:30:00Z", from: adr("kees@example.com", "Kees de Boer"), toRecipients: [adr("ALIAS@example.org")] },
    { receivedDateTime: "2026-11-10T08:00:00Z", from: adr("piet@example.com", "Piet Jansen"), toRecipients: [adr("test@example.org"), adr("x@example.com")] },
    // Alleen in Cc (niet in Aan), onbekende afzender, van jezelf: telt niet.
    { receivedDateTime: "2026-11-10T08:00:00Z", from: adr("roos@example.com"), toRecipients: [adr("ander@example.com")] },
    { receivedDateTime: "2026-11-10T08:00:00Z", from: adr("nieuwsbrief@example.com"), toRecipients: [adr("test@example.org")] },
    { receivedDateTime: "2026-11-10T08:00:00Z", from: adr("test@example.org"), toRecipients: [adr("test@example.org")] },
  ];
  const tekst = formatChannel("mail", W, countMail(sent, received, SELF, W));
  assert.deepEqual(tekst.split("\n"), [
    "# contactdagen mail 2026-11-09..2026-11-12 regels=4",
    '["mail","kees@example.com","Kees de Boer","2026-11-12",1,0]',
    '["mail","piet@example.com","Jansen, Piet","2026-11-10",1,0]',
    '["mail","piet@example.com","Jansen, Piet","2026-11-12",0,2]',
    '["mail","roos@example.com","Roos Visser","2026-11-12",0,1]',
  ]);
});

test("teams: alleen echte berichten van een één-op-één-chat, in en uit", () => {
  const ander = otherMember([{ userId: "me", email: "test@example.org", displayName: "Ik" }, { userId: "u2", email: "Hilco@Example.com", displayName: "Hilco Lunsing" }], "me");
  assert.deepEqual(ander, { address: "hilco@example.com", name: "Hilco Lunsing" });
  assert.equal(otherMember([{ userId: "me", email: "test@example.org" }, { userId: "bot" }], "me"), null, "zonder adres geen regel");
  const counts = new DayCounts(W, SELF);
  countChat(counts, ander, [
    { messageType: "message", createdDateTime: "2026-11-12T08:00:00Z", from: { user: { id: "u2" } } },
    { messageType: "message", createdDateTime: "2026-11-12T08:01:00Z", from: { user: { id: "me" } } },
    { messageType: "message", createdDateTime: "2026-11-12T08:02:00Z", from: { user: { id: "me" } } },
    { messageType: "systemEventMessage", createdDateTime: "2026-11-12T08:03:00Z", from: null },
    { messageType: "message", createdDateTime: "2026-11-12T08:04:00Z", deletedDateTime: "2026-11-12T08:05:00Z", from: { user: { id: "u2" } } },
    { messageType: "message", createdDateTime: "2026-11-01T08:00:00Z", from: { user: { id: "u2" } } },
  ], "me");
  assert.deepEqual(formatChannel("teams", W, counts).split("\n"), ["# contactdagen teams 2026-11-09..2026-11-12 regels=1", '["teams","hilco@example.com","Hilco Lunsing","2026-11-12",1,2]']);
});

test("agenda: kleine afspraken, niet geannuleerd, niet de hele dag, niet afgewezen", () => {
  const ik = adr("test@example.org", "Ik");
  const ev = (extra) => ({ start: { dateTime: "2026-11-10T14:00:00.0000000" }, organizer: adr("niels@example.com"), responseStatus: { response: "accepted" }, attendees: [ik, adr("niels@example.com", "Niels Bezemer")], ...extra });
  const counts = countEvents([
    ev({}),
    ev({ start: { dateTime: "2026-11-11T09:00:00.0000000" }, organizer: ik, responseStatus: { response: "organizer" }, attendees: [ik, adr("niels@example.com", "Niels Bezemer"), { ...adr("weg@example.com"), status: { response: "declined" } }, { ...adr("zaal@example.com"), type: "resource" }] }),
    ev({ isCancelled: true }),
    ev({ isAllDay: true }),
    ev({ responseStatus: { response: "declined" } }),
    ev({ responseStatus: { response: "notResponded" } }),
    ev({ responseStatus: { response: "none" } }),
    ev({ attendees: [ik] }),
    ev({ attendees: [ik, ...Array.from({ length: 6 }, (_, i) => adr(`d${i}@example.com`))] }),
  ], SELF, W);
  assert.deepEqual(formatChannel("agenda", W, counts).split("\n"), [
    "# contactdagen agenda 2026-11-09..2026-11-12 regels=2",
    '["agenda","niels@example.com","Niels Bezemer","2026-11-10",1,0]',
    '["agenda","niels@example.com","Niels Bezemer","2026-11-11",0,1]',
  ]);
  assert.equal(formatChannel("agenda", W, countEvents([], SELF, W)), "# contactdagen agenda 2026-11-09..2026-11-12 regels=0");
});

test("bladeren stopt bij de bovengrens of zodra de rest niet meer nodig is, en meldt afkappen", async () => {
  const pages = { "/a": { value: [3, 2], "@odata.nextLink": "https://graph.microsoft.com/v1.0/b" }, "https://graph.microsoft.com/v1.0/b": { value: [1, 0], "@odata.nextLink": "https://graph.microsoft.com/v1.0/c" }, "https://graph.microsoft.com/v1.0/c": { value: [-1] } };
  const seen = [];
  const get = async (url) => (seen.push(url), pages[url]);
  assert.deepEqual(await graphGetPages("/a", { maxPages: 5, get }), { value: [3, 2, 1, 0, -1], pages: 3, truncated: false });
  assert.deepEqual(await graphGetPages("/a", { maxPages: 2, get }), { value: [3, 2, 1, 0], pages: 2, truncated: true });
  seen.length = 0;
  const gestopt = await graphGetPages("/a", { maxPages: 5, get, more: (page) => page.at(-1) > 0 });
  assert.deepEqual([gestopt.value, gestopt.truncated, seen.length], [[3, 2, 1, 0], false, 2], "na de pagina die onder de grens eindigt wordt niet verder gebladerd");
});
