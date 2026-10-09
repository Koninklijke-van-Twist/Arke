"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const Core = require("../web/js/core.js");

function fnv1a64Reference(str) {
  let h = 0xcbf29ce484222325n;
  for (const b of Buffer.from(str, "utf8")) {
    h ^= BigInt(b);
    h = (h * 0x100000001b3n) & 0xffffffffffffffffn;
  }
  return h.toString(16).padStart(16, "0");
}

test("messageKey = FNV-1a 64 van het genormaliseerde internetMessageId", () => {
  for (const id of ["<abc@x>", "<CAF1234+Ünïcödé@mail.gmail.com>", "<a@b>", "x".repeat(500)]) {
    assert.equal(Core.messageKey(id), fnv1a64Reference(Core.normalizeMessageId(id)));
  }
});

test("messageKey negeert <>, hoofdletters en spaties", () => {
  assert.equal(Core.messageKey(" <ABC@Example.COM> "), Core.messageKey("abc@example.com"));
  assert.notEqual(Core.messageKey("<a@x>"), Core.messageKey("<b@x>"));
  assert.match(Core.messageKey("<a@x>"), /^[0-9a-f]{16}$/);
});

test("messageKey gooit zonder internetMessageId", () => {
  assert.throws(() => Core.messageKey(""), /internetMessageId/);
  assert.throws(() => Core.messageKey(null), /internetMessageId/);
});

test("sanitizeFileNamePart: verboden tekens, prefixen, accenten, lengte", () => {
  assert.equal(Core.sanitizeFileNamePart("RE: FW: Offerte #123: pomp/klep?"), "Offerte-123-pomp-klep");
  assert.equal(Core.sanitizeFileNamePart("Antw: Café “Zuid”"), "Cafe-“Zuid”");
  assert.equal(Core.sanitizeFileNamePart(""), "zonder-onderwerp");
  assert.equal(Core.sanitizeFileNamePart("   ...   "), "zonder-onderwerp");
  assert.equal(Core.sanitizeFileNamePart("CON"), "zonder-onderwerp");
  const long = Core.sanitizeFileNamePart("a".repeat(200), 60);
  assert.equal(long.length, 60);
  assert.doesNotMatch(Core.sanitizeFileNamePart('a"b*c:d<e>f?g/h\\i|j#k%l&m~n'), /["*:<>?\/\\|#%&~]/);
});

test("buildBaseName is deterministisch en in UTC", () => {
  const a = Core.buildBaseName({ date: "2026-10-09T07:32:10Z", subject: "Re: Levering", internetMessageId: "<m1@kvt.nl>" });
  const b = Core.buildBaseName({ date: new Date("2026-10-09T09:32:10+02:00"), subject: "Re: Levering", internetMessageId: "<M1@kvt.nl>" });
  assert.equal(a, b);
  assert.match(a, /^2026-10-09_0732_Levering_[0-9a-f]{16}$/);
});

test("buildMetadata levert het sidecar-schema", () => {
  const meta = Core.buildMetadata({
    item: {
      subject: "Offerte pomp",
      from: { displayName: "Jan Klant", emailAddress: "Jan@Klant.NL" },
      to: [{ displayName: "Tim", emailAddress: "tfalken@kvt.nl" }, null],
      cc: [],
      dateTimeCreated: new Date("2026-10-08T12:00:00Z"),
      internetMessageId: "<xyz@klant.nl>",
      conversationId: "conv1",
      attachments: [
        { name: "offerte.pdf", size: 1234, contentType: "application/pdf", isInline: false },
        { name: "logo.png", size: 10, isInline: true }
      ]
    },
    user: { displayName: "Tim Falken", emailAddress: "TFalken@kvt.nl" },
    customerHint: "  Klant BV\n 10023 ",
    now: new Date("2026-10-09T07:00:00Z"),
    appVersion: "1.0.0"
  });
  assert.equal(meta.schema, "arke.mail-metadata/v1");
  assert.equal(meta.messageKey, Core.messageKey("<xyz@klant.nl>"));
  assert.equal(meta.internetMessageId, "<xyz@klant.nl>");
  assert.deepEqual(meta.from, { name: "Jan Klant", email: "jan@klant.nl" });
  assert.deepEqual(meta.to, [{ name: "Tim", email: "tfalken@kvt.nl" }]);
  assert.deepEqual(meta.savedBy, { name: "Tim Falken", email: "tfalken@kvt.nl" });
  assert.equal(meta.receivedAt, "2026-10-08T12:00:00.000Z");
  assert.equal(meta.savedAt, "2026-10-09T07:00:00.000Z");
  assert.equal(meta.customerHint, "Klant BV 10023");
  assert.equal(meta.hasAttachments, true);
  assert.equal(meta.attachments.length, 1);
  assert.equal(meta.files.eml, `2026-10-08_1200_Offerte-pomp_${meta.messageKey}.eml`);
  assert.equal(meta.files.metadata, meta.files.eml.replace(/\.eml$/, ".json"));
  assert.equal(meta.processing.status, "nieuw");
});

test("buildMetadata: lege klanthint wordt null", () => {
  const meta = Core.buildMetadata({ item: { internetMessageId: "<a@b>", dateTimeCreated: "2026-01-01T00:00:00Z" }, customerHint: "   " });
  assert.equal(meta.customerHint, null);
  assert.equal(meta.from, null);
  assert.deepEqual(meta.to, []);
});

test("chooseUploadStrategy: grens 4 MB", () => {
  assert.equal(Core.chooseUploadStrategy(0), "simple");
  assert.equal(Core.chooseUploadStrategy(4 * 1024 * 1024), "simple");
  assert.equal(Core.chooseUploadStrategy(4 * 1024 * 1024 + 1), "session");
});

test("chunkRanges: aaneengesloten, veelvoud van 320 KiB, juiste headers", () => {
  const total = 10 * 1024 * 1024 + 7;
  const ranges = Core.chunkRanges(total);
  assert.equal(ranges[0].start, 0);
  assert.equal(ranges[ranges.length - 1].end, total - 1);
  for (let i = 0; i < ranges.length; i++) {
    const r = ranges[i];
    if (i < ranges.length - 1) {
      assert.equal((r.end - r.start + 1) % (320 * 1024), 0);
      assert.equal(ranges[i + 1].start, r.end + 1);
    }
    assert.equal(r.header, `bytes ${r.start}-${r.end}/${total}`);
  }
  assert.throws(() => Core.chunkRanges(100, 1000), /320 KiB/);
});

test("dedupeDecision", () => {
  assert.deepEqual(Core.dedupeDecision({ registerExists: true, inboxExists: false }), { duplicate: true, reason: "register" });
  assert.deepEqual(Core.dedupeDecision({ registerExists: false, inboxExists: true }), { duplicate: true, reason: "inbox" });
  assert.deepEqual(Core.dedupeDecision({ registerExists: false, inboxExists: false }), { duplicate: false, reason: null });
  assert.deepEqual(Core.dedupeDecision(null), { duplicate: false, reason: null });
});

test("drivePath encodeert per segment", () => {
  assert.equal(Core.drivePath("Klantmail/_Inbox", "a b#c.eml"), "Klantmail/_Inbox/a%20b%23c.eml");
  assert.equal(Core.drivePath("/_Inbox/"), "_Inbox");
});

test("base64ToBytes", () => {
  assert.deepEqual(Array.from(Core.base64ToBytes("SGFs\nbG8=")), Array.from(Buffer.from("Hallo")));
});

test("validateConfig", () => {
  assert.ok(Core.validateConfig(undefined).length > 0);
  const ok = { clientId: "00000000-0000-0000-0000-000000000000", tenantId: "kvt.nl", driveId: "b!x", inboxFolder: "_Inbox", scopes: ["x"] };
  assert.deepEqual(Core.validateConfig(ok), []);
  assert.ok(Core.validateConfig({ ...ok, clientId: "nee" }).some((e) => /clientId/.test(e)));
});

test("friendlyError geeft Nederlandse meldingen", () => {
  assert.match(Core.friendlyError({ status: 403 }), /schrijfrechten/);
  assert.match(Core.friendlyError({ status: 409 }), /al in SharePoint/);
  assert.match(Core.friendlyError({ message: "AADSTS65001: consent required" }), /admin consent/);
  assert.match(Core.friendlyError(new Error("boem")), /boem/);
});
