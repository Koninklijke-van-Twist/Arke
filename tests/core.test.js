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
  assert.equal(meta.entityAttachmentGroup, "016_CORRESPONDENCE");
  assert.equal(meta.customerNo, "Klant BV 10023");
  assert.equal(meta.kvtCustomerName, "");
  assert.equal(meta.kvtSalesQuoteNo, "");
  assert.equal(meta.kvtSalesQuoteDescription, "");
});

test("buildMetadata: lege klanthint wordt null", () => {
  const meta = Core.buildMetadata({ item: { internetMessageId: "<a@b>", dateTimeCreated: "2026-01-01T00:00:00Z" }, customerHint: "   " });
  assert.equal(meta.customerHint, null);
  assert.equal(meta.customerNo, "");
  assert.equal(meta.entityAttachmentGroup, "016_CORRESPONDENCE");
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
  assert.deepEqual(Core.validateConfig({ ...ok, driveId: "", siteUrl: "https://kvtnl.sharepoint.com/sites/A/B" }), []);
  assert.ok(Core.validateConfig({ ...ok, driveId: "" }).some((e) => /siteUrl/.test(e)));
  assert.ok(Core.validateConfig({ ...ok, driveId: "", siteUrl: "https://example.com/x" }).some((e) => /siteUrl/.test(e)));
  assert.ok(Core.validateConfig({ ...ok, clientId: "nee" }).some((e) => /clientId/.test(e)));
});

test("friendlyError geeft Nederlandse meldingen", () => {
  assert.match(Core.friendlyError({ status: 403 }), /schrijfrechten/);
  assert.match(Core.friendlyError({ status: 409 }), /al in SharePoint/);
  assert.match(Core.friendlyError({ message: "AADSTS65001: consent required" }), /admin consent/);
  assert.match(Core.friendlyError(new Error("boem")), /boem/);
});

test("defaults.js bevat de KVT-tenant en is geldig", () => {
  const fs = require("node:fs");
  const vm = require("node:vm");
  const ctx = { window: {} };
  vm.runInNewContext(fs.readFileSync(require("node:path").join(__dirname, "../web/js/defaults.js"), "utf8"), ctx);
  const d = ctx.window.ARKE_DEFAULTS;
  assert.equal(d.clientId, "d2a6ccbd-3981-4e5f-a07c-83a277f978f8");
  assert.equal(d.tenantId, "e7f5c109-53a0-45fc-8779-4c9d83a4572a");
  assert.deepEqual(Core.validateConfig(Core.mergeConfig(d, undefined)), []);
});

test("mergeConfig: override wint, lege waarden niet", () => {
  const m = Core.mergeConfig({ a: 1, b: "x", c: ["s"] }, { a: 2, b: "", c: ["t"], d: null });
  assert.deepEqual(m, { a: 2, b: "x", c: ["t"] });
  assert.deepEqual(Core.mergeConfig({ a: 1 }, undefined), { a: 1 });
});

test("parseSiteUrl", () => {
  assert.deepEqual(Core.parseSiteUrl("https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Customer"),
    { host: "kvtnl.sharepoint.com", segments: ["sites", "BCDocumentRepository", "Customer"] });
  assert.deepEqual(Core.parseSiteUrl("https://KVTNL.sharepoint.com/sites/A/Customer/Forms/AllItems.aspx?x=1").segments, ["sites", "A", "Customer"]);
  assert.deepEqual(Core.parseSiteUrl("https://kvtnl.sharepoint.com/sites/Mijn%20Site/").segments, ["sites", "Mijn Site"]);
  assert.equal(Core.parseSiteUrl("https://kvtnl.sharepoint.com/"), null);
  assert.equal(Core.parseSiteUrl("https://evil.example.com/sites/A"), null);
  assert.equal(Core.parseSiteUrl("http://kvtnl.sharepoint.com/sites/A"), null);
});

test("siteCandidates: eerst bibliotheek, dan subsite", () => {
  const c = Core.siteCandidates("https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Customer");
  assert.deepEqual(c.map((x) => [x.sitePath, x.library, x.kind]), [
    ["kvtnl.sharepoint.com:/sites/BCDocumentRepository", "Customer", "library"],
    ["kvtnl.sharepoint.com:/sites/BCDocumentRepository/Customer", null, "subsite"]
  ]);
  const withLib = Core.siteCandidates("https://kvtnl.sharepoint.com/sites/A/B", "Klantmail");
  assert.equal(withLib[0].library, "Klantmail");
  assert.equal(withLib[1].library, "Klantmail");
  const top = Core.siteCandidates("https://kvtnl.sharepoint.com/sites/A");
  assert.deepEqual(top.map((x) => [x.sitePath, x.library, x.kind]), [["kvtnl.sharepoint.com:/sites/A", null, "site"]]);
});

test("pickDrive: op naam of op webUrl", () => {
  const drives = [
    { id: "1", name: "Documenten", webUrl: "https://kvtnl.sharepoint.com/sites/A/Gedeelde%20documenten" },
    { id: "2", name: "Klanten", webUrl: "https://kvtnl.sharepoint.com/sites/A/Customer" },
    { id: "3", name: "customer", webUrl: "https://kvtnl.sharepoint.com/sites/A/Other" }
  ];
  assert.equal(Core.pickDrive(drives, "Customer").id, "3"); // naam gaat voor
  assert.equal(Core.pickDrive(drives.slice(0, 2), "Customer").id, "2"); // anders webUrl
  assert.equal(Core.pickDrive(drives, "Gedeelde documenten").id, "1");
  assert.equal(Core.pickDrive(drives, "Bestaat-niet"), null);
  assert.equal(Core.pickDrive(drives, null), null);
});

test("targetCacheKey wisselt mee met siteUrl en bibliotheek", () => {
  const a = Core.targetCacheKey({ siteUrl: "https://kvtnl.sharepoint.com/sites/A/B/" });
  assert.equal(a, Core.targetCacheKey({ siteUrl: "https://KVTNL.sharepoint.com/sites/A/B" }));
  assert.notEqual(a, Core.targetCacheKey({ siteUrl: "https://kvtnl.sharepoint.com/sites/A/B", libraryName: "X" }));
});

test("foldersToEnsure: tussenmappen eerst, geen dubbelen", () => {
  assert.deepEqual(Core.foldersToEnsure({ inboxFolder: "_Inbox", registerFolder: "_Register" }), ["_Inbox", "_Register"]);
  assert.deepEqual(Core.foldersToEnsure({ inboxFolder: "Klantmail/_Inbox", registerFolder: "Klantmail/_Register" }),
    ["Klantmail", "Klantmail/_Inbox", "Klantmail/_Register"]);
  assert.deepEqual(Core.foldersToEnsure({ inboxFolder: "_Inbox", registerFolder: "" }), ["_Inbox"]);
});

test("mergeConfig: registerFolder \"\"/false/null in override zet het register uit", () => {
  const d = { registerFolder: "_Register", inboxFolder: "_Inbox", driveId: "", libraryName: "" };
  for (const off of ["", false, null]) {
    const m = Core.mergeConfig(d, { registerFolder: off });
    assert.equal(m.registerFolder, "", `registerFolder=${JSON.stringify(off)}`);
    assert.deepEqual(Core.foldersToEnsure(m), ["_Inbox"]);
  }
  // Niet genoemd of undefined = default blijft.
  assert.equal(Core.mergeConfig(d, {}).registerFolder, "_Register");
  assert.equal(Core.mergeConfig(d, { registerFolder: undefined }).registerFolder, "_Register");
  // Andere lege velden blijven "niet ingesteld".
  const m = Core.mergeConfig({ ...d, driveId: "b!default", inboxFolder: "_Inbox" }, { driveId: "", inboxFolder: null, libraryName: false });
  assert.equal(m.driveId, "b!default");
  assert.equal(m.inboxFolder, "_Inbox");
  assert.equal(m.libraryName, "");
});

test("foldersToEnsure: false/null registerFolder maakt geen map 'false'/'null'", () => {
  assert.deepEqual(Core.foldersToEnsure({ inboxFolder: "_Inbox", registerFolder: false }), ["_Inbox"]);
  assert.deepEqual(Core.foldersToEnsure({ inboxFolder: "_Inbox", registerFolder: null }), ["_Inbox"]);
});
