"use strict";
// taskpane.js met een nep-DOM, nep-Office en nep-Graph: gedrag van een vastgezette taskpane.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function el() {
  return { hidden: false, disabled: false, textContent: "", innerHTML: "", className: "", value: "", style: {},
    listeners: {}, addEventListener(t, f) { this.listeners[t] = f; } };
}

function mail(id, subject) {
  return {
    itemType: "message",
    internetMessageId: id,
    subject,
    from: { displayName: "Klant", emailAddress: "k@klant.nl" },
    to: [], cc: [], attachments: [],
    dateTimeCreated: new Date("2026-10-09T10:00:00Z"),
    emlCalls: 0,
    getAsFileAsync(cb) { this.emlCalls++; cb({ status: "succeeded", value: Buffer.from("EML " + id).toString("base64") }); }
  };
}

function setup() {
  const els = {};
  const document = { getElementById: (id) => (els[id] = els[id] || el()) };
  let itemChangedHandler = null;
  const mailbox = {
    item: null,
    userProfile: { displayName: "Tim", emailAddress: "tfalken@kvt.nl" },
    addHandlerAsync: (type, h) => { itemChangedHandler = h; }
  };
  let ready;
  const Office = {
    HostType: { Outlook: "Outlook" },
    MailboxEnums: { ItemType: { Message: "message" } },
    AsyncResultStatus: { Succeeded: "succeeded" },
    EventType: { ItemChanged: "itemChanged" },
    context: { mailbox, requirements: { isSetSupported: () => true }, diagnostics: null },
    onReady: (f) => { ready = f; }
  };
  const uploads = [];
  let gate = null; // laat de eerste Graph-aanroep wachten
  const ctx = {
    console: { log() {}, warn() {}, error() {} }, TextEncoder, Buffer, atob: (s) => Buffer.from(s, "base64").toString("binary"),
    document, Office,
    ARKE_DEFAULTS: {
      clientId: "d2a6ccbd-3981-4e5f-a07c-83a277f978f8", tenantId: "t", scopes: ["s"],
      siteUrl: "https://kvtnl.sharepoint.com/sites/A/B", inboxFolder: "_Inbox", registerFolder: "_Register"
    },
    ArkeAuth: { init: async () => "naa", getToken: async () => "tok", mode: () => "naa" },
    ArkeGraph: {
      resolveTarget: async () => { if (gate) await gate.promise; return { siteId: "S", driveId: "D", source: "graph" }; },
      forgetTarget() {},
      ensureFolders: async () => {},
      exists: async () => false,
      upload: async (tok, cfg, folder, name, bytes) => { uploads.push({ folder, name, bytes: Buffer.from(bytes).toString() }); return { webUrl: "https://sp/" + name }; }
    }
  };
  ctx.window = ctx;
  ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../web/js/core.js"), "utf8"), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../web/js/taskpane.js"), "utf8"), ctx);
  return {
    els, mailbox, uploads,
    start(item) { mailbox.item = item; ready({ host: "Outlook" }); },
    switchTo(item) { mailbox.item = item; itemChangedHandler(); },
    hold() { let r; gate = { promise: new Promise((res) => (r = res)) }; return () => { gate = null; r(); }; },
    click: () => els.save.listeners.click()
  };
}

const tick = () => new Promise((r) => setImmediate(r));

test("normaal opslaan: .eml, .json en register voor de geopende mail", async () => {
  const t = setup();
  t.start(mail("<a@x>", "Mail A"));
  await t.click();
  assert.deepEqual(t.uploads.map((u) => u.folder), ["_Inbox", "_Inbox", "_Register"]);
  assert.equal(t.uploads[0].bytes, "EML <a@x>");
  assert.match(t.els.message.className, /ok/);
});

test("vastgezette pane: wissel tijdens opslaan vóór het ophalen -> niets opgeslagen, kaart toont nieuwe mail", async () => {
  const t = setup();
  const a = mail("<a@x>", "Mail A");
  const b = mail("<b@x>", "Mail B");
  t.start(a);
  assert.equal(t.els["m-subject"].textContent, "Mail A");
  const release = t.hold();
  const p = t.click();
  await tick();
  t.switchTo(b); // genegeerd voor render tijdens busy, maar onthouden
  assert.equal(t.els["m-subject"].textContent, "Mail A");
  release();
  await p;
  assert.equal(t.uploads.length, 0, "geen upload van de verkeerde (of halve) mail");
  assert.equal(a.emlCalls + b.emlCalls, 0);
  assert.equal(t.els["m-subject"].textContent, "Mail B", "na afloop opnieuw gerenderd");
  assert.equal(t.els.save.disabled, false);
  assert.match(t.els.message.innerHTML, /andere mail geopend/);
  // Volgende klik slaat B op.
  await t.click();
  assert.equal(t.uploads[0].bytes, "EML <b@x>");
  assert.match(t.uploads[1].name, /Mail-B/);
});

test("vastgezette pane: wissel na het ophalen -> A wordt opgeslagen, daarna B getoond", async () => {
  const t = setup();
  const a = mail("<a@x>", "Mail A");
  const b = mail("<b@x>", "Mail B");
  t.start(a);
  // Wissel zodra de eerste upload (A) start.
  const origUpload = t.uploads.push.bind(t.uploads);
  let switched = false;
  t.uploads.push = (u) => { if (!switched) { switched = true; t.switchTo(b); } return origUpload(u); };
  await t.click();
  assert.equal(t.uploads[0].bytes, "EML <a@x>");
  const meta = JSON.parse(t.uploads[1].bytes);
  assert.equal(meta.internetMessageId, "<a@x>");
  assert.equal(meta.subject, "Mail A");
  assert.equal(t.els["m-subject"].textContent, "Mail B");
  assert.match(t.els.message.className, /ok/);
  assert.match(t.els.message.innerHTML, /vorige mail/);
});
