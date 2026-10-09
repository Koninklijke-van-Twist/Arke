"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const { readFileSync } = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");

test("build-manifest vult alle placeholders en levert geldige JSON", () => {
  execFileSync("node", [path.join(root, "tools/build-manifest.mjs"), "--host", "https://sleutels.kvt.nl/arke/"], { cwd: root });
  const xml = readFileSync(path.join(root, "dist/manifest.xml"), "utf8");
  const json = JSON.parse(readFileSync(path.join(root, "dist/unified/manifest.json"), "utf8"));
  assert.doesNotMatch(xml, /\{\{/);
  assert.match(xml, /https:\/\/sleutels\.kvt\.nl\/arke\/taskpane\.html/);
  assert.match(xml, /<Permissions>ReadItem<\/Permissions>/);
  assert.deepEqual(json.validDomains, ["sleutels.kvt.nl"]);
});

test("build-manifest weigert een niet-https host", () => {
  assert.throws(() => execFileSync("node", [path.join(root, "tools/build-manifest.mjs"), "--host", "http://x"], { cwd: root, stdio: "pipe" }));
});
