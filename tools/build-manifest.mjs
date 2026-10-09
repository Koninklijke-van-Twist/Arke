#!/usr/bin/env node
// Vult de manifest-templates in met de hosting-URL.
// Gebruik: node tools/build-manifest.mjs --host https://sleutels.kvt.nl/arke
// Output:  dist/manifest.xml  (voor het M365-beheercentrum)
//          dist/unified/manifest.json + iconen (optioneel; zip de map-inhoud voor de unified manifest)
import { readFileSync, writeFileSync, mkdirSync, copyFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const i = args.indexOf("--host");
const host = (i >= 0 ? args[i + 1] : "").replace(/\/+$/, "");
if (!/^https:\/\/[^/]+(\/.*)?$/.test(host)) {
  console.error("Gebruik: node tools/build-manifest.mjs --host https://sleutels.kvt.nl/arke");
  process.exit(1);
}
const domain = new URL(host).host;

export function fill(text) {
  return text.replaceAll("{{HOST_URL}}", host).replaceAll("{{HOST_ORIGIN}}", new URL(host).origin).replaceAll("{{HOST_DOMAIN}}", domain);
}

mkdirSync(join(root, "dist/unified"), { recursive: true });
const xml = fill(readFileSync(join(root, "manifest/manifest.template.xml"), "utf8"));
writeFileSync(join(root, "dist/manifest.xml"), xml);
const json = fill(readFileSync(join(root, "manifest/manifest.template.json"), "utf8"));
JSON.parse(json); // valideert dat het nog geldige JSON is
writeFileSync(join(root, "dist/unified/manifest.json"), json);
copyFileSync(join(root, "web/assets/icon-192.png"), join(root, "dist/unified/icon-192.png"));
copyFileSync(join(root, "web/assets/outline-32.png"), join(root, "dist/unified/outline-32.png"));
if (/\{\{[A-Z_]+\}\}/.test(xml + json)) {
  console.error("Let op: er staan nog niet-ingevulde placeholders in de manifest.");
  process.exit(2);
}
console.log("Klaar: dist/manifest.xml en dist/unified/ (host " + host + ")");
