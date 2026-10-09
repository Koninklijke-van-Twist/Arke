"use strict";
// Test web/js/graph.js tegen een nep-Graph (fetch-mock), zonder tenant.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadGraph(routes, storage) {
  const calls = [];
  const store = storage || new Map();
  const ctx = {
    console,
    URL,
    TextEncoder,
    setTimeout,
    localStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k)
    },
    fetch: async (url, opts) => {
      const method = (opts && opts.method) || "GET";
      const u = decodeURIComponent(url.replace("https://graph.microsoft.com/v1.0", "").split("?")[0]);
      calls.push({ method, url: u, body: opts && opts.body });
      for (const r of routes) {
        if (r.method === method && (r.url instanceof RegExp ? r.url.test(u) : r.url === u)) {
          const out = typeof r.reply === "function" ? r.reply(opts) : r.reply;
          return new Response(JSON.stringify(out.body || {}), { status: out.status || 200, headers: { "Content-Type": "application/json" } });
        }
      }
      return new Response(JSON.stringify({ error: { code: "itemNotFound", message: "nf" } }), { status: 404 });
    }
  };
  ctx.self = ctx;
  vm.createContext(ctx);
  for (const f of ["core.js", "graph.js"]) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, "../web/js", f), "utf8"), ctx);
  }
  return { Graph: ctx.ArkeGraph, calls, store };
}

const cfg = {
  siteUrl: "https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Customer",
  libraryName: "",
  inboxFolder: "_Inbox",
  registerFolder: "_Register"
};

test("resolveTarget: site + bibliotheek 'Customer', daarna uit cache", async () => {
  const routes = [
    { method: "GET", url: "/sites/kvtnl.sharepoint.com:/sites/BCDocumentRepository", reply: { body: { id: "SITE1" } } },
    { method: "GET", url: "/sites/SITE1/drives", reply: { body: { value: [
      { id: "D-docs", name: "Documenten", webUrl: "https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Shared%20Documents" },
      { id: "D-cust", name: "Customer", webUrl: "https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Customer" }
    ] } } }
  ];
  const { Graph, calls } = loadGraph(routes);
  const t = await Graph.resolveTarget("tok", cfg);
  assert.equal(t.driveId, "D-cust");
  assert.equal(t.siteId, "SITE1");
  assert.equal(t.kind, "library");
  assert.equal(t.source, "graph");
  const n = calls.length;
  const t2 = await Graph.resolveTarget("tok", cfg);
  assert.equal(t2.source, "cache");
  assert.equal(t2.driveId, "D-cust");
  assert.equal(calls.length, n, "geen extra Graph-calls bij cache-hit");
});

test("resolveTarget: valt terug op subsite 'Customer' met standaardbibliotheek", async () => {
  const routes = [
    { method: "GET", url: "/sites/kvtnl.sharepoint.com:/sites/BCDocumentRepository", reply: { body: { id: "SITE1" } } },
    { method: "GET", url: "/sites/SITE1/drives", reply: { body: { value: [{ id: "D-docs", name: "Documenten", webUrl: "https://x/Shared%20Documents" }] } } },
    { method: "GET", url: "/sites/kvtnl.sharepoint.com:/sites/BCDocumentRepository/Customer", reply: { body: { id: "SUB1" } } },
    { method: "GET", url: "/sites/SUB1/drive", reply: { body: { id: "D-sub", name: "Documenten" } } }
  ];
  const { Graph } = loadGraph(routes);
  const t = await Graph.resolveTarget("tok", cfg);
  assert.equal(t.driveId, "D-sub");
  assert.equal(t.kind, "subsite");
});

test("resolveTarget: override in config slaat Graph over", async () => {
  const { Graph, calls } = loadGraph([]);
  const t = await Graph.resolveTarget("tok", { ...cfg, driveId: "b!override" });
  assert.equal(t.driveId, "b!override");
  assert.equal(t.source, "config");
  assert.equal(calls.length, 0);
});

test("resolveTarget: 403 (geen Sites.Selected-grant) komt als 403 terug", async () => {
  const routes = [
    { method: "GET", url: /^\/sites\/kvtnl/, reply: { status: 403, body: { error: { code: "accessDenied", message: "Access denied" } } } }
  ];
  const { Graph } = loadGraph(routes);
  await assert.rejects(Graph.resolveTarget("tok", cfg), (e) => e.status === 403);
});

test("resolveTarget: niets gevonden geeft 404", async () => {
  const { Graph } = loadGraph([]);
  await assert.rejects(Graph.resolveTarget("tok", cfg), (e) => e.status === 404);
});

test("ensureFolders: maakt mappen aan, 409 (bestaat al) wordt genegeerd", async () => {
  const routes = [
    { method: "POST", url: "/drives/D1/root/children", reply: (opts) => {
      const b = JSON.parse(opts.body);
      assert.equal(b["@microsoft.graph.conflictBehavior"], "fail");
      assert.deepEqual(b.folder, {});
      return b.name === "_Inbox" ? { status: 409, body: { error: { code: "nameAlreadyExists" } } } : { status: 201, body: { id: "x" } };
    } }
  ];
  const { Graph, calls } = loadGraph(routes);
  await Graph.ensureFolders("tok", { ...cfg, driveId: "D1" });
  assert.deepEqual(calls.map((c) => JSON.parse(c.body).name), ["_Inbox", "_Register"]);
});

test("ensureFolders: geneste map gebruikt het pad van de ouder", async () => {
  const routes = [
    { method: "POST", url: "/drives/D1/root/children", reply: { status: 201 } },
    { method: "POST", url: "/drives/D1/root:/Klantmail:/children", reply: { status: 201 } }
  ];
  const { Graph, calls } = loadGraph(routes);
  await Graph.ensureFolders("tok", { ...cfg, driveId: "D1", inboxFolder: "Klantmail/_Inbox", registerFolder: "" });
  assert.deepEqual(calls.map((c) => c.url), ["/drives/D1/root/children", "/drives/D1/root:/Klantmail:/children"]);
});

test("ensureFolders: 403 wordt doorgegeven", async () => {
  const routes = [{ method: "POST", url: /children$/, reply: { status: 403, body: { error: { code: "accessDenied" } } } }];
  const { Graph } = loadGraph(routes);
  await assert.rejects(Graph.ensureFolders("tok", { ...cfg, driveId: "D1" }), (e) => e.status === 403);
});
