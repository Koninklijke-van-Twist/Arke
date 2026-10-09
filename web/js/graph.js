/*
 * Arke - Microsoft Graph-aanroepen (drive-items). Vereist ArkeCore.
 * Alle aanroepen zijn gedelegeerd: ze gebeuren met het token en de rechten van de gebruiker.
 */
(function (root) {
  "use strict";
  var Core = root.ArkeCore;
  var GRAPH = "https://graph.microsoft.com/v1.0";

  function graphError(res, body) {
    var err = new Error((body && body.error && body.error.message) || ("HTTP " + res.status));
    err.status = res.status;
    err.code = body && body.error && body.error.code;
    return err;
  }

  async function request(token, method, url, opts) {
    opts = opts || {};
    var headers = Object.assign({ Authorization: "Bearer " + token }, opts.headers || {});
    for (var attempt = 0; attempt < 3; attempt++) {
      var res = await fetch(url.indexOf("https://") === 0 ? url : GRAPH + url, {
        method: method,
        headers: headers,
        body: opts.body
      });
      if ((res.status === 429 || res.status === 503) && attempt < 2) {
        var wait = Math.min(parseInt(res.headers.get("Retry-After") || "2", 10), 10) * 1000;
        await new Promise(function (r) { setTimeout(r, wait); });
        continue;
      }
      return res;
    }
    return res;
  }

  async function json(res) {
    try { return await res.json(); } catch (e) { return null; }
  }

  function itemUrl(cfg, folder, fileName) {
    return "/drives/" + encodeURIComponent(cfg.driveId) + "/root:/" + Core.drivePath(folder, fileName);
  }

  /** true als het item bestaat, false bij 404; gooit bij andere fouten. */
  async function exists(token, cfg, folder, fileName) {
    var res = await request(token, "GET", itemUrl(cfg, folder, fileName) + "?$select=id");
    if (res.status === 404) return false;
    if (res.ok) return true;
    throw graphError(res, await json(res));
  }

  /** Controleert of de doelmap bereikbaar is (geeft nette 403/404-fouten). */
  async function checkFolder(token, cfg) {
    var res = await request(token, "GET", itemUrl(cfg, cfg.inboxFolder) + "?$select=id,name,webUrl,folder");
    if (!res.ok) throw graphError(res, await json(res));
    return json(res);
  }

  /** Kleine upload (<= 4 MB) met PUT; conflictBehavior=fail voorkomt overschrijven. */
  async function putSmall(token, cfg, folder, fileName, bytes, contentType) {
    var res = await request(
      token, "PUT",
      itemUrl(cfg, folder, fileName) + ":/content?@microsoft.graph.conflictBehavior=fail",
      { headers: { "Content-Type": contentType || "application/octet-stream" }, body: bytes }
    );
    if (!res.ok) throw graphError(res, await json(res));
    return json(res);
  }

  /** Grote upload via upload session, in chunks van een veelvoud van 320 KiB. */
  async function putLarge(token, cfg, folder, fileName, bytes, onProgress) {
    var res = await request(token, "POST", itemUrl(cfg, folder, fileName) + ":/createUploadSession", {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "fail", name: fileName } })
    });
    if (!res.ok) throw graphError(res, await json(res));
    var session = await json(res);
    var ranges = Core.chunkRanges(bytes.length);
    var last = null;
    try {
      for (var i = 0; i < ranges.length; i++) {
        var r = ranges[i];
        // Upload-URL is vooraf geautoriseerd: GEEN Authorization-header meesturen.
        var chunkRes = await fetch(session.uploadUrl, {
          method: "PUT",
          headers: { "Content-Range": r.header },
          body: bytes.subarray(r.start, r.end + 1)
        });
        if (!chunkRes.ok && chunkRes.status !== 202) throw graphError(chunkRes, await json(chunkRes));
        last = chunkRes;
        if (onProgress) onProgress((r.end + 1) / bytes.length);
      }
    } catch (e) {
      fetch(session.uploadUrl, { method: "DELETE" }).catch(function () {});
      throw e;
    }
    return json(last);
  }

  async function upload(token, cfg, folder, fileName, bytes, contentType, onProgress) {
    if (Core.chooseUploadStrategy(bytes.length) === "simple") {
      return putSmall(token, cfg, folder, fileName, bytes, contentType);
    }
    return putLarge(token, cfg, folder, fileName, bytes, onProgress);
  }

  // ---- Site/drive-resolutie (runtime, met het gedelegeerde token) ----

  function cacheGet(key) {
    try {
      var v = root.localStorage && root.localStorage.getItem(key);
      if (v) return JSON.parse(v);
    } catch (e) { /* localStorage kan geblokkeerd zijn */ }
    try {
      var rs = root.Office && Office.context && Office.context.roamingSettings;
      var r = rs && rs.get(key);
      if (r) return typeof r === "string" ? JSON.parse(r) : r;
    } catch (e) { /* geen roamingSettings */ }
    return null;
  }

  function cacheSet(key, value) {
    try { root.localStorage && root.localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* negeren */ }
    try {
      var rs = root.Office && Office.context && Office.context.roamingSettings;
      if (rs) { rs.set(key, value); rs.saveAsync(function () {}); }
    } catch (e) { /* negeren */ }
  }

  function cacheClear(key) {
    try { root.localStorage && root.localStorage.removeItem(key); } catch (e) { /* negeren */ }
    try {
      var rs = root.Office && Office.context && Office.context.roamingSettings;
      if (rs) { rs.remove(key); rs.saveAsync(function () {}); }
    } catch (e) { /* negeren */ }
  }

  async function getJson(token, url) {
    var res = await request(token, "GET", url);
    var body = await json(res);
    if (!res.ok) throw graphError(res, body);
    return body;
  }

  /**
   * Zoekt site-id en drive-id op bij cfg.siteUrl. Probeert eerst "site + bibliotheek"
   * (/sites/A met bibliotheek B) en daarna "subsite" (/sites/A/B, standaardbibliotheek).
   * Gooit een nette fout als niets past.
   */
  async function discoverTarget(token, cfg) {
    var candidates = Core.siteCandidates(cfg.siteUrl, cfg.libraryName);
    var lastErr = null;
    for (var i = 0; i < candidates.length; i++) {
      var c = candidates[i];
      try {
        var site = await getJson(token, "/sites/" + c.sitePath + "?$select=id,webUrl,displayName");
        var drive = null;
        if (c.library) {
          var drives = await getJson(token, "/sites/" + encodeURIComponent(site.id) + "/drives?$select=id,name,webUrl");
          drive = Core.pickDrive(drives && drives.value, c.library);
        } else {
          drive = await getJson(token, "/sites/" + encodeURIComponent(site.id) + "/drive?$select=id,name,webUrl");
        }
        if (drive && drive.id) {
          return { siteId: site.id, driveId: drive.id, driveName: drive.name, webUrl: drive.webUrl, kind: c.kind };
        }
      } catch (e) {
        if (e.status && e.status !== 404 && e.status !== 400) lastErr = e; // 403 bewaren voor een duidelijke melding
      }
    }
    if (lastErr) throw lastErr;
    var err = new Error("SharePoint-bibliotheek niet gevonden bij " + cfg.siteUrl);
    err.status = 404;
    throw err;
  }

  /**
   * Geeft { siteId, driveId } terug: uit config (override), uit de cache, of via Graph.
   * opts.refresh = true negeert de cache.
   */
  async function resolveTarget(token, cfg, opts) {
    if (cfg.driveId) return { siteId: cfg.siteId || null, driveId: cfg.driveId, source: "config" };
    var key = Core.targetCacheKey(cfg);
    if (!(opts && opts.refresh)) {
      var cached = cacheGet(key);
      if (cached && cached.driveId) { cached.source = "cache"; return cached; }
    }
    var found = await discoverTarget(token, cfg);
    cacheSet(key, { siteId: found.siteId, driveId: found.driveId, driveName: found.driveName, webUrl: found.webUrl, kind: found.kind });
    found.source = "graph";
    return found;
  }

  function forgetTarget(cfg) {
    cacheClear(Core.targetCacheKey(cfg));
  }

  /** Maakt _Inbox/_Register (en tussenmappen) aan als ze ontbreken; 409 = bestaat al. */
  async function ensureFolders(token, cfg) {
    var folders = Core.foldersToEnsure(cfg);
    for (var i = 0; i < folders.length; i++) {
      var parts = folders[i].split("/");
      var name = parts.pop();
      var parent = parts.join("/");
      var url = "/drives/" + encodeURIComponent(cfg.driveId) +
        (parent ? "/root:/" + Core.drivePath(parent) + ":/children" : "/root/children");
      var res = await request(token, "POST", url, {
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
      });
      if (res.ok || res.status === 409) continue;
      throw graphError(res, await json(res));
    }
  }

  root.ArkeGraph = {
    resolveTarget: resolveTarget,
    forgetTarget: forgetTarget,
    ensureFolders: ensureFolders,
    exists: exists,
    checkFolder: checkFolder,
    upload: upload
  };
})(self);
