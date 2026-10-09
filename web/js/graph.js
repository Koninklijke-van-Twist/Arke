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

  root.ArkeGraph = {
    exists: exists,
    checkFolder: checkFolder,
    upload: upload
  };
})(self);
