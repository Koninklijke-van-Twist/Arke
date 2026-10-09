/*
 * Arke - pure helpers (geen Office.js / netwerk).
 * Werkt in de browser (window.ArkeCore) en in Node (module.exports) zodat
 * de logica met `node --test` getest kan worden.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.ArkeCore = api;
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var METADATA_SCHEMA = "arke.mail-metadata/v1";
  // Graph: "simple upload" tot 4 MB, daarboven een upload session.
  var SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;
  // Chunks moeten een veelvoud van 320 KiB zijn (Graph-eis). 10 x 320 KiB = 3,125 MiB.
  var CHUNK_SIZE = 10 * 320 * 1024;

  var ENTITY_ATTACHMENT_GROUP = "016_CORRESPONDENCE";

  var RESERVED_NAMES = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i;

  /** Maakt een tekst veilig als (deel van) een SharePoint/OneDrive-bestandsnaam. */
  function sanitizeFileNamePart(input, maxLength) {
    var max = typeof maxLength === "number" ? maxLength : 60;
    var s = String(input == null ? "" : input);
    if (typeof s.normalize === "function") {
      s = s.normalize("NFKD").replace(/[\u0300-\u036f]/g, "");
    }
    s = s
      .replace(/^\s*((re|fw|fwd|antw|doorst|tr|wg|aw)\s*:\s*)+/i, "") // reply/forward-prefixen
      .replace(/["*:<>?\/\\|#%&{}~]/g, " ") // tekens die SharePoint niet toestaat of lastig vindt
      .replace(/[\u0000-\u001f\u007f]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/ /g, "-")
      .replace(/-+/g, "-")
      .replace(/^[.\-_]+|[.\-_]+$/g, "");
    if (s.length > max) {
      s = s.slice(0, max).replace(/[.\-_]+$/g, "");
    }
    if (!s || RESERVED_NAMES.test(s)) {
      s = "zonder-onderwerp";
    }
    return s;
  }

  /** Normaliseert een internetMessageId (zonder <>, getrimd, lowercase). */
  function normalizeMessageId(id) {
    return String(id == null ? "" : id)
      .trim()
      .replace(/^<+|>+$/g, "")
      .trim()
      .toLowerCase();
  }

  /**
   * Synchrone, deterministische hash (FNV-1a 64-bit, hex) van een internetMessageId.
   * Geen cryptografische eis: alleen een stabiele, korte sleutel voor bestandsnamen/dedupe.
   */
  function messageKey(internetMessageId) {
    var norm = normalizeMessageId(internetMessageId);
    if (!norm) {
      throw new Error("internetMessageId ontbreekt");
    }
    // FNV-1a 64-bit met twee 32-bit helften (BigInt-vrij voor oude webviews).
    var h1 = 0x84222325 >>> 0; // laag
    var h2 = 0xcbf29ce4 >>> 0; // hoog
    var bytes = utf8Bytes(norm);
    for (var i = 0; i < bytes.length; i++) {
      h1 = (h1 ^ bytes[i]) >>> 0;
      // vermenigvuldigen met FNV-prime 0x100000001b3 = (2^40 + 0x1b3)
      var lo = h1 * 0x1b3;
      var carry = Math.floor(lo / 0x100000000);
      var newLo = lo >>> 0;
      var newHi = (h2 * 0x1b3 + carry + ((h1 << 8) >>> 0)) >>> 0;
      h1 = newLo;
      h2 = newHi;
    }
    return pad8(h2.toString(16)) + pad8(h1.toString(16));
  }

  function pad8(s) {
    while (s.length < 8) s = "0" + s;
    return s;
  }

  function utf8Bytes(str) {
    if (typeof TextEncoder !== "undefined") {
      return new TextEncoder().encode(str);
    }
    var out = [];
    var enc = unescape(encodeURIComponent(str));
    for (var i = 0; i < enc.length; i++) out.push(enc.charCodeAt(i));
    return out;
  }

  function two(n) {
    return (n < 10 ? "0" : "") + n;
  }

  /** "2026-10-09_0932" in UTC (stabiel, ongeacht de tijdzone van de pc). */
  function formatStamp(date) {
    var d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) {
      throw new Error("Ongeldige datum");
    }
    return (
      d.getUTCFullYear() + "-" + two(d.getUTCMonth() + 1) + "-" + two(d.getUTCDate()) +
      "_" + two(d.getUTCHours()) + two(d.getUTCMinutes())
    );
  }

  /**
   * Basisnaam (zonder extensie) voor .eml en .json.
   * Vorm: JJJJ-MM-DD_HHMM_<onderwerp>_<key16>  (tijd in UTC)
   * Deterministisch per mail, dus dezelfde mail levert altijd dezelfde naam op.
   */
  function buildBaseName(opts) {
    var stamp = formatStamp(opts.date);
    var subject = sanitizeFileNamePart(opts.subject, 60);
    var key = opts.key || messageKey(opts.internetMessageId);
    return stamp + "_" + subject + "_" + key;
  }

  function normalizeAddress(a) {
    if (!a) return null;
    var email = String(a.emailAddress || a.address || a.email || "").trim();
    var name = String(a.displayName || a.name || "").trim();
    if (!email && !name) return null;
    return { name: name || email, email: email.toLowerCase() };
  }

  function normalizeList(list) {
    var out = [];
    (list || []).forEach(function (a) {
      var n = normalizeAddress(a);
      if (n) out.push(n);
    });
    return out;
  }

  function toIso(d) {
    if (!d) return null;
    var date = d instanceof Date ? d : new Date(d);
    return isNaN(date.getTime()) ? null : date.toISOString();
  }

  function cleanHint(s) {
    var t = String(s == null ? "" : s).replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim();
    return t.length > 120 ? t.slice(0, 120) : t;
  }

  /**
   * Bouwt het JSON-sidecar-object (schema arke.mail-metadata/v1).
   * item: { subject, from, to[], cc[], dateTimeCreated, internetMessageId, conversationId,
   *         itemId, attachments[] }  (namen zoals Office.js ze levert)
   * user: { displayName, emailAddress }
   */
  function buildMetadata(input) {
    var item = input.item || {};
    var user = input.user || {};
    var key = messageKey(item.internetMessageId);
    var baseName = input.baseName || buildBaseName({
      date: item.dateTimeCreated,
      subject: item.subject,
      key: key
    });
    var attachments = (item.attachments || [])
      .filter(function (a) { return a && !a.isInline; })
      .map(function (a) {
        return { name: String(a.name || ""), size: typeof a.size === "number" ? a.size : null, contentType: a.contentType || null };
      });
    return {
      schema: METADATA_SCHEMA,
      messageKey: key,
      internetMessageId: String(item.internetMessageId || "").trim(),
      subject: String(item.subject || ""),
      from: normalizeAddress(item.from),
      to: normalizeList(item.to),
      cc: normalizeList(item.cc),
      receivedAt: toIso(item.dateTimeCreated),
      conversationId: item.conversationId || null,
      hasAttachments: attachments.length > 0,
      attachments: attachments,
      savedBy: normalizeAddress(user),
      savedAt: toIso(input.now || new Date()),
      customerHint: cleanHint(input.customerHint) || null,
      // Vaste waarde + lege velden voor deel 2 (Power Automate/Copilot vult ze in).
      entityAttachmentGroup: ENTITY_ATTACHMENT_GROUP,
      customerNo: cleanHint(input.customerHint),
      kvtCustomerName: "",
      kvtSalesQuoteNo: "",
      kvtSalesQuoteDescription: "",
      files: {
        eml: baseName + ".eml",
        metadata: baseName + ".json"
      },
      source: {
        app: "Arke",
        version: input.appVersion || null,
        host: input.host || null
      },
      processing: {
        status: "nieuw"
      }
    };
  }

  /** Kiest de uploadmethode op basis van de grootte in bytes. */
  function chooseUploadStrategy(size) {
    return size <= SIMPLE_UPLOAD_LIMIT ? "simple" : "session";
  }

  /** Byte-ranges voor een upload session: [{start, end, header}] (end inclusief). */
  function chunkRanges(total, chunkSize) {
    var size = chunkSize || CHUNK_SIZE;
    if (size % (320 * 1024) !== 0) {
      throw new Error("Chunkgrootte moet een veelvoud van 320 KiB zijn");
    }
    var ranges = [];
    for (var start = 0; start < total; start += size) {
      var end = Math.min(start + size, total) - 1;
      ranges.push({ start: start, end: end, header: "bytes " + start + "-" + end + "/" + total });
    }
    return ranges;
  }

  /**
   * Dedupe-beslissing. checks: { registerExists, inboxExists } (true/false/null=onbekend).
   * Geeft { duplicate: bool, reason: string|null }.
   */
  function dedupeDecision(checks) {
    if (checks && checks.registerExists) return { duplicate: true, reason: "register" };
    if (checks && checks.inboxExists) return { duplicate: true, reason: "inbox" };
    return { duplicate: false, reason: null };
  }

  /** Graph-pad (relatief t.o.v. de drive-root) met URL-encoding per segment. */
  function drivePath(folder, fileName) {
    var parts = String(folder || "").split("/").filter(Boolean);
    if (fileName) parts.push(fileName);
    return parts.map(encodeURIComponent).join("/");
  }

  /** Decodeert base64 naar Uint8Array (browser en Node). */
  function base64ToBytes(b64) {
    var clean = String(b64 || "").replace(/\s+/g, "");
    if (typeof atob === "function") {
      var bin = atob(clean);
      var out = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    return new Uint8Array(Buffer.from(clean, "base64"));
  }

  /**
   * Splitst een SharePoint-URL: https://kvtnl.sharepoint.com/sites/A/B
   * -> { host: "kvtnl.sharepoint.com", segments: ["sites","A","B"] } of null.
   */
  function parseSiteUrl(url) {
    var m = /^https:\/\/([a-z0-9.-]+\.sharepoint\.com)(\/[^?#]*)?(?:[?#].*)?$/i.exec(String(url || "").trim());
    if (!m) return null;
    var segments = (m[2] || "").split("/").filter(Boolean).map(function (x) {
      try { return decodeURIComponent(x); } catch (e) { return x; }
    });
    // Veelvoorkomende pagina-achtervoegsels negeren (bv. .../Forms/AllItems.aspx).
    var formsIdx = segments.indexOf("Forms");
    if (formsIdx > 0) segments = segments.slice(0, formsIdx);
    if (segments.length && /\.aspx$/i.test(segments[segments.length - 1])) segments.pop();
    if (segments.length < 2 || !/^(sites|teams)$/i.test(segments[0])) return null;
    return { host: m[1].toLowerCase(), segments: segments };
  }

  /**
   * Kandidaten voor (site, bibliotheek) in volgorde van proberen.
   * Voor .../sites/A/B:
   *   1. site /sites/A met bibliotheek "B"      (B = documentbibliotheek)
   *   2. site /sites/A/B met bibliotheek libraryName of de standaardbibliotheek (B = subsite)
   * Elk: { sitePath: "kvtnl.sharepoint.com:/sites/A", library: "B" | null }
   */
  function siteCandidates(siteUrl, libraryName) {
    var p = parseSiteUrl(siteUrl);
    if (!p) return [];
    var out = [];
    var segs = p.segments;
    function path(list) {
      return p.host + ":/" + list.map(encodeURIComponent).join("/");
    }
    if (segs.length >= 3) {
      out.push({ sitePath: path(segs.slice(0, -1)), library: libraryName || segs[segs.length - 1], kind: "library" });
    }
    out.push({ sitePath: path(segs), library: libraryName || null, kind: segs.length >= 3 ? "subsite" : "site" });
    return out;
  }

  /** Kiest de drive die bij een bibliotheeknaam hoort (op naam of op het laatste stuk van webUrl). */
  function pickDrive(drives, library) {
    var list = drives || [];
    if (!library) return null;
    var want = String(library).toLowerCase();
    var byName = list.filter(function (d) { return String(d.name || "").toLowerCase() === want; })[0];
    if (byName) return byName;
    return list.filter(function (d) {
      var url = String(d.webUrl || "");
      var last = url.split("/").filter(Boolean).pop() || "";
      try { last = decodeURIComponent(last); } catch (e) { /* laat staan */ }
      return last.toLowerCase() === want;
    })[0] || null;
  }

  /** Cachesleutel voor de opgeloste site/drive (wisselt mee met URL en bibliotheek). */
  function targetCacheKey(cfg) {
    return "arke.target.v1|" + String(cfg.siteUrl || "").toLowerCase().replace(/\/+$/, "") + "|" + String(cfg.libraryName || "").toLowerCase();
  }

  /** Mappen (incl. tussenliggende) die moeten bestaan: ["_Inbox", "_Register"] of ["A", "A/_Inbox"]. */
  function foldersToEnsure(cfg) {
    var out = [];
    [folderValue(cfg.inboxFolder), folderValue(cfg.registerFolder)].forEach(function (f) {
      var parts = f.split("/").filter(Boolean);
      for (var i = 1; i <= parts.length; i++) {
        var p = parts.slice(0, i).join("/");
        if (out.indexOf(p) < 0) out.push(p);
      }
    });
    return out;
  }

  // Velden waarbij een lege waarde ("", false, null) in de override bewust iets UITzet.
  var CLEARABLE_KEYS = ["registerFolder"];

  /**
   * Combineert defaults met een (optionele) override. Lege waarden in de override tellen als
   * "niet ingesteld", behalve voor CLEARABLE_KEYS: daar zet "", false of null de functie uit ("").
   */
  function mergeConfig(defaults, override) {
    var out = {};
    var k;
    for (k in (defaults || {})) out[k] = defaults[k];
    for (k in (override || {})) {
      if (!Object.prototype.hasOwnProperty.call(override, k)) continue;
      var v = override[k];
      var empty = v === undefined || v === null || v === "" || v === false;
      if (empty) {
        if (CLEARABLE_KEYS.indexOf(k) >= 0 && v !== undefined) out[k] = "";
        continue;
      }
      out[k] = v;
    }
    return out;
  }

  /** Mapnaam als string; false/null/"" -> "" (uit). */
  function folderValue(f) {
    return typeof f === "string" ? f.trim() : "";
  }

  /** Valideert de config; geeft een lijst met Nederlandse foutmeldingen terug. */
  function validateConfig(cfg) {
    var errors = [];
    var guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!cfg) return ["config.js ontbreekt of is leeg."];
    if (!cfg.clientId || !guid.test(cfg.clientId)) errors.push("clientId ontbreekt of is geen GUID.");
    if (!cfg.tenantId) errors.push("tenantId ontbreekt (GUID of bv. kvt.nl).");
    if (!cfg.driveId && !cfg.siteUrl) errors.push("siteUrl (of een driveId-override) ontbreekt.");
    if (cfg.siteUrl && !parseSiteUrl(cfg.siteUrl)) errors.push("siteUrl is geen geldige SharePoint-URL (https://<tenant>.sharepoint.com/sites/...).");
    if (!cfg.inboxFolder) errors.push("inboxFolder ontbreekt (bv. _Inbox).");
    if (!Array.isArray(cfg.scopes) || cfg.scopes.length === 0) errors.push("scopes ontbreekt.");
    return errors;
  }

  /** Vertaalt een fout naar een begrijpelijke Nederlandse melding. */
  function friendlyError(err) {
    var status = err && err.status;
    var code = (err && (err.code || err.errorCode)) || "";
    if (status === 401) return "Je sessie is verlopen. Probeer het opnieuw.";
    if (status === 403) return "Je hebt geen schrijfrechten op de SharePoint-map. Vraag ICT om toegang.";
    if (status === 404) return "De SharePoint-site of -bibliotheek is niet gevonden. Controleer siteUrl/libraryName in de configuratie.";
    if (status === 409) return "Deze mail staat al in SharePoint.";
    if (status === 413) return "De mail is te groot om op te slaan.";
    if (status === 429 || status === 503) return "SharePoint is even druk. Probeer het over een minuut opnieuw.";
    if (/consent|AADSTS65001/i.test(code + " " + (err && err.message))) {
      return "De add-in heeft nog geen toestemming (admin consent). Neem contact op met ICT.";
    }
    if (/user_cancelled|popup_window_error|12006/i.test(code + " " + (err && err.message))) {
      return "Inloggen is afgebroken. Probeer het opnieuw.";
    }
    return "Opslaan is mislukt" + (err && err.message ? ": " + err.message : ".");
  }

  return {
    METADATA_SCHEMA: METADATA_SCHEMA,
    SIMPLE_UPLOAD_LIMIT: SIMPLE_UPLOAD_LIMIT,
    CHUNK_SIZE: CHUNK_SIZE,
    sanitizeFileNamePart: sanitizeFileNamePart,
    normalizeMessageId: normalizeMessageId,
    messageKey: messageKey,
    formatStamp: formatStamp,
    buildBaseName: buildBaseName,
    buildMetadata: buildMetadata,
    chooseUploadStrategy: chooseUploadStrategy,
    chunkRanges: chunkRanges,
    dedupeDecision: dedupeDecision,
    drivePath: drivePath,
    base64ToBytes: base64ToBytes,
    validateConfig: validateConfig,
    mergeConfig: mergeConfig,
    parseSiteUrl: parseSiteUrl,
    siteCandidates: siteCandidates,
    pickDrive: pickDrive,
    targetCacheKey: targetCacheKey,
    foldersToEnsure: foldersToEnsure,
    ENTITY_ATTACHMENT_GROUP: ENTITY_ATTACHMENT_GROUP,
    friendlyError: friendlyError
  };
});
