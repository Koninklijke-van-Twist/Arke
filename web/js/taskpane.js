/*
 * Arke - taskpane: slaat de geopende mail op als .eml + JSON-sidecar in SharePoint.
 * Alleen handmatig: er gebeurt niets zonder klik op "Opslaan in SharePoint".
 */
(function () {
  "use strict";
  var APP_VERSION = "1.1.0";
  var Core = window.ArkeCore;
  var cfg = Core.mergeConfig(window.ARKE_DEFAULTS, window.ARKE_CONFIG);
  var busy = false;
  var foldersReady = {}; // per driveId: mappen deze sessie al gecontroleerd

  /** Zoekt de doel-drive op en zorgt dat de mappen bestaan. Geeft een runtime-config terug. */
  async function prepareTarget(token, refresh) {
    var target = await ArkeGraph.resolveTarget(token, cfg, { refresh: refresh });
    var runCfg = Object.assign({}, cfg, { siteId: target.siteId, driveId: target.driveId });
    if (!foldersReady[runCfg.driveId]) {
      try {
        await ArkeGraph.ensureFolders(token, runCfg);
      } catch (e) {
        // Gecachte drive bestaat niet meer: cache weggooien en één keer opnieuw zoeken.
        if (e.status === 404 && target.source === "cache") {
          ArkeGraph.forgetTarget(cfg);
          return prepareTarget(token, true);
        }
        throw e;
      }
      foldersReady[runCfg.driveId] = true;
    }
    return runCfg;
  }

  function $(id) { return document.getElementById(id); }

  function show(kind, html) {
    var el = $("message");
    el.className = "message " + kind;
    el.innerHTML = html;
    el.hidden = false;
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function progress(fraction) {
    var p = $("progress");
    if (fraction == null) { p.hidden = true; return; }
    p.hidden = false;
    $("progress-bar").style.width = Math.round(fraction * 100) + "%";
  }

  function currentItem() {
    return Office.context.mailbox && Office.context.mailbox.item;
  }

  function supportsEml() {
    var item = currentItem();
    return !!(item && typeof item.getAsFileAsync === "function" &&
      Office.context.requirements.isSetSupported("Mailbox", "1.14"));
  }

  function render() {
    var item = currentItem();
    $("message").hidden = true;
    progress(null);
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || !item.internetMessageId) {
      $("mail-card").hidden = true;
      $("save").disabled = true;
      show("info", "Open of selecteer een ontvangen of verzonden mail om die op te slaan.");
      return;
    }
    $("m-subject").textContent = item.subject || "(geen onderwerp)";
    $("m-from").textContent = item.from ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "-";
    $("m-date").textContent = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("nl-NL") : "-";
    var att = (item.attachments || []).filter(function (a) { return !a.isInline; });
    $("m-attachments").textContent = att.length ? att.length + " (worden in de .eml meegenomen)" : "geen";
    $("mail-card").hidden = false;
    $("customer-hint").value = "";
    if (!supportsEml()) {
      $("save").disabled = true;
      show("err", "Deze versie van Outlook ondersteunt het opslaan van mails nog niet (Mailbox 1.14 nodig). " +
        "Werk Outlook bij of gebruik Outlook op het web / de nieuwe Outlook.");
      return;
    }
    $("save").disabled = false;
  }

  function getEmlBase64(item) {
    return new Promise(function (resolve, reject) {
      item.getAsFileAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(new Error("Mail kon niet worden uitgelezen (" + (res.error && res.error.message) + ")"));
      });
    });
  }

  async function save() {
    if (busy) return;
    var item = currentItem();
    if (!item) return;
    busy = true;
    $("save").disabled = true;
    show("info", "Bezig met opslaan&hellip;");
    progress(0.05);
    try {
      var token = await ArkeAuth.getToken(cfg);
      $("auth-mode").textContent = "· aanmelding: " + (ArkeAuth.mode() === "naa" ? "SSO" : "dialoog");
      progress(0.1);
      var rc = await prepareTarget(token, false);
      progress(0.15);

      var profile = Office.context.mailbox.userProfile;
      var meta = Core.buildMetadata({
        item: {
          subject: item.subject,
          from: item.from,
          to: item.to,
          cc: item.cc,
          dateTimeCreated: item.dateTimeCreated,
          internetMessageId: item.internetMessageId,
          conversationId: item.conversationId,
          attachments: item.attachments
        },
        user: { displayName: profile.displayName, emailAddress: profile.emailAddress },
        customerHint: $("customer-hint").value,
        appVersion: APP_VERSION,
        host: Office.context.diagnostics ? (Office.context.diagnostics.platform + " " + Office.context.diagnostics.version) : null
      });

      // Dedupe op internetMessageId: register (blijft bestaan als deel 2 de mail verplaatst) + _Inbox.
      var registerName = meta.messageKey + ".json";
      var checks = {
        registerExists: rc.registerFolder ? await ArkeGraph.exists(token, rc, rc.registerFolder, registerName) : false,
        inboxExists: await ArkeGraph.exists(token, rc, rc.inboxFolder, meta.files.metadata)
      };
      var decision = Core.dedupeDecision(checks);
      if (decision.duplicate) {
        show("warn", "Deze mail is al eerder opgeslagen in SharePoint. Er is niets dubbel opgeslagen.");
        return;
      }
      progress(0.25);

      var eml = Core.base64ToBytes(await getEmlBase64(item));
      progress(0.35);

      // Volgorde: eerst .eml, dan de sidecar .json. Deel 2 start op de .json, dan is de .eml er gegarandeerd.
      var emlItem;
      try {
        emlItem = await ArkeGraph.upload(token, rc, rc.inboxFolder, meta.files.eml, eml, "message/rfc822",
          function (f) { progress(0.35 + f * 0.5); });
      } catch (e) {
        if (e.status === 409) {
          // .eml bestaat al (bv. eerdere poging brak af na de .eml): sidecar alsnog schrijven.
          emlItem = null;
        } else {
          throw e;
        }
      }
      meta.files.emlSize = eml.length;
      if (emlItem && emlItem.webUrl) meta.files.emlWebUrl = emlItem.webUrl;
      var metaBytes = new TextEncoder().encode(JSON.stringify(meta, null, 2));
      await ArkeGraph.upload(token, rc, rc.inboxFolder, meta.files.metadata, metaBytes, "application/json");
      progress(0.95);

      if (rc.registerFolder) {
        try {
          var reg = { schema: "arke.register/v1", messageKey: meta.messageKey, internetMessageId: meta.internetMessageId,
            savedBy: meta.savedBy, savedAt: meta.savedAt, files: meta.files, inboxFolder: rc.inboxFolder };
          await ArkeGraph.upload(token, rc, rc.registerFolder, registerName,
            new TextEncoder().encode(JSON.stringify(reg, null, 2)), "application/json");
        } catch (e) {
          if (e.status !== 409) console.warn("Register bijwerken mislukt", e);
        }
      }
      progress(1);
      var link = emlItem && emlItem.webUrl ? ' <a href="' + esc(emlItem.webUrl) + '" target="_blank" rel="noopener">Bekijk in SharePoint</a>' : "";
      show("ok", "&#10003; Opgeslagen in SharePoint (" + esc(cfg.inboxFolder) + ")." + link);
    } catch (e) {
      console.error(e);
      if (e && e.status === 409) show("warn", "Deze mail staat al in SharePoint.");
      else show("err", esc(Core.friendlyError(e)));
    } finally {
      busy = false;
      progress(null);
      $("save").disabled = !supportsEml();
    }
  }

  Office.onReady(function (info) {
    $("version").textContent = "Arke v" + APP_VERSION;
    var errors = Core.validateConfig(cfg);
    if (errors.length) {
      show("err", "De add-in is nog niet goed ingesteld. Neem contact op met ICT.<br><small>" +
        errors.map(esc).join("<br>") + "</small>");
      return;
    }
    if (info.host !== Office.HostType.Outlook) {
      show("err", "Deze add-in werkt alleen in Outlook.");
      return;
    }
    $("save").addEventListener("click", save);
    ArkeAuth.init(cfg).then(function (mode) {
      $("auth-mode").textContent = "· aanmelding: " + (mode === "naa" ? "SSO" : "dialoog");
    }).catch(function (e) { console.warn("Auth-init", e); });
    // Vastgezette taskpane: bij wisselen van mail opnieuw tonen.
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        if (!busy) render();
      });
    }
    render();
  });
})();
