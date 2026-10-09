/*
 * Arke - aanmelden via Nested App Authentication (NAA) met MSAL.js.
 * Fallback: als NAA niet beschikbaar is, een Office-dialoog (dialog.html) met de
 * standaard MSAL-redirectflow. Vereist window.msal (vendor/msal-browser.min.js).
 */
(function (root) {
  "use strict";
  var pca = null;
  var mode = null; // "naa" | "dialog"

  function naaSupported() {
    try {
      return !!(root.Office && Office.context && Office.context.requirements &&
        Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
    } catch (e) {
      return false;
    }
  }

  function authority(cfg) {
    return "https://login.microsoftonline.com/" + cfg.tenantId;
  }

  async function init(cfg) {
    if (naaSupported()) {
      pca = await msal.createNestablePublicClientApplication({
        auth: { clientId: cfg.clientId, authority: authority(cfg) },
        cache: { cacheLocation: "localStorage" }
      });
      mode = "naa";
    } else {
      mode = "dialog";
    }
    return mode;
  }

  function loginHint() {
    try { return Office.context.mailbox.userProfile.emailAddress; } catch (e) { return undefined; }
  }

  async function tokenViaNaa(cfg) {
    var req = { scopes: cfg.scopes, loginHint: loginHint() };
    try {
      var r = await pca.acquireTokenSilent(req);
      return r.accessToken;
    } catch (silentErr) {
      // Interactief: in NAA toont de Outlook-host zelf het (meestal onzichtbare) consent/login.
      var r2 = await pca.acquireTokenPopup(req);
      return r2.accessToken;
    }
  }

  function tokenViaDialog(cfg) {
    return new Promise(function (resolve, reject) {
      var base = new URL("dialog.html", root.location.href);
      base.searchParams.set("hint", loginHint() || "");
      Office.context.ui.displayDialogAsync(base.toString(), { height: 60, width: 30, promptBeforeOpen: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          var e = new Error("Aanmeldvenster kon niet openen (" + (res.error && res.error.message) + ")");
          e.code = String(res.error && res.error.code);
          reject(e);
          return;
        }
        var dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          dialog.close();
          var msg;
          try { msg = JSON.parse(arg.message); } catch (e) { msg = { ok: false, error: "Ongeldig antwoord" }; }
          if (msg.ok && msg.accessToken) resolve(msg.accessToken);
          else { var err = new Error(msg.error || "Aanmelden mislukt"); err.code = msg.code; reject(err); }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
          var err = new Error("Aanmeldvenster gesloten");
          err.code = String(arg.error);
          reject(err);
        });
      });
    });
  }

  async function getToken(cfg) {
    if (!mode) await init(cfg);
    if (mode === "naa") {
      try {
        return await tokenViaNaa(cfg);
      } catch (e) {
        // NAA beschikbaar maar mislukt (bv. oude build): eenmalig terugvallen op de dialoog.
        if (/consent|AADSTS65001/i.test(String(e && (e.errorCode || e.message)))) throw e;
        mode = "dialog";
      }
    }
    return tokenViaDialog(cfg);
  }

  root.ArkeAuth = {
    init: init,
    getToken: getToken,
    mode: function () { return mode; },
    naaSupported: naaSupported
  };
})(self);
