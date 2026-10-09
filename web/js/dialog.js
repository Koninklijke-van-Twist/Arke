/*
 * Arke - fallback-aanmelding in een Office-dialoog (als NAA niet beschikbaar is).
 * Standaard MSAL-redirectflow; het token gaat via messageParent terug naar de taskpane.
 * Redirect-URI (SPA) in Entra: https://<host>/<pad>/dialog.html
 */
(function () {
  "use strict";
  var cfg = window.ArkeCore.mergeConfig(window.ARKE_DEFAULTS, window.ARKE_CONFIG);

  function send(obj) {
    Office.context.ui.messageParent(JSON.stringify(obj));
  }

  Office.onReady(async function () {
    try {
      var redirectUri = window.location.origin + window.location.pathname;
      var pca = new msal.PublicClientApplication({
        auth: {
          clientId: cfg.clientId,
          authority: "https://login.microsoftonline.com/" + cfg.tenantId,
          redirectUri: redirectUri
        },
        cache: { cacheLocation: "localStorage" }
      });
      await pca.initialize();
      var result = await pca.handleRedirectPromise();
      if (result && result.accessToken) {
        send({ ok: true, accessToken: result.accessToken });
        return;
      }
      var hint = new URLSearchParams(window.location.search).get("hint") || undefined;
      var accounts = pca.getAllAccounts();
      if (accounts.length) {
        try {
          var silent = await pca.acquireTokenSilent({ scopes: cfg.scopes, account: accounts[0] });
          send({ ok: true, accessToken: silent.accessToken });
          return;
        } catch (e) { /* door naar redirect */ }
      }
      await pca.loginRedirect({ scopes: cfg.scopes, loginHint: hint, redirectUri: redirectUri });
    } catch (e) {
      send({ ok: false, error: e && e.message, code: e && (e.errorCode || e.code) });
    }
  });
})();
