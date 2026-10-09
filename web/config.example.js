/*
 * Arke - configuratie. Kopieer naar config.js (staat NIET in git) en vul in.
 * config.js wordt eenmalig handmatig op de server gezet; de FTP-deploy laat hem staan.
 * Er staan hier geen geheimen in (een SPA heeft geen client secret), maar we houden
 * tenant-specifieke ID's buiten de repo.
 */
window.ARKE_CONFIG = {
  // Entra ID > App-registraties > Arke > "Toepassings-id (client)"
  clientId: "00000000-0000-0000-0000-000000000000",
  // Tenant-ID (GUID) of het primaire domein, bv. "kvt.nl"
  tenantId: "kvt.nl",
  // Gedelegeerde Graph-scopes. Zie README (Sites.Selected = minimaal, Files.ReadWrite.All = eenvoudig alternatief)
  scopes: ["https://graph.microsoft.com/Sites.Selected"],
  // Ter documentatie; de add-in gebruikt alleen driveId.
  siteId: "kvtnl.sharepoint.com,xxxxxxxx-....,yyyyyyyy-....",
  // drive-id van de documentbibliotheek (begint meestal met "b!")
  driveId: "b!xxxxxxxxxxxxxxxxxxxxxxxx",
  // Map (relatief t.o.v. de bibliotheek-root) waar nieuwe mails landen
  inboxFolder: "_Inbox",
  // Map voor het dedupe-register (blijft staan als deel 2 mails verplaatst). Leeg = alleen _Inbox controleren.
  registerFolder: "_Register"
};
