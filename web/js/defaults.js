/*
 * Arke - standaardconfiguratie (KVT-tenant). Geen geheimen: een SPA heeft geen client secret,
 * en client-/tenant-id zijn publiek zichtbaar in elke aanmeld-URL.
 * Overschrijven kan per server met web/config.js (niet in git, zie config.example.js).
 */
window.ARKE_DEFAULTS = {
  clientId: "d2a6ccbd-3981-4e5f-a07c-83a277f978f8",
  tenantId: "e7f5c109-53a0-45fc-8779-4c9d83a4572a",
  scopes: ["https://graph.microsoft.com/Sites.Selected"],
  // Site + documentbibliotheek. Arke zoekt site-id en drive-id zelf op via Graph:
  // eerst site /sites/BCDocumentRepository met bibliotheek "Customer",
  // anders subsite /sites/BCDocumentRepository/Customer (standaardbibliotheek).
  siteUrl: "https://kvtnl.sharepoint.com/sites/BCDocumentRepository/Customer",
  libraryName: "",
  // Optionele overrides (sla de Graph-resolutie over):
  siteId: "",
  driveId: "",
  inboxFolder: "_Inbox",
  registerFolder: "_Register"
};
