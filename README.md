# Arke – Outlook-knop "Opslaan in SharePoint"

**Arke** (Ἄρκη) is in de Griekse mythologie de tweelingzus van Iris en de boodschapper van de Titanen. Deze add-in brengt berichten (mails) naar het archief.

Arke is een Outlook-add-in die ICT centraal uitrolt naar alle medewerkers. Een knop **Opslaan in SharePoint** in het lint slaat de geopende mail op als `.eml` in een vaste map (`_Inbox`) op een SharePoint-site, samen met een JSON-bestand met metadata. Dit is **deel 1** van de Ponos-taak *Klantgegevens automatisch in Sharepoint zetten*. Deel 2 (verwerking met Power Automate/Copilot) en deel 3 (doorzoeken) staan los van deze repo.

- **Alleen handmatig.** Er wordt nooit automatisch iets opgeslagen (privacy/AVG). De medewerker kiest zelf de mail en klikt op Opslaan.
- **Rechten van de gebruiker.** Alles gebeurt gedelegeerd met het Microsoft-account van de medewerker (SSO via Nested App Authentication). Er is geen server, geen client secret en geen service-account.
- **Hosting.** Statische bestanden op `https://sleutels.kvt.nl/arke/` (pagina-root `web/`, FTP-deploy bij een push naar `master`).

---

## Hoe het werkt

1. De medewerker opent of selecteert een mail en klikt op **Opslaan in SharePoint** (groep *Klantmail* op het tabblad Start). Het taakvenster opent en toont het onderwerp, de afzender, de datum en het aantal bijlagen.
2. Optioneel vult de medewerker een **klant** in (naam of klantnummer). Dat is een hint voor deel 2.
3. Na een klik op **Opslaan in SharePoint**:
   1. haalt de add-in een Graph-token op via **NAA** (`msal.createNestablePublicClientApplication`). Is NAA niet beschikbaar, dan opent een klein aanmeldvenster (Office-dialoog met de standaard MSAL-redirectflow, `dialog.html`);
   2. **dedupe**: de add-in berekent een sleutel uit de `internetMessageId` en controleert of `_Register/<sleutel>.json` of de sidecar in `_Inbox` al bestaat. Zo ja, dan krijgt de medewerker de melding "al eerder opgeslagen" en wordt er niets geüpload;
   3. haalt de complete mail op als `.eml` via Office.js `item.getAsFileAsync()` (Mailbox 1.14). Daarvoor is geen Graph-mailrecht nodig;
   4. uploadt de `.eml` naar `_Inbox` (tot 4 MB met één `PUT`, groter via een *upload session* in stukken van 3,125 MiB), daarna de **JSON-sidecar** en tot slot een kleine registerregel in `_Register`. Er wordt nooit overschreven (`conflictBehavior=fail`);
   5. toont "✓ Opgeslagen" met een link naar het bestand, of een duidelijke Nederlandse foutmelding (geen rechten, map niet gevonden, geen admin consent, te groot, SharePoint druk, enzovoort).

### Bestandsnamen

```
_Inbox/2026-10-09_0732_Offerte-pomp-P-200_3f9a1c2b7d4e5f60.eml
_Inbox/2026-10-09_0732_Offerte-pomp-P-200_3f9a1c2b7d4e5f60.json
_Register/3f9a1c2b7d4e5f60.json
```

Opbouw: `JJJJ-MM-DD_HHMM` (ontvangsttijd in **UTC**), dan het onderwerp (zonder `RE:`/`FW:`/`Antw:`, zonder tekens die SharePoint weigert, maximaal 60 tekens), dan een sleutel van 16 hextekens (FNV-1a-64 van de genormaliseerde `internetMessageId`). Dezelfde mail levert dus altijd dezelfde naam op, ook als een collega hem opslaat.

### Metadata-formaat: JSON-sidecar (schema `arke.mail-metadata/v1`)

**Keuze: een JSON-bestand naast elke `.eml`, geen SharePoint-kolommen.** Waarom:

- **Makkelijk voor deel 2.** In Power Automate is het genoeg om te triggeren op *"Wanneer een bestand wordt gemaakt (alleen eigenschappen)"* in `_Inbox` met een filter op `.json`. Daarna volgen *Bestandsinhoud ophalen* en *JSON parseren*, en dan heb je alle velden plus de bestandsnaam van de `.eml`. De `.json` wordt pas **na** de `.eml` geschreven, dus als de `.json` er staat, is de `.eml` compleet.
- **Geen bibliotheekinrichting en geen extra rechten.** Voor kolommen moet iemand eerst de kolommen in de bibliotheek aanmaken, en de add-in heeft dan een extra Graph-aanroep (listItem/fields) nodig. Dat breekt zodra iemand een kolom hernoemt.
- **Reist mee.** Als deel 2 de mail naar `Klantmail/<Klant>/…` verplaatst, gaat de sidecar mee. Deel 2 kan daarna alsnog kolommen of tags zetten (klant, type, project) voor zoeken en filteren in deel 3.
- **Uitbreidbaar.** Nieuwe velden breken bestaande flows niet. Het veld `schema` geeft de versie aan.

Voorbeeld:

```json
{
  "schema": "arke.mail-metadata/v1",
  "messageKey": "3f9a1c2b7d4e5f60",
  "internetMessageId": "<CAF123@mail.klant.nl>",
  "subject": "RE: Offerte pomp P-200",
  "from": { "name": "Jan Klant", "email": "jan@klant.nl" },
  "to": [{ "name": "Tim Falken", "email": "tfalken@kvt.nl" }],
  "cc": [],
  "receivedAt": "2026-10-09T07:32:10.000Z",
  "conversationId": "AAQkAD…",
  "hasAttachments": true,
  "attachments": [{ "name": "offerte.pdf", "size": 123456, "contentType": "application/pdf" }],
  "savedBy": { "name": "Tim Falken", "email": "tfalken@kvt.nl" },
  "savedAt": "2026-10-09T08:01:44.120Z",
  "customerHint": "Klant BV / 10023",
  "files": {
    "eml": "2026-10-09_0732_Offerte-pomp-P-200_3f9a1c2b7d4e5f60.eml",
    "metadata": "2026-10-09_0732_Offerte-pomp-P-200_3f9a1c2b7d4e5f60.json",
    "emlSize": 187654,
    "emlWebUrl": "https://kvtnl.sharepoint.com/…"
  },
  "source": { "app": "Arke", "version": "1.0.0", "host": "OfficeOnline 16.0…" },
  "processing": { "status": "nieuw" }
}
```

`customerHint` is `null` als de medewerker niets invult. `processing.status` kan deel 2 bijwerken (bijvoorbeeld naar `verwerkt` of `onbekend`).

### Dedupe

- **Primair:** `_Register/<messageKey>.json`. Deze map blijft staan als deel 2 de mails uit `_Inbox` verplaatst, zodat "al opgeslagen" ook dan klopt. Deel 2 hoeft niets met `_Register` te doen. Laat het register vooral **niet** leegmaken.
- **Secundair:** de sidecar met dezelfde naam in `_Inbox`.
- **Vangnet:** uploads gebruiken `conflictBehavior=fail`, dus er wordt nooit iets overschreven.
- Wil je geen register, zet dan `registerFolder: ""` in `config.js`. Dedupe werkt dan alleen zolang de mail nog in `_Inbox` staat.

### Graph-rechten: waarom `Sites.Selected`

| Recht (gedelegeerd) | Wat het betekent | Advies |
|---|---|---|
| **`Sites.Selected`** | De app kan **alleen** bij sites waarop een beheerder hem expliciet toegang geeft, en ook dan nooit meer dan de gebruiker zelf mag. | **Aanbevolen (minimaal).** Kost één extra stap (stap 1e hieronder). |
| `Files.ReadWrite.All` | De app mag namens de gebruiker bij álle bestanden waar die gebruiker bij kan (OneDrive en alle sites). | Eenvoudiger alternatief als stap 1e niet lukt. Het blijft beperkt tot de rechten van de gebruiker, maar de reikwijdte is veel breder. |
| ~~`Mail.Read`~~ | Niet nodig: de `.eml` komt via Office.js (`getAsFileAsync`) met het manifestrecht `ReadItem`. | Niet aanvragen. |

De add-in vraagt alleen de scopes die in `config.js` staan. Wissel je van recht, pas dan alleen `scopes` aan, bijvoorbeeld naar `["https://graph.microsoft.com/Files.ReadWrite.All"]`.

### Ondersteunde Outlook-versies

| Client | Knop + opslaan | Aanmelden |
|---|---|---|
| Outlook op het web | ✅ | NAA (SSO) |
| Nieuwe Outlook voor Windows | ✅ | NAA (SSO) |
| Klassiek Outlook voor Windows, Microsoft 365 | ✅ vanaf **Versie 2404 (build 17530.15000)**, de eerste build met Mailbox 1.14 | NAA bij recente builds, anders de aanmelddialoog |
| Outlook 2024 LTSC (volumelicentie) | ✅ (Mailbox 1.14) | dialoog of NAA |
| Outlook 2021 LTSC / 2019 / 2016 | ❌ te oud (max. Mailbox 1.9). De knop verschijnt, maar het venster meldt "werk Outlook bij". | – |
| Outlook voor Mac (nieuwe interface) | ✅ (Mailbox 1.14) | NAA |
| Outlook iOS/Android | ❌ (de manifest heeft geen mobiele knop; mobiel ondersteunt geen `getAsFileAsync`) | – |

Bron: [Outlook API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). De manifest vraagt bewust minimaal Mailbox 1.1/1.5, zodat de knop overal verschijnt. De add-in controleert pas tijdens het gebruik op 1.14 en geeft dan een nette melding.

---

## Repo-indeling

```
web/                       pagina-root (wordt naar sleutels.kvt.nl/arke/ gedeployd)
  taskpane.html            het taakvenster
  dialog.html              fallback-aanmeldvenster (als NAA niet kan)
  privacy.html             korte privacyverklaring
  config.example.js        voorbeeldconfig → kopieer naar config.js (NIET in git)
  js/core.js               pure logica: bestandsnaam, metadata, dedupe, chunks (getest)
  js/graph.js              Graph-uploads (PUT / upload session)
  js/auth.js               NAA + fallback
  js/taskpane.js           UI en flow
  vendor/msal-browser.min.js  @azure/msal-browser 5.25.0 (MIT)
  assets/                  css + iconen
manifest/
  manifest.template.xml    add-in-only manifest (voor het beheercentrum)
  manifest.template.json   unified manifest (optioneel)
tools/build-manifest.mjs   vult {{HOST_URL}} in → dist/
tests/                     node --test
```

Testen: `npm test` (Node 20+, geen dependencies). De workflow **Tests** draait ze bij elke PR.

---

## Wat Tim moet doen (stap voor stap)

> Tip: doe alles met een account dat **Globale beheerder** of **Toepassingsbeheerder + SharePoint-beheerder** is. Ga ervan uit dat het samen ongeveer een uur kost.

### Stap 1 – App-registratie in Entra ID

1. **a. Registreren.** Ga naar <https://entra.microsoft.com> → **Identiteit → Toepassingen → App-registraties → Nieuwe registratie**.
   - Naam: `Arke – Outlook naar SharePoint`
   - Ondersteunde accounttypen: **Alleen accounts in deze organisatiemap (één tenant)**
   - Omleidings-URI: platform **Single-page application (SPA)**, waarde `brk-multihub://sleutels.kvt.nl`
     (alleen het domein, zonder `/arke`. Dit is de NAA-redirect voor Outlook.)
   - Klik op **Registreren**. Noteer de **Toepassings-id (client)** en de **Map-id (tenant)**.
2. **b. Tweede redirect voor de fallback.** Ga in de app naar **Verificatie → Single-page application → URI toevoegen**: `https://sleutels.kvt.nl/arke/dialog.html`. Klik op **Opslaan**.
   (Laat bij Verificatie de opties *Toegangstokens*/*ID-tokens* onder "Impliciete toekenning" **uit**. Die zijn niet nodig.)
3. **c. API-machtigingen.** Ga naar **API-machtigingen → Een machtiging toevoegen → Microsoft Graph → Gedelegeerde machtigingen**:
   - `Sites.Selected` (aanbevolen; óf `Files.ReadWrite.All`, zie de tabel hierboven)
   - `User.Read` staat er standaard al in. Laat die staan.
   - Er is **geen** clientgeheim of certificaat nodig. Maak die ook niet aan.
4. **d. Admin consent.** Klik op **Beheerderstoestemming verlenen voor Koninklijke van Twist** en bevestig. Achter elke machtiging moet een groen vinkje komen.
5. **e. (Alleen bij `Sites.Selected`) De app toegang geven tot de doelsite.** Zonder deze stap krijgt iedereen "geen schrijfrechten".
   - Open <https://developer.microsoft.com/graph/graph-explorer> en meld je aan als beheerder.
   - Geef Graph Explorer eenmalig het recht `Sites.FullControl.All` (tabblad **Modify permissions**, toestemming geven).
   - Voer uit: `POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions` (site-id: zie stap 2) met als body:
     ```json
     {
       "roles": ["write"],
       "grantedToIdentities": [
         { "application": { "id": "<client-id van Arke>", "displayName": "Arke" } }
       ]
     }
     ```
   - Je krijgt `201 Created` terug. Controleer met `GET …/sites/{site-id}/permissions`.
   - Alternatief in PowerShell (PnP): `Grant-PnPAzureADAppSitePermission -AppId <client-id> -DisplayName Arke -Site https://kvtnl.sharepoint.com/sites/<site> -Permissions Write`.
   - Medewerkers moeten daarnaast zelf **bijdragen** (bewerken) op de site of bibliotheek mogen. `Sites.Selected` geeft nooit méér dan de gebruiker al heeft.

### Stap 2 – De doelsite en de map `_Inbox`

1. Maak (of kies) de SharePoint-site, bijvoorbeeld `https://kvtnl.sharepoint.com/sites/Klantmail` (zie de open vragen), en de documentbibliotheek (standaard *Documenten*).
2. Maak in de bibliotheek de mappen **`_Inbox`** en **`_Register`** aan.
3. **Site-id vinden** in Graph Explorer:
   `GET https://graph.microsoft.com/v1.0/sites/kvtnl.sharepoint.com:/sites/Klantmail`
   → het veld `id` ziet eruit als `kvtnl.sharepoint.com,1111…,2222…`. Dat is de **site-id**.
4. **Drive-id vinden:**
   `GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives`
   → zoek de bibliotheek (bv. `"name": "Documenten"`) en neem het `id` (begint met `b!`). Dat is de **drive-id**.
5. Controle: `GET https://graph.microsoft.com/v1.0/drives/{drive-id}/root:/_Inbox` moet de map teruggeven.
6. Vul `web/config.example.js` in en sla hem op als **`config.js`**: `clientId`, `tenantId`, `scopes`, `siteId`, `driveId`, `inboxFolder` (`_Inbox`) en `registerFolder` (`_Register`). Zet `config.js` eenmalig met een FTP-programma in de map op de server (dezelfde map als `taskpane.html`). Het bestand staat niet in git en de deploy laat het staan.
   Wil je een submap gebruiken, bijvoorbeeld `Klantmail/_Inbox`? Dan zet je dat pad als `inboxFolder`.

### Stap 3 – FTP-secret en de eerste deploy

1. GitHub → repo **Arke** → **Settings → Secrets and variables → Actions → New repository secret**:
   `FTP_REMOTE_DIR` = de map op de FTP-server voor `sleutels.kvt.nl/arke` (zelfde opbouw als bij de andere repo's, bv. `…/arke`). `FTP_HOST`, `FTP_USERNAME` en `FTP_PASSWORD` staan al als organisatie-secrets (zichtbaar voor alle repo's) en hoef je niet opnieuw te zetten.
2. Merge de PR naar `master`. De workflow **Deploy to FTP on master push** zet `web/` live.
3. Upload `config.js` (stap 2.6) en controleer dat <https://sleutels.kvt.nl/arke/taskpane.html> opent. Buiten Outlook doet de pagina niets, maar hij mag geen 404 geven.

### Stap 4 – Uitrol via het Microsoft 365-beheercentrum

1. Maak de definitieve manifest: `npm run manifest -- https://sleutels.kvt.nl/arke` → `dist/manifest.xml`.
   (Of open `manifest/manifest.template.xml` en vervang elke `{{HOST_URL}}` door `https://sleutels.kvt.nl/arke` en `{{HOST_ORIGIN}}` door `https://sleutels.kvt.nl`.)
2. Ga naar <https://admin.microsoft.com> → **Instellingen → Geïntegreerde apps → Aangepaste apps uploaden**.
3. Kies als app-type **Office-invoegtoepassing** en daarna **Manifestbestand uploaden (.xml)** → `dist/manifest.xml`.
4. **Gebruikers kiezen.** Test eerst met **Specifieke gebruikers/groepen** (bv. Tim en Milan) en zet daarna **Hele organisatie** aan. Kies uitrolmethode **Vast** (standaard), zodat gebruikers de add-in niet kunnen verwijderen. **Beschikbaar** mag ook.
5. Accepteer de machtigingen ("item lezen") en klik op **Implementeren**.
6. **Doorlooptijd:** meestal binnen een paar uur, maar Microsoft noemt **tot 24 uur** (soms tot 72 uur) voordat de knop bij iedereen verschijnt. Gebruikers moeten Outlook daarna opnieuw starten. In klassiek Outlook staat de knop op het tabblad *Start* in de groep *Klantmail*, of onder **Apps** als het lint vol is.
7. **Updates:** wijzigingen in `web/` zijn direct live na een merge, zonder nieuwe uitrol. Alleen bij een wijziging in de manifest (knoptekst, iconen, rechten) hoog je `<Version>` op en upload je hem opnieuw bij *Geïntegreerde apps → Arke → Bijwerken*.

### Eerste test (Tim)

1. Open een willekeurige mail in Outlook op het web → **Opslaan in SharePoint** → **Opslaan in SharePoint**.
2. Verwacht: "✓ Opgeslagen". In `_Inbox` staan een `.eml` en een `.json`, in `_Register` één `.json`.
3. Klik nogmaals: de melding wordt "al eerder opgeslagen".
4. Test ook een mail met een grote bijlage (> 4 MB, upload session) en test in klassiek Outlook.

---

## Open vragen (uit de Ponos-taak) met advies

| Vraag | Advies |
|---|---|
| **Welke site/bibliotheek?** | Een **nieuwe, aparte site** `Klantmail` (Teams-loze communicatiesite of teamsite) met één bibliotheek. Gebruik dus niet een bestaande afdelingssite: zo zijn rechten, bewaarbeleid en de Copilot-kennisbron (deel 3) los te regelen. |
| **Wie heeft toegang?** | Twee niveaus. **Schrijven** in `_Inbox` en `_Register`: alle medewerkers die de knop krijgen (anders werkt opslaan niet). **Lezen** van de verwerkte klantmappen: start met de groepen die met klanten werken (verkoop, service, projecten), niet "iedereen". `_Inbox` kun je verbergen of lezen beperken tot de beheerders en het flow-account. Let op: in `_Inbox` kan iedere schrijver ook de mails van collega's zien zolang ze niet verwerkt zijn. Zet daarom deel 2 snel op, of geef de map unieke rechten (*bijdragen zonder lezen* bestaat niet standaard; het alternatief is een aparte uploadbibliotheek waarin alleen de flow leest). |
| **Kiest de medewerker zelf een klant?** | **Ja, optioneel als vrij tekstveld.** Dat zit er nu in (`customerHint`). Het maakt deel 2 betrouwbaarder zonder de medewerker te blokkeren. Een keuzelijst uit BC is een logische vervolgstap (vereist een kleine read-only klanten-endpoint op sleutels.kvt.nl). |
| **Bewaartermijn / privacy?** | Stel een **retentielabel** in op de site (Purview), bijvoorbeeld **7 jaar** voor klant- en ordercorrespondentie (fiscale bewaarplicht), en korter (bv. 2 jaar) voor overige mails. Laat `_Inbox` niet langer dan nodig vol staan. **Ja, leg het voor aan de privacyfunctionaris**: het gaat om persoonsgegevens van klantcontactpersonen en er komt AI-verwerking bij (deel 2). Handmatig opslaan, rechten volgens SharePoint en geen opslag buiten de M365-tenant helpen daarbij. |

---

## Bekende beperkingen

- Alleen in **leesmodus** (niet bij een mail die je nog aan het opstellen bent). Dat is een beperking van `getAsFileAsync`.
- Eén mail tegelijk. Meerdere geselecteerde mails tegelijk opslaan (item multi-select) kan later.
- Er zijn geen kolommen/tags in SharePoint. Dat is werk voor deel 2.
- De sleutel in de bestandsnaam is geen cryptografische hash. Hij dient alleen als stabiele, korte ID.
