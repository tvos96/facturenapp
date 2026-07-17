# Facturen App

Een eenvoudige, mobielvriendelijke webapp om facturen en offertes te maken.
Er is geen eigen server of database — de app draait volledig in de browser.
Iedereen logt in met zijn **eigen** account (Google of Microsoft) en krijgt
een eigen, volledig gescheiden archief in zijn eigen cloud-opslag:

| Inloggen met | Opslag van facturen/offertes/klanten |
|---|---|
| Google | Map "Facturen App Data" in jouw eigen Google Drive |
| Microsoft (Outlook.com, Hotmail, Microsoft 365/Exchange) | Verborgen app-map in jouw eigen OneDrive |
| Apple | Alleen identiteit — je kiest daarna zelf Google Drive of OneDrive als opslag (zie onder) |

**Waarom geen echte iCloud-opslag?** Apple biedt (in tegenstelling tot Google
en Microsoft) geen publieke API waarmee een losse webapp bestanden mag
lezen/schrijven in iemands persoonlijke iCloud Drive-map. "Inloggen met
Apple" werkt daarom alleen als identiteit (je naam/mailadres); de facturen
zelf worden dan alsnog in Google Drive of OneDrive bewaard, gekoppeld aan dat
Apple-account.

Je hoeft niet alle drie de providers in te richten. **Als je alleen Google
gebruikt (zoals nu), kun je de Microsoft- en Apple-stappen gewoon overslaan.**

---

## Google — al ingericht ✅

Client ID staat al in `config.js`. Niks meer te doen, tenzij je een nieuw
Google Cloud-project wilt gebruiken. Kort samengevat was dit de opzet:
1. Project + Drive API aan in [Google Cloud Console](https://console.cloud.google.com/).
2. OAuth-consentscherm op "Testing", met je eigen mailadres(sen) als **Test user**.
3. OAuth Client ID (Web application) met bij **Authorized JavaScript origins**
   exact `https://tvos96.github.io` (dus zonder `/facturenapp/` erachter).

---

## Microsoft (Outlook / Exchange / OneDrive) — optioneel

### Stap 1 — App registreren in Azure
1. Ga naar https://portal.azure.com/ en log in (een gratis Microsoft-account volstaat).
2. Ga naar **Microsoft Entra ID** → **App registrations** → **New registration**.
3. Naam: bv. "Facturen App".
4. **Supported account types**: kies **Accounts in any organizational
   directory and personal Microsoft accounts** (dit is nodig zodat zowel
   privé Outlook/Hotmail-accounts als werk/school-accounts met Exchange
   kunnen inloggen).
5. **Redirect URI**: kies platformtype **Single-page application (SPA)** en
   vul in: `https://tvos96.github.io/facturenapp/`
6. Klik **Register** en kopieer de **Application (client) ID**.

### Stap 2 — API-rechten instellen
1. Ga naar **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Voeg toe: `Files.ReadWrite.AppFolder` en `User.Read`.
3. Bij een persoonlijk/eigen account is geen admin-goedkeuring nodig; bij een
   zakelijk account kan IT-goedkeuring vereist zijn.

### Stap 3 — Client ID invullen
Open `config.js` en vervang:
```js
export const MICROSOFT_CLIENT_ID = "VUL_HIER_JE_MICROSOFT_CLIENT_ID_IN";
```
door de Application (client) ID uit Stap 1.

Dat is alles — er is geen aparte hosting-stap nodig, dit draait op dezelfde
GitHub Pages-URL als de rest van de app.

---

## Apple ("Inloggen met Apple") — optioneel, betaald

Dit vereist een **Apple Developer Program**-account (€99/jaar) en kost
duidelijk meer moeite dan Google/Microsoft. Sla dit gerust over als het
alleen Google/Microsoft-gebruikers betreft.

### Stap 1 — Apple Developer account
1. Meld je aan op https://developer.apple.com/programs/ (betaald, jaarlijks).

### Stap 2 — Services ID aanmaken
1. Ga naar **Certificates, Identifiers & Profiles** → **Identifiers** → **+**.
2. Kies **Services IDs**, geef een naam en een identifier, bv. `nl.zig.facturenapp.web`.
3. Vink **Sign in with Apple** aan en klik **Configure**:
   - **Primary App ID**: koppel aan een App ID (moet je eventueel eerst apart aanmaken).
   - **Domains and Subdomains**: `tvos96.github.io`
   - **Return URLs**: `https://tvos96.github.io/facturenapp/`
4. Sla op.

### Stap 3 — Domeinverificatie
Apple vraagt een verificatiebestand te hosten op jouw domein
(`.well-known/apple-developer-domain-association.txt`). Zet dat bestand in
de root van je GitHub Pages-repository zodat het bereikbaar is op
`https://tvos96.github.io/.well-known/apple-developer-domain-association.txt`.
Let op: dit moet in de root van `tvos96.github.io` staan, niet per se binnen
`/facturenapp/` — dit hangt af van hoe je repository is ingericht.

### Stap 4 — Client ID invullen
Open `config.js` en vervang:
```js
export const APPLE_CLIENT_ID = "VUL_HIER_JE_APPLE_SERVICES_ID_IN";
export const APPLE_REDIRECT_URI = "VUL_HIER_JE_GEHOSTE_APP_URL_IN";
```
door je Services ID (bv. `nl.zig.facturenapp.web`) en de exacte URL
(`https://tvos96.github.io/facturenapp/`).

**Beperking**: Apple geeft de naam van de gebruiker alléén door bij de
allereerste keer inloggen ooit met dat Apple ID bij deze app. Daarna toont de
app alleen nog het e-mailadres.

---

## Gebruik

1. Open de app en kies een inlogknop (Google, Microsoft, of Apple).
   - Bij Apple: kies daarna nog Google Drive of OneDrive als opslag.
2. Vul eerst bij **Instellingen** je bedrijfsgegevens in (naam, adres,
   btw-nummer, IBAN, logo, standaard betalingstermijn).
3. Voeg eventueel klanten toe bij **Klanten**.
4. Maak een nieuwe factuur of offerte bij **Nieuw**: kies het type, kies een
   klant (of vul handmatig in), voeg regels toe, kies btw-percentages, en sla
   op. Nummering gebeurt automatisch maar is altijd handmatig aan te passen.
5. Download de PDF direct na het aanmaken, of later via **Archief** — filter
   daar op type (factuur/offerte), status of zoek op klantnaam/nummer.

## Belangrijk om te weten

- **Eigen archief per persoon**: iedereen die inlogt gebruikt zijn eigen
  Google Drive- of OneDrive-account als opslag. Er is geen gedeeld archief
  tussen verschillende mensen — als twee collega's allebei inloggen (elk met
  hun eigen account), zien ze elk hún eigen facturen, niet die van de ander.
- **Inloggen**: om veiligheidsredenen moet je bij Google soms opnieuw inloggen
  in een nieuw tabblad/sessie. Microsoft onthoudt de sessie iets langer dankzij
  ingebouwde single sign-on.
- **Google Testgebruikers**: zolang het OAuth-consentscherm op "Testing"
  staat, kunnen alleen toegevoegde e-mailadressen inloggen (zie Google-sectie
  hierboven).
- **PDF**: wordt altijd opnieuw gegenereerd uit de opgeslagen gegevens, dus
  ook oude facturen in het archief zijn altijd opnieuw te downloaden.

## Bestanden in dit project
- `index.html` — de pagina/app-shell, met login-scherm voor alle providers
- `style.css` — vormgeving (responsive voor telefoon/tablet/laptop)
- `config.js` — jouw Client IDs voor Google/Microsoft/Apple
- `app.js` — alle logica (login-providers, opslag, formulieren, PDF-generatie)
- `manifest.json`, `icon-192.png`, `icon-512.png` — voor "zet op beginscherm"
