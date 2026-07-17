# Facturen App

Een eenvoudige, mobielvriendelijke webapp om facturen en offertes te maken.
Er is geen eigen server of database — de app draait volledig in de browser.
Iedereen logt in met zijn **eigen** account (Google of Microsoft) en krijgt
een eigen, volledig gescheiden archief in zijn eigen cloud-opslag:

| Inloggen met | Opslag van facturen/offertes/klanten/PDF's |
|---|---|
| Google | Map "Facturen App Data" in jouw eigen Google Drive |
| Microsoft (Outlook.com, Hotmail, Microsoft 365/Exchange) | Verborgen app-map in jouw eigen OneDrive |

Naast de gegevens (`settings.json`, `clients.json`, `invoices.json`) wordt ook
een kopie van elke gegenereerde PDF in diezelfde map/app-map opgeslagen.

Er is geen "Inloggen met Apple"/iCloud-optie: Apple biedt geen publieke API
waarmee een losse webapp bestanden mag lezen/schrijven in iemands
persoonlijke iCloud Drive-map, in tegenstelling tot Google en Microsoft.

Je hoeft niet beide providers in te richten. **Als je alleen Google gebruikt
(zoals nu), kun je de Microsoft-stappen gewoon overslaan.**

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

## Gebruik

1. Open de app en kies een inlogknop (Google of Microsoft).
2. Vul eerst bij **Instellingen** je bedrijfsgegevens in (naam, adres,
   btw-nummer, KvK-nummer, IBAN, banknaam, BIC/SWIFT, telefoonnummer, e-mail,
   logo, accentkleur voor de PDF, standaard betalingstermijn en valuta).
3. Voeg eventueel klanten toe bij **Klanten** — of vul bij **Nieuw** de
   klantgegevens direct in en klik op "Opslaan als nieuwe klant" om ze meteen
   te bewaren voor volgende keren.
4. Maak een nieuwe factuur of offerte bij **Nieuw**: kies het type, kies een
   klant (of vul handmatig in), vul optioneel een referentie in en kies de
   taal van het document (Nederlands/Engels), voeg regels toe (met aantal,
   eenheid, prijs en btw-percentage), en sla op. Nummering gebeurt automatisch
   maar is altijd handmatig aan te passen.
5. Download de PDF direct na het aanmaken, of later via **Archief** — filter
   daar op type (factuur/offerte), status of zoek op klantnaam/nummer. In het
   Archief kun je per regel ook direct de status wijzigen of de factuur/
   offerte verwijderen.

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
- **PDF**: wordt bij elke download opnieuw gegenereerd uit de opgeslagen
  gegevens (dus ook oude facturen in het archief zijn altijd opnieuw te
  downloaden), én automatisch als kopie geüpload naar dezelfde Drive/OneDrive-
  map als de brongegevens. Lukt die upload een keer niet (bv. even geen
  internet), dan krijg je een melding maar is de download zelf nooit
  geblokkeerd.

## Bestanden in dit project
- `index.html` — de pagina/app-shell, met login-scherm
- `style.css` — vormgeving (responsive voor telefoon/tablet/laptop)
- `config.js` — jouw Client IDs voor Google/Microsoft
- `app.js` — alle logica (login-providers, opslag, formulieren, PDF-generatie)
- `manifest.json`, `icon-192.png`, `icon-512.png` — voor "zet op beginscherm"
