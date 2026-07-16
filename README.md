# Facturen App

Een eenvoudige, mobielvriendelijke webapp om facturen en offertes te maken.
Alle gegevens (klanten, instellingen, facturen/offertes) worden bewaard in een
map in **jouw eigen Google Drive** ("Facturen App Data"), zodat je hetzelfde
archief ziet op je iPhone, iPad en laptop.

Er is geen eigen server of database nodig — de app draait volledig in de
browser en praat rechtstreeks met de Google Drive API.

## Wat moet je zelf instellen (eenmalig, ± 10 minuten)

De app heeft een eigen "sleutel" (OAuth Client ID) nodig om met jouw Google
account te mogen praten. Dit kan alleen jijzelf aanmaken (ik kan geen Google
account namens jou aanmaken).

### Stap 1 — Google Cloud project aanmaken
1. Ga naar https://console.cloud.google.com/
2. Log in met het Google-account waar je facturen in wilt bewaren.
3. Maak een nieuw project aan, bijvoorbeeld genaamd "Facturen App".

### Stap 2 — Drive API inschakelen
1. Ga naar "APIs & Services" → "Library".
2. Zoek "Google Drive API" en klik op **Enable**.

### Stap 3 — OAuth-consentscherm instellen
1. Ga naar "APIs & Services" → "OAuth consent screen".
2. Kies **External** (tenzij je een Google Workspace-organisatie hebt, dan mag Internal ook).
3. Vul een appnaam in (bv. "Facturen App") en je eigen e-mailadres bij support/contact.
4. Sla op. Je hoeft de app niet te publiceren/laten verifiëren: laat de status op
   **Testing** staan.
5. Voeg jezelf (en eventuele collega's die de app mogen gebruiken) toe onder
   **Test users**.

### Stap 4 — OAuth Client ID aanmaken
1. Ga naar "APIs & Services" → "Credentials" → "Create Credentials" → "OAuth client ID".
2. Kies applicatietype **Web application**.
3. Bij **Authorized JavaScript origins** voeg je de URL toe waar de app straks
   bereikbaar is, bijvoorbeeld:
   - `http://localhost:5500` (voor lokaal testen)
   - `https://jouw-app-naam.vercel.app` (na het hosten, zie Stap 6)
4. Klik op **Create** en kopieer de **Client ID** (eindigt op `.apps.googleusercontent.com`).

### Stap 5 — Client ID invullen in de app
Open `config.js` in deze map en vervang:
```js
export const GOOGLE_CLIENT_ID = "VUL_HIER_JE_GOOGLE_CLIENT_ID_IN.apps.googleusercontent.com";
```
door je eigen Client ID.

### Stap 6 — Hosten (gratis)
De app is een statische website (alleen HTML/CSS/JS) en kan gratis gehost
worden, bijvoorbeeld via **Vercel**:
1. Ga naar https://vercel.com en log in (kan met je Google-account).
2. Maak een nieuw project en upload deze map (`facturen-app`), of verbind een
   GitHub-repository met deze bestanden.
3. Na het deployen krijg je een URL zoals `https://jouw-app-naam.vercel.app`.
4. Ga terug naar stap 4 in Google Cloud Console en zet deze URL bij
   **Authorized JavaScript origins** (en klik op Save).

Alternatieven: Netlify, GitHub Pages, Cloudflare Pages — werkt allemaal, zolang
het een HTTPS-adres oplevert dat je kunt invullen bij "Authorized JavaScript origins".

### Stap 7 — Toevoegen aan beginscherm (iPhone/iPad)
Open de gehoste URL in Safari op je iPhone/iPad → deelknop → "Zet op
beginscherm". De app krijgt dan een eigen icoon en opent zonder browserbalk,
net als een echte app.

## Gebruik

1. Open de app en log in met Google (eerste keer moet je akkoord gaan met de
   gevraagde toestemming — de app vraagt alleen toegang tot bestanden die de
   app zelf aanmaakt, niet tot de rest van je Drive).
2. Vul eerst bij **Instellingen** je bedrijfsgegevens in (naam, adres,
   btw-nummer, IBAN, logo, standaard betalingstermijn).
3. Voeg eventueel klanten toe bij **Klanten**.
4. Maak een nieuwe factuur of offerte bij **Nieuw**: kies het type, kies een
   klant (of vul handmatig in), voeg regels toe, kies btw-percentages, en sla
   op. Nummering gebeurt automatisch maar is altijd handmatig aan te passen.
5. Download de PDF direct na het aanmaken, of later via **Archief** — filter
   daar op type (factuur/offerte), status of zoek op klantnaam/nummer.

## Belangrijk om te weten

- **Inloggen**: om veiligheidsredenen moet je opnieuw inloggen als je de app
  in een nieuw tabblad/sessie opent (het toegangstoken wordt niet permanent
  op het apparaat bewaard). Binnen dezelfde sessie hoef je niet steeds opnieuw
  in te loggen.
- **Testgebruikers**: zolang het OAuth-consentscherm op "Testing" staat,
  kunnen alleen de e-mailadressen die je bij Stap 3 hebt toegevoegd inloggen.
  Wil je dat meer mensen (bv. een boekhouder) de app kunnen gebruiken? Voeg ze
  toe als test user, of laat het consentscherm verifiëren door Google
  (alleen nodig bij veel gebruikers).
- **Data**: alles staat in de map "Facturen App Data" in jouw Google Drive.
  Verwijder deze map niet handmatig, tenzij je alles kwijt wilt.
- **PDF**: wordt altijd opnieuw gegenereerd uit de opgeslagen gegevens, dus
  ook oude facturen in het archief zijn altijd opnieuw te downloaden.

## Bestanden in dit project
- `index.html` — de pagina/app-shell
- `style.css` — vormgeving (responsive voor telefoon/tablet/laptop)
- `config.js` — jouw Google Client ID (stap 5 hierboven)
- `app.js` — alle logica (inloggen, Drive-opslag, formulieren, PDF-generatie)
- `manifest.json`, `icon-192.png`, `icon-512.png` — voor "zet op beginscherm"
