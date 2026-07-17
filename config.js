// ==========================================================================
// CONFIGURATIE
// ==========================================================================
// Vul hier je eigen Google OAuth Client ID in. Zie README.md voor de
// stap-voor-stap uitleg hoe je die aanmaakt (gratis, Google Cloud Console).
//
// Voorbeeld: "123456789012-abcdefghijklmnopqrstuvwxyz123456.apps.googleusercontent.com"
export const GOOGLE_CLIENT_ID = "811734404657-o2v6ca8gkva2ehmlvvu78sf3gvu37u2l.apps.googleusercontent.com";

// Scope: de app kan alleen bestanden lezen/schrijven die ze zelf heeft
// aangemaakt in jouw Google Drive. Geen toegang tot de rest van je Drive.
export const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";

// Naam van de map die automatisch in je Google Drive wordt aangemaakt
// en waarin alle gegevens van de app worden bewaard.
export const APP_FOLDER_NAME = "Facturen App Data";

// Standaard btw-percentages die in de keuzelijst verschijnen.
export const DEFAULT_VAT_RATES = [21, 9, 0];

// --------------------------------------------------------------------------
// MICROSOFT (Outlook / Exchange / OneDrive)
// --------------------------------------------------------------------------
// Vul hier je Azure "Application (client) ID" in. Zie README.md voor de
// stap-voor-stap uitleg (gratis, Azure Portal - App registrations).
export const MICROSOFT_CLIENT_ID = "VUL_HIER_JE_MICROSOFT_CLIENT_ID_IN";

// "common" ondersteunt zowel persoonlijke Microsoft-accounts (Outlook.com,
// Hotmail) als werk-/schoolaccounts (Microsoft 365 / Exchange).
export const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";

// Files.ReadWrite.AppFolder = alleen toegang tot een eigen map
// ("OneDrive > Apps > Facturen App"), niet tot de rest van iemands OneDrive.
export const MICROSOFT_SCOPES = ["Files.ReadWrite.AppFolder", "User.Read"];

// --------------------------------------------------------------------------
// APPLE (Inloggen met Apple - alleen identiteit, geen bestandsopslag)
// --------------------------------------------------------------------------
// Vereist een betaald Apple Developer Program account ($99/jaar) + een
// geverifieerd domein. Zie README.md. Laat op de placeholder staan als je
// dit (nog) niet gebruikt - de Apple-knop toont dan een duidelijke melding.
export const APPLE_CLIENT_ID = "VUL_HIER_JE_APPLE_SERVICES_ID_IN"; // bv. nl.zig.facturenapp.web
export const APPLE_REDIRECT_URI = "VUL_HIER_JE_GEHOSTE_APP_URL_IN"; // exact zoals bij Apple geconfigureerd, bv. https://tvos96.github.io/facturenapp/
