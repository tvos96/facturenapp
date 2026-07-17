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
export const MICROSOFT_CLIENT_ID = "5f09388c-7728-49c4-8214-50ab592eb721";

// "common" ondersteunt zowel persoonlijke Microsoft-accounts (Outlook.com,
// Hotmail) als werk-/schoolaccounts (Microsoft 365 / Exchange).
export const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";

// Files.ReadWrite.AppFolder = alleen toegang tot een eigen map
// ("OneDrive > Apps > Facturen App"), niet tot de rest van iemands OneDrive.
export const MICROSOFT_SCOPES = ["Files.ReadWrite.AppFolder", "User.Read"];
