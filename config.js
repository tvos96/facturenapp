// ==========================================================================
// CONFIGURATIE
// ==========================================================================
// Vul hier je eigen Google OAuth Client ID in. Zie README.md voor de
// stap-voor-stap uitleg hoe je die aanmaakt (gratis, Google Cloud Console).
//
// Voorbeeld: "123456789012-abcdefghijklmnopqrstuvwxyz123456.apps.googleusercontent.com"
export const GOOGLE_CLIENT_ID = "VUL_HIER_JE_GOOGLE_CLIENT_ID_IN.apps.googleusercontent.com";

// Scope: de app kan alleen bestanden lezen/schrijven die ze zelf heeft
// aangemaakt in jouw Google Drive. Geen toegang tot de rest van je Drive.
export const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";

// Naam van de map die automatisch in je Google Drive wordt aangemaakt
// en waarin alle gegevens van de app worden bewaard.
export const APP_FOLDER_NAME = "Facturen App Data";

// Standaard btw-percentages die in de keuzelijst verschijnen.
export const DEFAULT_VAT_RATES = [21, 9, 0];
