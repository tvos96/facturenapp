import {
  GOOGLE_CLIENT_ID,
  DRIVE_SCOPE,
  APP_FOLDER_NAME,
  DEFAULT_VAT_RATES,
  MICROSOFT_CLIENT_ID,
  MICROSOFT_AUTHORITY,
  MICROSOFT_SCOPES,
} from "./config.js";

// ==========================================================================
// STATE
// ==========================================================================
const state = {
  activeProviderId: null, // 'google' | 'microsoft'
  identity: { label: "" }, // label getoond rechtsboven in de app
  fileIds: { settings: null, clients: null, invoices: null },
  settings: null,
  clients: [],
  invoices: [],
  view: "new",
  draft: null,
  archiveFilter: { type: "alle", status: "alle", q: "" },
};

const DEFAULT_SETTINGS = {
  companyName: "",
  // Adres als losse velden (internationaal bruikbare indeling: straat +
  // huisnummer, postcode, plaats, land), i.p.v. één vrij tekstveld.
  addressLine: "",
  postalCode: "",
  city: "",
  country: "",
  vatNumber: "",
  kvkNumber: "",
  iban: "",
  bankName: "",
  bic: "",
  phone: "",
  email: "",
  logoDataUrl: "",
  accentColor: "#1d4ed8",
  currency: "EUR",
  paymentTermDays: 14,
  nextFactuurNummer: 1,
  nextOfferteNummer: 1,
  numberPrefix: { factuur: "F", offerte: "O" },
  // Notitiesjablonen: elk sjabloon heeft een tekst in het Nederlands en in het
  // Engels, zodat de juiste taalversie op de PDF komt. defaultNoteTemplateId
  // wijst naar het sjabloon dat automatisch wordt ingevuld bij een nieuwe
  // factuur/offerte.
  noteTemplates: [
    {
      id: "default-standaard",
      name: "Standaard",
      nl: "Bedankt voor uw vertrouwen. Heeft u vragen, neem dan gerust contact met ons op.",
      en: "Thank you for your business. If you have any questions, please don't hesitate to contact us.",
    },
  ],
  defaultNoteTemplateId: "default-standaard",
};

const STATUS_OPTIONS = {
  factuur: ["concept", "verzonden", "betaald"],
  offerte: ["concept", "verzonden", "geaccepteerd", "afgewezen"],
};

const STATUS_LABELS = {
  concept: "Concept",
  verzonden: "Verzonden",
  betaald: "Betaald",
  geaccepteerd: "Geaccepteerd",
  afgewezen: "Afgewezen",
};

// Vaste eenheid-opties per regel. "custom" (zelf in te vullen) heeft geen
// vertaling: de ingevoerde tekst wordt letterlijk overgenomen, ongeacht de
// documenttaal.
const UNIT_PRESETS = [
  { key: "dagen", nl: "Dag(en)", en: "Day(s)" },
  { key: "km", nl: "km", en: "km" },
  { key: "eenheden", nl: "eenheid", en: "unit" },
];

function unitDisplayLabel(item, lang) {
  const preset = UNIT_PRESETS.find((u) => u.key === item.unitKey);
  if (preset) return preset[lang] || preset.nl;
  return item.unit || "";
}

// Zet de losse "van"/"tot" datumkiezers van een regel om naar één leesbare
// tekst voor op de PDF: een periode als beide zijn ingevuld en verschillen,
// anders gewoon de ene datum die er wél is.
function itemDateLabel(item) {
  if (!item.dateFrom && !item.dateTo) return "";
  if (item.dateFrom && item.dateTo && item.dateTo !== item.dateFrom) {
    return fmtDate(item.dateFrom) + " - " + fmtDate(item.dateTo);
  }
  return fmtDate(item.dateFrom || item.dateTo);
}

function activeProvider() {
  return providers[state.activeProviderId];
}

// ==========================================================================
// PROVIDER: GOOGLE DRIVE
// ==========================================================================
const googleProvider = {
  id: "google",
  name: "Google Drive",
  token: null,
  tokenClient: null,
  folderId: null,

  init() {
    if (typeof google === "undefined" || !google.accounts) return;
    this.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: DRIVE_SCOPE,
      callback: () => {}, // wordt per login-poging overschreven
    });
  },

  login() {
    return new Promise((resolve, reject) => {
      if (GOOGLE_CLIENT_ID.includes("VUL_HIER")) {
        reject(new Error("Er is nog geen Google Client ID ingevuld in config.js. Zie README.md."));
        return;
      }
      if (!this.tokenClient) {
        reject(new Error("Google Identity Services kon niet geladen worden. Controleer je internetverbinding."));
        return;
      }
      this.tokenClient.callback = (resp) => {
        if (resp.error) {
          reject(new Error(resp.error));
          return;
        }
        this.token = resp.access_token;
        sessionStorage.setItem("google_token", this.token);
        sessionStorage.setItem("google_token_exp", String(Date.now() + (resp.expires_in - 60) * 1000));
        resolve();
      };
      this.tokenClient.requestAccessToken({ prompt: "" });
    });
  },

  async restoreSession() {
    const savedToken = sessionStorage.getItem("google_token");
    const savedExp = Number(sessionStorage.getItem("google_token_exp") || 0);
    if (savedToken && savedExp > Date.now()) {
      this.token = savedToken;
      this.folderId = sessionStorage.getItem("google_folder_id") || null;
      return true;
    }
    return false;
  },

  logout() {
    if (this.token && typeof google !== "undefined") {
      google.accounts.oauth2.revoke(this.token, () => {});
    }
    this.token = null;
    sessionStorage.removeItem("google_token");
    sessionStorage.removeItem("google_token_exp");
    sessionStorage.removeItem("google_folder_id");
  },

  async request(url, options = {}) {
    const res = await fetch(url, {
      ...options,
      headers: { Authorization: "Bearer " + this.token, ...(options.headers || {}) },
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`Google Drive fout (${res.status}): ${text.slice(0, 200)}`);
    }
    return res;
  },

  async findByName(name, parentId) {
    const q = encodeURIComponent(
      `name='${name.replace(/'/g, "\\'")}' and trashed=false` + (parentId ? ` and '${parentId}' in parents` : "")
    );
    const res = await this.request(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&spaces=drive`);
    const data = await res.json();
    return data.files && data.files.length ? data.files[0] : null;
  },

  async ensureStorage() {
    let folder = await this.findByName(APP_FOLDER_NAME, null);
    if (!folder) {
      const res = await this.request("https://www.googleapis.com/drive/v3/files", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: APP_FOLDER_NAME, mimeType: "application/vnd.google-apps.folder" }),
      });
      folder = await res.json();
    }
    this.folderId = folder.id;
    sessionStorage.setItem("google_folder_id", folder.id);
  },

  async readJSON(fileName, fallback) {
    const file = await this.findByName(fileName, this.folderId);
    if (!file) return { fileId: null, data: fallback };
    const res = await this.request(`https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`);
    return { fileId: file.id, data: await res.json() };
  },

  async writeJSON(fileName, fileId, data) {
    const content = JSON.stringify(data, null, 2);
    if (fileId) {
      await this.request(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: content,
      });
      return fileId;
    }
    const boundary = "-------facturenapp" + Date.now();
    const metadata = { name: fileName, parents: [this.folderId] };
    const body =
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\nContent-Type: application/json\r\n\r\n${content}\r\n--${boundary}--`;
    const res = await this.request("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
      method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
      body,
    });
    const created = await res.json();
    return created.id;
  },

  // Slaat een binair bestand (bv. de gegenereerde PDF) op in dezelfde map.
  async writeBinary(fileName, blob) {
    const existing = await this.findByName(fileName, this.folderId);
    if (existing) {
      await this.request(`https://www.googleapis.com/upload/drive/v3/files/${existing.id}?uploadType=media`, {
        method: "PATCH",
        headers: { "Content-Type": "application/pdf" },
        body: blob,
      });
      return existing.id;
    }
    const boundary = "-------facturenapp" + Date.now();
    const metadata = { name: fileName, parents: [this.folderId] };
    const head =
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\nContent-Type: application/pdf\r\n\r\n`;
    const tail = `\r\n--${boundary}--`;
    const body = new Blob([head, blob, tail]);
    const res = await this.request("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
      method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
      body,
    });
    const created = await res.json();
    return created.id;
  },

  // Verwijdert een bestand (bv. een PDF) uit de app-map, als het bestaat.
  async deleteFile(fileName) {
    const file = await this.findByName(fileName, this.folderId);
    if (!file) return;
    await this.request(`https://www.googleapis.com/drive/v3/files/${file.id}`, { method: "DELETE" });
  },
};

// ==========================================================================
// PROVIDER: MICROSOFT ONEDRIVE (Outlook / Exchange / Microsoft 365)
// ==========================================================================
let msalInstance = null;

const microsoftProvider = {
  id: "microsoft",
  name: "OneDrive",
  token: null,
  account: null,

  init() {
    if (typeof msal === "undefined") return;
    if (MICROSOFT_CLIENT_ID.includes("VUL_HIER")) return; // nog niet geconfigureerd, sla stil over
    try {
      msalInstance = new msal.PublicClientApplication({
        auth: {
          clientId: MICROSOFT_CLIENT_ID,
          authority: MICROSOFT_AUTHORITY,
          redirectUri: window.location.origin + window.location.pathname,
        },
        cache: { cacheLocation: "sessionStorage" },
      });
    } catch (e) {
      console.warn("Microsoft-configuratie kon niet worden geïnitialiseerd", e);
    }
  },

  async login() {
    if (MICROSOFT_CLIENT_ID.includes("VUL_HIER")) {
      throw new Error("Er is nog geen Microsoft Client ID ingevuld in config.js. Zie README.md.");
    }
    if (!msalInstance) {
      throw new Error("Microsoft-inlogbibliotheek kon niet geladen worden. Controleer je internetverbinding.");
    }
    const resp = await msalInstance.loginPopup({ scopes: MICROSOFT_SCOPES });
    this.account = resp.account;
    await this.acquireToken();
  },

  async acquireToken() {
    try {
      const resp = await msalInstance.acquireTokenSilent({ scopes: MICROSOFT_SCOPES, account: this.account });
      this.token = resp.accessToken;
    } catch (e) {
      const resp = await msalInstance.acquireTokenPopup({ scopes: MICROSOFT_SCOPES, account: this.account });
      this.token = resp.accessToken;
    }
  },

  async restoreSession() {
    if (!msalInstance) return false;
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return false;
    this.account = accounts[0];
    try {
      await this.acquireToken();
      return true;
    } catch (e) {
      return false;
    }
  },

  logout() {
    this.token = null;
    if (msalInstance && this.account) {
      msalInstance.logoutPopup({ account: this.account }).catch(() => {});
    }
  },

  // approot = speciale, verborgen "Apps"-map per gebruiker; wordt door
  // Microsoft automatisch aangemaakt zodra we er iets in zetten.
  async ensureStorage() {
    // niets te doen: approot bestaat impliciet
  },

  async graphRequest(path, options = {}, allow404 = false) {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: { Authorization: "Bearer " + this.token, ...(options.headers || {}) },
    });
    if (res.status === 404 && allow404) return null;
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`OneDrive fout (${res.status}): ${text.slice(0, 200)}`);
    }
    return res;
  },

  async readJSON(fileName, fallback) {
    const res = await this.graphRequest(`/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`, {}, true);
    if (!res) return { fileId: null, data: fallback };
    return { fileId: fileName, data: await res.json() };
  },

  async writeJSON(fileName, _fileId, data) {
    await this.graphRequest(`/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data, null, 2),
    });
    return fileName; // OneDrive-bestanden zijn pad-gebaseerd, geen apart id nodig
  },

  // Slaat een binair bestand (bv. de gegenereerde PDF) op in dezelfde app-map.
  async writeBinary(fileName, blob) {
    await this.graphRequest(`/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/pdf" },
      body: blob,
    });
    return fileName;
  },

  // Verwijdert een bestand (bv. een PDF) uit de app-map, als het bestaat.
  async deleteFile(fileName) {
    await this.graphRequest(`/me/drive/special/approot:/${encodeURIComponent(fileName)}:`, { method: "DELETE" }, true);
  },
};

const providers = { google: googleProvider, microsoft: microsoftProvider };

// ==========================================================================
// LOGIN WIRING
// ==========================================================================
function setupLoginButtons() {
  document.getElementById("login-google").addEventListener("click", () => doProviderLogin("google"));
  document.getElementById("login-microsoft").addEventListener("click", () => doProviderLogin("microsoft"));

  document.getElementById("logout-btn").addEventListener("click", () => {
    const p = activeProvider();
    if (p) p.logout();
    sessionStorage.removeItem("active_provider");
    location.reload();
  });
}

async function doProviderLogin(providerId) {
  showLoginError("");
  const provider = providers[providerId];
  if (!provider) return;
  try {
    setSyncStatus("Inloggen...");
    await provider.login();
    await provider.ensureStorage();
    state.activeProviderId = providerId;
    state.identity = { label: provider.name };
    sessionStorage.setItem("active_provider", providerId);
    setSyncStatus("");
    await afterLogin();
  } catch (err) {
    console.error(err);
    setSyncStatus("");
    showLoginError("Inloggen mislukt: " + err.message);
  }
}

function showLoginError(msg) {
  document.getElementById("login-error").textContent = msg;
}

async function tryRestoreSession() {
  const savedProviderId = sessionStorage.getItem("active_provider");
  if (!savedProviderId) return false;
  const provider = providers[savedProviderId];
  if (!provider) return false;
  const ok = await provider.restoreSession();
  if (!ok) return false;
  state.activeProviderId = savedProviderId;
  state.identity = { label: provider.name };
  return true;
}

async function afterLogin() {
  document.getElementById("login-screen").classList.add("hidden");
  document.getElementById("app").classList.remove("hidden");
  document.getElementById("account-label").textContent = "Opgeslagen in: " + state.identity.label;
  setSyncStatus("Synchroniseren...");
  try {
    await loadAllData();
    setSyncStatus("");
    renderView();
  } catch (err) {
    console.error(err);
    setSyncStatus("");
    toast("Fout bij laden van je opslag: " + err.message);
  }
}

function setSyncStatus(txt) {
  document.getElementById("sync-status").textContent = txt;
}

// Zet een oud, vrij ingevuld adres (meerdere regels tekst) om naar de
// nieuwe losse velden. Best-effort: regel 1 = straat, regel 2 = "postcode
// plaats", regel 3 = land. Wordt alleen gebruikt als de nieuwe velden nog
// helemaal leeg zijn, zodat bestaande data (Instellingen, Klanten, oude
// facturen/offertes) niet stilzwijgend hun adres kwijtraken.
function migrateLegacyAddress(obj) {
  if (!obj) return obj;
  const hasNewFields = obj.addressLine || obj.postalCode || obj.city || obj.country;
  if (hasNewFields || !obj.address) return obj;
  const lines = String(obj.address).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  obj.addressLine = lines[0] || "";
  if (lines[1]) {
    const m = lines[1].match(/^([0-9]{3,5}\s?[A-Za-z]{0,3})\s+(.*)$/);
    if (m) {
      obj.postalCode = m[1].trim();
      obj.city = m[2].trim();
    } else {
      obj.city = lines[1];
    }
  }
  if (lines[2]) obj.country = lines[2];
  return obj;
}

// ==========================================================================
// STORE (settings / clients / invoices) - provider-onafhankelijk
// ==========================================================================
async function loadAllData() {
  const p = activeProvider();
  const [settingsRes, clientsRes, invoicesRes] = await Promise.all([
    p.readJSON("settings.json", DEFAULT_SETTINGS),
    p.readJSON("clients.json", []),
    p.readJSON("invoices.json", []),
  ]);
  state.settings = { ...DEFAULT_SETTINGS, ...settingsRes.data };
  migrateLegacyAddress(state.settings);
  state.fileIds.settings = settingsRes.fileId;
  state.clients = clientsRes.data || [];
  state.clients.forEach(migrateLegacyAddress);
  state.fileIds.clients = clientsRes.fileId;
  state.invoices = invoicesRes.data || [];
  state.invoices.forEach((inv) => migrateLegacyAddress(inv.clientSnapshot));
  state.fileIds.invoices = invoicesRes.fileId;
  // Basislijn voor de "ongeopslagen wijzigingen"-check in Instellingen.
  state.settingsSavedSnapshot = JSON.stringify(state.settings);
}

async function saveSettings() {
  setSyncStatus("Opslaan...");
  state.fileIds.settings = await activeProvider().writeJSON("settings.json", state.fileIds.settings, state.settings);
  state.settingsSavedSnapshot = JSON.stringify(state.settings);
  setSyncStatus("");
}

// Zijn er in Instellingen wijzigingen gemaakt die nog niet zijn opgeslagen?
function hasUnsavedSettings() {
  if (state.settingsSavedSnapshot === undefined) return false;
  return JSON.stringify(state.settings) !== state.settingsSavedSnapshot;
}

async function saveClients() {
  setSyncStatus("Opslaan...");
  state.fileIds.clients = await activeProvider().writeJSON("clients.json", state.fileIds.clients, state.clients);
  setSyncStatus("");
}

async function saveInvoices() {
  setSyncStatus("Opslaan...");
  state.fileIds.invoices = await activeProvider().writeJSON("invoices.json", state.fileIds.invoices, state.invoices);
  setSyncStatus("");
}

// ==========================================================================
// HELPERS: nummering, datums, totalen
// ==========================================================================
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function addDays(isoDate, days) {
  const d = new Date(isoDate + "T00:00:00");
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString().slice(0, 10);
}

function suggestNumber(type) {
  const year = new Date().getFullYear();
  const n = type === "factuur" ? state.settings.nextFactuurNummer : state.settings.nextOfferteNummer;
  const prefix = state.settings.numberPrefix[type];
  return `${prefix}${year}-${String(n).padStart(3, "0")}`;
}

function bumpNumberCounterIfMatches(type, usedNumber) {
  if (usedNumber === suggestNumber(type)) {
    if (type === "factuur") state.settings.nextFactuurNummer++;
    else state.settings.nextOfferteNummer++;
  }
}

function calcTotals(items) {
  const groups = {};
  let subtotal = 0;
  for (const it of items) {
    const qty = Number(it.qty) || 0;
    const price = Number(it.price) || 0;
    const rate = Number(it.vatRate) || 0;
    const lineTotal = qty * price;
    subtotal += lineTotal;
    if (!groups[rate]) groups[rate] = { subtotal: 0, vat: 0 };
    groups[rate].subtotal += lineTotal;
    groups[rate].vat += lineTotal * (rate / 100);
  }
  const vatTotal = Object.values(groups).reduce((s, g) => s + g.vat, 0);
  return { subtotal, groups, vatTotal, grandTotal: subtotal + vatTotal };
}

function fmtMoney(n, currencyOverride) {
  // currencyOverride laat toe om de valuta van een specifieke factuur/offerte
  // te gebruiken i.p.v. de standaardvaluta uit Instellingen.
  const currency = currencyOverride || (state.settings && state.settings.currency) || "EUR";
  try {
    return new Intl.NumberFormat("nl-NL", { style: "currency", currency }).format(n || 0);
  } catch (e) {
    // ongeldige/onbekende valutacode: toon toch een bedrag i.p.v. te crashen
    return currency + " " + (n || 0).toFixed(2);
  }
}

function fmtDate(iso) {
  if (!iso) return "";
  return new Date(iso + "T00:00:00").toLocaleDateString("nl-NL");
}

function hexToRgb(hex) {
  const clean = (hex || "#1d4ed8").replace("#", "");
  const full = clean.length === 3 ? clean.split("").map((c) => c + c).join("") : clean;
  const num = parseInt(full, 16) || 0x1d4ed8;
  return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
}

function emptyDraft(type = "factuur") {
  const date = todayISO();
  const language = "nl";
  const defaultTemplate = ((state.settings && state.settings.noteTemplates) || []).find(
    (t) => t.id === state.settings.defaultNoteTemplateId
  );
  return {
    id: crypto.randomUUID(),
    type,
    number: suggestNumber(type),
    reference: "",
    language,
    currency: (state.settings && state.settings.currency) || "EUR",
    status: "concept",
    date,
    paymentTermDays: state.settings.paymentTermDays,
    dueDate: addDays(date, state.settings.paymentTermDays),
    clientId: null,
    clientSnapshot: { name: "", addressLine: "", postalCode: "", city: "", country: "", vatNumber: "", email: "" },
    items: [{ dateFrom: "", dateTo: "", desc: "", qty: 1, unit: "", unitKey: "", price: 0, vatRate: DEFAULT_VAT_RATES[0] }],
    notes: defaultTemplate ? defaultTemplate[language] || "" : "",
  };
}

// ==========================================================================
// TOAST
// ==========================================================================
let toastTimer = null;
function toast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), 3000);
}

// ==========================================================================
// KLEINE GEDEELDE HELPERS
// ==========================================================================

// Vraagt een getal via prompt() en valideert het, i.p.v. dat een ongeldige
// invoer (of een typfout) stilzwijgend NaN of 0 wordt. Geeft null terug bij
// annuleren of ongeldige invoer (met een toast-uitleg in dat laatste geval) -
// de aanroeper laat de bestaande waarde dan gewoon ongewijzigd.
function promptNumber(message, currentValue, { min = -Infinity, max = Infinity } = {}) {
  const input = prompt(message, currentValue);
  if (input === null) return null; // geannuleerd
  const n = Number(String(input).trim().replace(",", "."));
  if (!Number.isFinite(n) || n < min || n > max) {
    toast("Ongeldige invoer: vul een getal in" + (Number.isFinite(min) || Number.isFinite(max) ? ` tussen ${min} en ${max}` : "") + ".");
    return null;
  }
  return n;
}

// Zelfde idee maar voor een 3-letterige ISO-valutacode.
function promptCurrencyCode(currentValue) {
  const input = prompt("ISO-valutacode (3 letters), bv. SEK, NOK, JPY:", currentValue);
  if (input === null) return null;
  const code = input.trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(code)) {
    toast("Ongeldige valutacode: gebruik precies 3 letters, bv. SEK.");
    return null;
  }
  return code;
}

// Zet een knop tijdelijk uit tijdens een async actie, zodat een dubbele klik
// niet leidt tot dubbele cloud-writes of dubbele PDF-downloads.
async function withBusy(btn, fn) {
  if (!btn || btn.disabled) return;
  btn.disabled = true;
  try {
    await fn();
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Genereert unieke id's voor formuliervelden, zodat elk <label> een geldig
// for-attribuut kan krijgen (nodig voor schermlezers).
let fieldIdCounter = 0;
function nextFieldId(prefix) {
  fieldIdCounter += 1;
  return `${prefix}-${fieldIdCounter}`;
}

// Maakt een modal-backdrop/modal-paar toegankelijk: dialog-semantiek,
// focus op het eerste veld, en Escape sluit de modal (zonder de listener
// te laten "hangen" nadat de modal via een andere weg is gesloten).
// Geeft een `close()` functie terug die overal waar de modal normaal
// gesproken sluit (Opslaan-knop, Annuleren-knop, klik buiten de modal)
// aangeroepen moet worden in plaats van rechtstreeks `backdrop.remove()`.
function attachModalA11y(backdrop, modal) {
  modal.setAttribute("role", "dialog");
  modal.setAttribute("aria-modal", "true");

  const previouslyFocused = document.activeElement;

  function onKeydown(e) {
    if (e.key === "Escape") close();
  }
  document.addEventListener("keydown", onKeydown);

  function close() {
    document.removeEventListener("keydown", onKeydown);
    backdrop.remove();
    if (previouslyFocused && typeof previouslyFocused.focus === "function") {
      previouslyFocused.focus();
    }
  }

  setTimeout(() => {
    const first = modal.querySelector("input, textarea, select, button");
    if (first) first.focus();
  }, 0);

  return close;
}

// ==========================================================================
// NAVIGATION
// ==========================================================================
function setupNav() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const view = btn.dataset.view;
      if (view === state.view) return;
      if (state.view === "settings" && hasUnsavedSettings()) {
        const ok = confirm("Je hebt niet-opgeslagen wijzigingen in Instellingen. Toch weg navigeren en wijzigingen verliezen?");
        if (!ok) return;
      }
      if (view === "new" && (!state.draft || state.draft.__saved)) {
        state.draft = emptyDraft();
      }
      state.view = view;
      renderView();
    });
  });
  window.addEventListener("beforeunload", (e) => {
    if (state.view === "settings" && hasUnsavedSettings()) {
      e.preventDefault();
      e.returnValue = "";
    }
  });
}

function updateActiveTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.view === state.view);
  });
}

function renderView() {
  updateActiveTabs();
  const root = document.getElementById("view-root");
  root.innerHTML = "";
  if (state.view === "new") root.appendChild(renderNewView());
  else if (state.view === "archive") root.appendChild(renderArchiveView());
  else if (state.view === "clients") root.appendChild(renderClientsView());
  else if (state.view === "settings") root.appendChild(renderSettingsView());
}

// ==========================================================================
// VIEW: NIEUWE FACTUUR / OFFERTE
// ==========================================================================
function renderNewView() {
  if (!state.draft) state.draft = emptyDraft();
  const draft = state.draft;
  const wrap = document.createElement("div");

  const title = document.createElement("div");
  title.className = "section-title";
  title.textContent = draft.__editing ? "Bewerken" : "Nieuwe factuur / offerte";
  wrap.appendChild(title);

  // Eerste-gebruik nudge: zonder bedrijfsnaam/adres mist de PDF de
  // belangrijkste bedrijfsgegevens - wijs de gebruiker daar hier al op,
  // i.p.v. dat dit pas opvalt bij het bekijken van de gedownloade PDF.
  if (!state.settings.companyName || !state.settings.companyName.trim()) {
    const notice = document.createElement("div");
    notice.className = "notice-card";
    notice.innerHTML = `Je hebt nog geen bedrijfsgegevens ingevuld. Vul deze eerst in bij <strong>Instellingen</strong>, anders ontbreken ze op je PDF.`;
    const noticeBtn = document.createElement("button");
    noticeBtn.className = "btn btn-secondary btn-sm";
    noticeBtn.textContent = "Naar Instellingen";
    noticeBtn.addEventListener("click", () => {
      state.view = "settings";
      renderView();
    });
    notice.appendChild(document.createElement("br"));
    notice.appendChild(noticeBtn);
    wrap.appendChild(notice);
  }

  const typeToggle = document.createElement("div");
  typeToggle.className = "type-toggle";
  ["factuur", "offerte"].forEach((t) => {
    const b = document.createElement("button");
    b.textContent = t === "factuur" ? "Factuur" : "Offerte";
    b.className = "type-" + t + (draft.type === t ? " active" : "");
    b.addEventListener("click", () => {
      if (draft.type === t) return;
      draft.type = t;
      if (!draft.__numberManuallyEdited) draft.number = suggestNumber(t);
      if (!STATUS_OPTIONS[t].includes(draft.status)) draft.status = "concept";
      renderView();
    });
    typeToggle.appendChild(b);
  });
  wrap.appendChild(typeToggle);

  const card = document.createElement("div");
  card.className = "card";

  const clientField = document.createElement("div");
  clientField.className = "field";
  const clientSelectId = nextFieldId("field");
  clientField.innerHTML = `<label for="${clientSelectId}">Klant</label>`;
  const clientSelect = document.createElement("select");
  clientSelect.id = clientSelectId;
  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "-- Kies een klant of vul handmatig in --";
  clientSelect.appendChild(optNone);
  state.clients.forEach((c) => {
    const o = document.createElement("option");
    o.value = c.id;
    o.textContent = c.name;
    if (draft.clientId === c.id) o.selected = true;
    clientSelect.appendChild(o);
  });
  clientSelect.addEventListener("change", () => {
    const c = state.clients.find((c) => c.id === clientSelect.value);
    draft.clientId = c ? c.id : null;
    draft.clientSnapshot = c
      ? {
          name: c.name,
          addressLine: c.addressLine || "",
          postalCode: c.postalCode || "",
          city: c.city || "",
          country: c.country || "",
          vatNumber: c.vatNumber,
          email: c.email,
        }
      : { name: "", addressLine: "", postalCode: "", city: "", country: "", vatNumber: "", email: "" };
    renderView();
  });
  clientField.appendChild(clientSelect);
  card.appendChild(clientField);

  const row1 = document.createElement("div");
  row1.className = "row";
  row1.appendChild(makeTextField("Klantnaam", draft.clientSnapshot.name, (v) => (draft.clientSnapshot.name = v)));
  row1.appendChild(makeTextField("E-mail", draft.clientSnapshot.email, (v) => (draft.clientSnapshot.email = v)));
  card.appendChild(row1);

  const row2 = document.createElement("div");
  row2.className = "row";
  row2.appendChild(
    makeTextField("Straat en huisnummer", draft.clientSnapshot.addressLine, (v) => (draft.clientSnapshot.addressLine = v))
  );
  row2.appendChild(makeTextField("Btw-nummer klant", draft.clientSnapshot.vatNumber, (v) => (draft.clientSnapshot.vatNumber = v)));
  card.appendChild(row2);

  const row2b = document.createElement("div");
  row2b.className = "row";
  row2b.appendChild(
    makeTextField("Postcode", draft.clientSnapshot.postalCode, (v) => (draft.clientSnapshot.postalCode = v))
  );
  row2b.appendChild(makeTextField("Plaats", draft.clientSnapshot.city, (v) => (draft.clientSnapshot.city = v)));
  row2b.appendChild(makeTextField("Land", draft.clientSnapshot.country, (v) => (draft.clientSnapshot.country = v)));
  card.appendChild(row2b);

  const saveClientBtn = document.createElement("button");
  saveClientBtn.className = "btn btn-secondary btn-sm";
  saveClientBtn.textContent = draft.clientId ? "Wijzigingen opslaan als klant" : "Opslaan als nieuwe klant";
  saveClientBtn.addEventListener("click", () =>
    withBusy(saveClientBtn, async () => {
      if (!draft.clientSnapshot.name.trim()) {
        toast("Vul minimaal een klantnaam in.");
        return;
      }
      const nameTaken = state.clients.some(
        (c) => c.id !== draft.clientId && c.name.trim().toLowerCase() === draft.clientSnapshot.name.trim().toLowerCase()
      );
      if (nameTaken) {
        const ok = confirm(`Er bestaat al een klant met de naam "${draft.clientSnapshot.name}". Toch opslaan als (extra) nieuwe klant?`);
        if (!ok) return;
      }
      try {
        if (draft.clientId && state.clients.some((c) => c.id === draft.clientId)) {
          const idx = state.clients.findIndex((c) => c.id === draft.clientId);
          state.clients[idx] = { id: draft.clientId, ...draft.clientSnapshot };
        } else {
          const newClient = { id: crypto.randomUUID(), ...draft.clientSnapshot };
          state.clients.push(newClient);
          draft.clientId = newClient.id;
        }
        await saveClients();
        toast("Klant \"" + draft.clientSnapshot.name + "\" opgeslagen.");
        renderView();
      } catch (err) {
        toast("Opslaan mislukt: " + err.message);
      }
    })
  );
  card.appendChild(saveClientBtn);

  wrap.appendChild(card);

  const card2 = document.createElement("div");
  card2.className = "card";
  const row3 = document.createElement("div");
  row3.className = "row";
  row3.appendChild(
    makeTextField(draft.type === "factuur" ? "Factuurnummer" : "Offertenummer", draft.number, (v) => {
      draft.number = v;
      draft.__numberManuallyEdited = true;
    })
  );
  row3.appendChild(
    makeDateField("Datum", draft.date, (v) => {
      draft.date = v;
      if (!draft.__dueDateManuallyEdited) {
        draft.dueDate = addDays(v, draft.paymentTermDays);
      }
      renderView();
    })
  );
  card2.appendChild(row3);

  const row4 = document.createElement("div");
  row4.className = "row";
  row4.appendChild(
    makeSelectField(
      "Betalingstermijn",
      String(draft.paymentTermDays),
      [
        ["0", "Direct"],
        ["7", "7 dagen"],
        ["14", "14 dagen"],
        ["30", "30 dagen"],
        ["custom", "Anders..."],
      ],
      (v) => {
        if (v === "custom") {
          const n = promptNumber("Aantal dagen betalingstermijn:", draft.paymentTermDays, { min: 0, max: 365 });
          if (n === null) return;
          draft.paymentTermDays = n;
        } else {
          draft.paymentTermDays = Number(v);
        }
        if (!draft.__dueDateManuallyEdited) {
          draft.dueDate = addDays(draft.date, draft.paymentTermDays);
        }
        renderView();
      }
    )
  );
  row4.appendChild(
    makeDateField(draft.type === "factuur" ? "Vervaldatum" : "Geldig tot", draft.dueDate, (v) => {
      draft.dueDate = v;
      draft.__dueDateManuallyEdited = true;
    })
  );
  row4.appendChild(
    makeSelectField(
      "Status",
      draft.status,
      STATUS_OPTIONS[draft.type].map((s) => [s, STATUS_LABELS[s]]),
      (v) => (draft.status = v)
    )
  );
  card2.appendChild(row4);

  const row5 = document.createElement("div");
  row5.className = "row";
  row5.appendChild(makeTextField("Referentie (optioneel)", draft.reference, (v) => (draft.reference = v)));
  row5.appendChild(
    makeSelectField(
      "Taal van het document",
      draft.language || "nl",
      [
        ["nl", "Nederlands"],
        ["en", "English"],
      ],
      (v) => {
        const oldLang = draft.language || "nl";
        // Als de huidige notities exact overeenkomen met een bekend sjabloon
        // in de oude taal, meteen meeswitchen naar dezelfde tekst in de
        // nieuwe taal. Handmatig aangepaste notities blijven ongemoeid.
        const templates = (state.settings && state.settings.noteTemplates) || [];
        const matchedTemplate = templates.find(
          (tpl) => (draft.notes || "").trim() !== "" && (tpl[oldLang] || "") === draft.notes
        );
        draft.language = v;
        if (matchedTemplate) {
          draft.notes = matchedTemplate[v] || matchedTemplate.nl || "";
        }
        renderView();
      }
    )
  );
  row5.appendChild(
    makeSelectField(
      "Valuta",
      draft.currency || (state.settings && state.settings.currency) || "EUR",
      [
        ["EUR", "Euro (€)"],
        ["USD", "US dollar ($)"],
        ["GBP", "Brits pond (£)"],
        ["CHF", "Zwitserse frank (CHF)"],
        ["custom", "Andere (ISO-code)..."],
      ],
      (v) => {
        if (v === "custom") {
          const code = promptCurrencyCode(draft.currency || "EUR");
          if (code === null) return;
          draft.currency = code;
        } else {
          draft.currency = v;
        }
        // Volledige re-render zodat regeltotalen en de totalen-box meteen
        // de nieuwe valuta tonen.
        renderView();
      }
    )
  );
  card2.appendChild(row5);
  wrap.appendChild(card2);

  const card3 = document.createElement("div");
  card3.className = "card";
  card3.innerHTML = `<h2>Regels</h2>`;
  card3.appendChild(renderItemsTable(draft));

  const addBtn = document.createElement("button");
  addBtn.className = "btn btn-secondary btn-sm";
  addBtn.textContent = "+ Regel toevoegen";
  addBtn.addEventListener("click", () => {
    draft.items.push({ dateFrom: "", dateTo: "", desc: "", qty: 1, unit: "", unitKey: "", price: 0, vatRate: DEFAULT_VAT_RATES[0] });
    renderView();
  });
  card3.appendChild(addBtn);

  const totals = calcTotals(draft.items);
  const totalsBox = document.createElement("div");
  totalsBox.className = "totals-box";
  let totalsHtml = `<div class="totals-row"><span>Subtotaal excl. btw</span><span>${fmtMoney(totals.subtotal, draft.currency)}</span></div>`;
  Object.keys(totals.groups)
    .sort()
    .forEach((rate) => {
      totalsHtml += `<div class="totals-row"><span>Btw ${rate}%</span><span>${fmtMoney(totals.groups[rate].vat, draft.currency)}</span></div>`;
    });
  totalsHtml += `<div class="totals-row grand"><span>Totaal incl. btw</span><span>${fmtMoney(totals.grandTotal, draft.currency)}</span></div>`;
  totalsBox.innerHTML = totalsHtml;
  card3.appendChild(totalsBox);
  wrap.appendChild(card3);

  const card4 = document.createElement("div");
  card4.className = "card";
  card4.innerHTML = `<h2>Notities</h2>`;
  const noteTemplates = (state.settings && state.settings.noteTemplates) || [];
  if (noteTemplates.length) {
    card4.appendChild(
      makeSelectField(
        "Notitiesjabloon invoegen",
        "",
        [["", "— Kies een sjabloon —"], ...noteTemplates.map((tpl) => [tpl.id, tpl.name])],
        (v) => {
          if (!v) return;
          const tpl = noteTemplates.find((x) => x.id === v);
          if (tpl) {
            draft.notes = tpl[draft.language || "nl"] || tpl.nl || "";
            renderView();
          }
        }
      )
    );
  }
  card4.appendChild(makeTextAreaField("Notities (optioneel, komt op de " + draft.type + ")", draft.notes, (v) => (draft.notes = v)));
  wrap.appendChild(card4);

  const actions = document.createElement("div");
  actions.className = "row";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = draft.__editing ? "Wijzigingen opslaan" : "Opslaan in archief";
  saveBtn.addEventListener("click", () => withBusy(saveBtn, () => saveDraft(draft)));
  actions.appendChild(saveBtn);

  const pdfBtn = document.createElement("button");
  pdfBtn.className = "btn btn-secondary";
  pdfBtn.textContent = "Download PDF";
  pdfBtn.addEventListener("click", () => withBusy(pdfBtn, () => generatePDF(draft)));
  actions.appendChild(pdfBtn);

  if (draft.__editing) {
    const cancelBtn = document.createElement("button");
    cancelBtn.className = "btn btn-ghost";
    cancelBtn.textContent = "Annuleren";
    cancelBtn.addEventListener("click", () => {
      state.draft = emptyDraft();
      state.view = "archive";
      renderView();
    });
    actions.appendChild(cancelBtn);
  }

  wrap.appendChild(actions);
  return wrap;
}

function renderItemsTable(draft) {
  const table = document.createElement("table");
  table.className = "items-table";
  table.innerHTML = `<thead><tr>
    <th class="col-date">Datum</th><th>Omschrijving</th><th class="col-qty">Aantal</th><th class="col-unit">Eenheid</th><th class="col-price">Prijs</th>
    <th class="col-vat">Btw</th><th class="col-total">Totaal</th><th class="col-remove"></th>
  </tr></thead>`;
  const tbody = document.createElement("tbody");
  draft.items.forEach((item, idx) => {
    const tr = document.createElement("tr");

    // Eén datumkiezer (kalender via de native date-input), met een kleine
    // "+ periode toevoegen"-link eronder die pas een tweede ("tot") kiezer
    // toont zodra je die nodig hebt - dat houdt de tabel rustiger voor de
    // meeste regels (die maar één datum of helemaal geen datum hebben) en is
    // tegelijk duidelijk zichtbaar dat een periode mogelijk is.
    const tdDate = document.createElement("td");
    tdDate.dataset.label = "Datum (optioneel)";
    tdDate.className = "col-date";

    const dateFromInput = document.createElement("input");
    dateFromInput.type = "date";
    dateFromInput.value = item.dateFrom || "";
    dateFromInput.addEventListener("input", () => (item.dateFrom = dateFromInput.value));
    tdDate.appendChild(dateFromInput);

    const dateToWrap = document.createElement("div");
    dateToWrap.style.marginTop = "4px";
    const dateToInput = document.createElement("input");
    dateToInput.type = "date";
    dateToInput.value = item.dateTo || "";
    dateToInput.title = "Tot (periode)";
    dateToInput.addEventListener("input", () => (item.dateTo = dateToInput.value));
    dateToWrap.appendChild(dateToInput);

    const periodToggle = document.createElement("button");
    periodToggle.type = "button";
    periodToggle.className = "link-btn";
    periodToggle.textContent = "+ periode (tot datum)";
    periodToggle.addEventListener("click", () => {
      dateToWrap.classList.remove("hidden");
      periodToggle.classList.add("hidden");
      dateToInput.focus();
    });

    if (item.dateTo) {
      periodToggle.classList.add("hidden");
    } else {
      dateToWrap.classList.add("hidden");
    }

    tdDate.appendChild(periodToggle);
    tdDate.appendChild(dateToWrap);
    tr.appendChild(tdDate);

    const tdDesc = document.createElement("td");
    tdDesc.dataset.label = "Omschrijving";
    tdDesc.className = "item-desc";
    const descInput = document.createElement("input");
    descInput.value = item.desc;
    descInput.placeholder = "Bijv. Consult / Product X";
    descInput.addEventListener("input", () => (item.desc = descInput.value));
    tdDesc.appendChild(descInput);
    tr.appendChild(tdDesc);

    const tdQty = document.createElement("td");
    tdQty.dataset.label = "Aantal";
    tdQty.className = "col-qty";
    const qtyInput = document.createElement("input");
    qtyInput.type = "number";
    qtyInput.className = "item-num";
    qtyInput.step = "any";
    qtyInput.value = item.qty;
    qtyInput.addEventListener("input", () => {
      item.qty = qtyInput.value;
      updateLineTotal(idx);
    });
    tdQty.appendChild(qtyInput);
    tr.appendChild(tdQty);

    const tdUnit = document.createElement("td");
    tdUnit.dataset.label = "Eenheid";
    tdUnit.className = "col-unit";
    const unitSelect = document.createElement("select");
    const unitOptions = [
      ["", "Geen"],
      ["dagen", "Dagen"],
      ["km", "Kilometers"],
      ["eenheden", "Eenheden"],
      ["custom", "Zelf in te vullen..."],
    ];
    unitOptions.forEach(([val, text]) => {
      const o = document.createElement("option");
      o.value = val;
      o.textContent = text;
      unitSelect.appendChild(o);
    });
    // Een bekend preset -> die staat geselecteerd. "custom" moet expliciet uit
    // item.unitKey blijven komen (niet afgeleid van of item.unit al tekst
    // bevat) - anders klapt de keuze na elke render meteen terug naar "Geen"
    // zodra het tekstveld nog leeg is, en lijkt de knop niet te werken.
    // Voor oude data zonder unitKey: als er toch al eigen tekst in item.unit
    // staat, ook als "Zelf in te vullen" tonen zodat niets verloren gaat.
    const isKnownPreset = UNIT_PRESETS.some((u) => u.key === item.unitKey);
    const isCustom = item.unitKey === "custom" || (!isKnownPreset && !!item.unit);
    unitSelect.value = isKnownPreset ? item.unitKey : isCustom ? "custom" : "";
    unitSelect.addEventListener("change", () => {
      const v = unitSelect.value;
      if (v === "custom") {
        item.unitKey = "custom";
        // Tekst blijft leeg totdat de gebruiker hem hieronder intikt - geen
        // prompt() meer, want dan was niet zichtbaar wat je had ingevuld.
        if (isKnownPreset || !item.unit) item.unit = "";
      } else if (v === "") {
        item.unitKey = "";
        item.unit = "";
      } else {
        const preset = UNIT_PRESETS.find((u) => u.key === v);
        item.unitKey = v;
        item.unit = preset ? preset.nl : "";
      }
      renderView();
    });
    tdUnit.appendChild(unitSelect);

    // Zodra "Zelf in te vullen" gekozen is, verschijnt hier gewoon een normaal
    // tekstveld met de ingevulde tekst zichtbaar erin - i.p.v. een prompt().
    if (unitSelect.value === "custom") {
      const customUnitInput = document.createElement("input");
      customUnitInput.type = "text";
      customUnitInput.value = item.unit || "";
      customUnitInput.placeholder = "eigen eenheid...";
      customUnitInput.style.marginTop = "6px";
      customUnitInput.addEventListener("input", () => (item.unit = customUnitInput.value));
      tdUnit.appendChild(customUnitInput);
    }
    tr.appendChild(tdUnit);

    const tdPrice = document.createElement("td");
    tdPrice.dataset.label = "Prijs (excl. btw)";
    tdPrice.className = "col-price";
    const priceInput = document.createElement("input");
    priceInput.type = "number";
    priceInput.className = "item-num";
    priceInput.step = "0.01";
    priceInput.value = item.price;
    priceInput.addEventListener("input", () => {
      item.price = priceInput.value;
      updateLineTotal(idx);
    });
    tdPrice.appendChild(priceInput);
    tr.appendChild(tdPrice);

    const tdVat = document.createElement("td");
    tdVat.dataset.label = "Btw%";
    tdVat.className = "col-vat";
    const vatSelect = document.createElement("select");
    const vatOptions = [...DEFAULT_VAT_RATES];
    if (!vatOptions.includes(Number(item.vatRate))) vatOptions.push(Number(item.vatRate));
    vatOptions.forEach((r) => {
      const o = document.createElement("option");
      o.value = r;
      o.textContent = r + "%";
      if (Number(item.vatRate) === r) o.selected = true;
      vatSelect.appendChild(o);
    });
    const customOpt = document.createElement("option");
    customOpt.value = "custom";
    customOpt.textContent = "Anders...";
    vatSelect.appendChild(customOpt);
    vatSelect.addEventListener("change", () => {
      if (vatSelect.value === "custom") {
        const n = promptNumber("Aangepast btw-percentage:", item.vatRate, { min: 0, max: 100 });
        if (n !== null) item.vatRate = n;
        renderView();
      } else {
        item.vatRate = Number(vatSelect.value);
        updateLineTotal(idx);
      }
    });
    tdVat.appendChild(vatSelect);
    tr.appendChild(tdVat);

    const tdTotal = document.createElement("td");
    tdTotal.dataset.label = "Totaal";
    tdTotal.className = "col-total";
    tdTotal.id = "line-total-" + idx;
    tdTotal.textContent = fmtMoney((Number(item.qty) || 0) * (Number(item.price) || 0), draft.currency);
    tr.appendChild(tdTotal);

    const tdRemove = document.createElement("td");
    tdRemove.className = "col-remove";
    if (draft.items.length > 1) {
      const rm = document.createElement("button");
      rm.className = "remove-line";
      rm.innerHTML = "&times;";
      rm.title = "Regel verwijderen";
      rm.addEventListener("click", () => {
        draft.items.splice(idx, 1);
        renderView();
      });
      tdRemove.appendChild(rm);
    }
    tr.appendChild(tdRemove);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  // Scrollbare wrapper op desktop, zodat de tabel kolommen niet hoeft dicht
  // te knijpen als het venster wat smaller is - mobiel (kaartweergave) blijft
  // ongewijzigd, zie de @media (max-width:700px) regels in style.css.
  const wrap = document.createElement("div");
  wrap.className = "items-table-wrap";
  wrap.appendChild(table);
  return wrap;
}

function updateLineTotal(idx) {
  const draft = state.draft;
  const item = draft.items[idx];
  const el = document.getElementById("line-total-" + idx);
  if (el) el.textContent = fmtMoney((Number(item.qty) || 0) * (Number(item.price) || 0), draft.currency);
  const totals = calcTotals(draft.items);
  const totalsBox = document.querySelector(".totals-box");
  if (totalsBox) {
    let html = `<div class="totals-row"><span>Subtotaal excl. btw</span><span>${fmtMoney(totals.subtotal, draft.currency)}</span></div>`;
    Object.keys(totals.groups)
      .sort()
      .forEach((rate) => {
        html += `<div class="totals-row"><span>Btw ${rate}%</span><span>${fmtMoney(totals.groups[rate].vat, draft.currency)}</span></div>`;
      });
    html += `<div class="totals-row grand"><span>Totaal incl. btw</span><span>${fmtMoney(totals.grandTotal, draft.currency)}</span></div>`;
    totalsBox.innerHTML = html;
  }
}

function makeTextField(label, value, onChange, opts = {}) {
  const div = document.createElement("div");
  div.className = "field";
  const id = nextFieldId("field");
  const l = document.createElement("label");
  l.textContent = label;
  l.setAttribute("for", id);
  const input = document.createElement("input");
  input.type = "text";
  input.id = id;
  input.value = value || "";
  if (opts.readOnly) input.readOnly = true;
  input.addEventListener("input", () => onChange(input.value));
  div.appendChild(l);
  div.appendChild(input);
  return div;
}

function makeTextAreaField(label, value, onChange) {
  const div = document.createElement("div");
  div.className = "field";
  const id = nextFieldId("field");
  const l = document.createElement("label");
  l.textContent = label;
  l.setAttribute("for", id);
  const ta = document.createElement("textarea");
  ta.id = id;
  ta.value = value || "";
  ta.addEventListener("input", () => onChange(ta.value));
  div.appendChild(l);
  div.appendChild(ta);
  return div;
}

function makeDateField(label, value, onChange) {
  const div = document.createElement("div");
  div.className = "field";
  const id = nextFieldId("field");
  const l = document.createElement("label");
  l.textContent = label;
  l.setAttribute("for", id);
  const input = document.createElement("input");
  input.type = "date";
  input.id = id;
  input.value = value || "";
  input.addEventListener("change", () => onChange(input.value));
  div.appendChild(l);
  div.appendChild(input);
  return div;
}

function makeSelectField(label, value, options, onChange) {
  const div = document.createElement("div");
  div.className = "field";
  const id = nextFieldId("field");
  const l = document.createElement("label");
  l.textContent = label;
  l.setAttribute("for", id);
  const select = document.createElement("select");
  select.id = id;
  options.forEach(([val, text]) => {
    const o = document.createElement("option");
    o.value = val;
    o.textContent = text;
    if (String(value) === String(val)) o.selected = true;
    select.appendChild(o);
  });
  select.addEventListener("change", () => onChange(select.value));
  div.appendChild(l);
  div.appendChild(select);
  return div;
}

async function saveDraft(draft) {
  if (!draft.clientSnapshot.name.trim()) {
    toast("Vul minimaal een klantnaam in.");
    return;
  }

  const duplicateNumber = state.invoices.some(
    (i) => i.id !== draft.id && i.type === draft.type && i.number.trim().toLowerCase() === draft.number.trim().toLowerCase()
  );
  if (duplicateNumber) {
    const ok = confirm(
      `Er bestaat al ${draft.type === "factuur" ? "een factuur" : "een offerte"} met nummer "${draft.number}". Toch opslaan met dit nummer (dit geeft dubbele nummers in het archief)?`
    );
    if (!ok) return;
  }

  const hasEmptyRow = draft.items.some((it) => !it.desc || !it.desc.trim());
  if (hasEmptyRow) {
    const ok = confirm("Eén of meer regels hebben geen omschrijving. Toch opslaan?");
    if (!ok) return;
  }

  const existingIdx = state.invoices.findIndex((i) => i.id === draft.id);
  const totals = calcTotals(draft.items);
  const toSave = {
    ...draft,
    total: totals.grandTotal,
    updatedAt: new Date().toISOString(),
    createdAt: draft.createdAt || new Date().toISOString(),
  };
  delete toSave.__editing;
  delete toSave.__saved;
  delete toSave.__numberManuallyEdited;
  delete toSave.__dueDateManuallyEdited;

  bumpNumberCounterIfMatches(draft.type, draft.number);

  if (existingIdx >= 0) {
    state.invoices[existingIdx] = toSave;
  } else {
    state.invoices.push(toSave);
  }

  try {
    await Promise.all([saveInvoices(), saveSettings()]);
    toast((draft.type === "factuur" ? "Factuur" : "Offerte") + " " + draft.number + " opgeslagen.");
    state.draft = emptyDraft();
    state.view = "archive";
    renderView();
  } catch (err) {
    console.error(err);
    toast("Opslaan mislukt: " + err.message);
  }
}

// ==========================================================================
// VIEW: ARCHIEF
// ==========================================================================
function renderArchiveView() {
  const wrap = document.createElement("div");
  const title = document.createElement("div");
  title.className = "section-title";
  title.textContent = "Archief";
  wrap.appendChild(title);

  const filters = document.createElement("div");
  filters.className = "filters";

  const typeSelect = document.createElement("select");
  [["alle", "Alle types"], ["factuur", "Facturen"], ["offerte", "Offertes"]].forEach(([v, t]) => {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = t;
    if (state.archiveFilter.type === v) o.selected = true;
    typeSelect.appendChild(o);
  });
  typeSelect.addEventListener("change", () => {
    state.archiveFilter.type = typeSelect.value;
    renderView();
  });
  filters.appendChild(typeSelect);

  const statusSelect = document.createElement("select");
  const allStatuses = ["alle", ...new Set([...STATUS_OPTIONS.factuur, ...STATUS_OPTIONS.offerte])];
  allStatuses.forEach((s) => {
    const o = document.createElement("option");
    o.value = s;
    o.textContent = s === "alle" ? "Alle statussen" : STATUS_LABELS[s];
    if (state.archiveFilter.status === s) o.selected = true;
    statusSelect.appendChild(o);
  });
  statusSelect.addEventListener("change", () => {
    state.archiveFilter.status = statusSelect.value;
    renderView();
  });
  filters.appendChild(statusSelect);

  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.placeholder = "Zoek op nummer of klantnaam...";
  searchInput.value = state.archiveFilter.q;
  searchInput.addEventListener("input", () => {
    state.archiveFilter.q = searchInput.value;
    renderArchiveList(list);
  });
  filters.appendChild(searchInput);

  wrap.appendChild(filters);

  const list = document.createElement("div");
  list.className = "archive-list";
  wrap.appendChild(list);
  renderArchiveList(list);

  return wrap;
}

function getFilteredInvoices() {
  const f = state.archiveFilter;
  return state.invoices
    .filter((inv) => f.type === "alle" || inv.type === f.type)
    .filter((inv) => f.status === "alle" || inv.status === f.status)
    .filter((inv) => {
      if (!f.q.trim()) return true;
      const q = f.q.trim().toLowerCase();
      return inv.number.toLowerCase().includes(q) || (inv.clientSnapshot.name || "").toLowerCase().includes(q);
    })
    .sort((a, b) => (b.date > a.date ? 1 : -1));
}

function renderArchiveList(list) {
  const items = getFilteredInvoices();
  list.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    if (!state.invoices.length) {
      empty.textContent = "Nog geen facturen of offertes. Maak er een aan bij \"Nieuw\".";
    } else {
      const f = state.archiveFilter;
      const filterActive = f.type !== "alle" || f.status !== "alle" || f.q.trim();
      empty.textContent = filterActive
        ? "Geen facturen of offertes gevonden voor dit filter/deze zoekopdracht."
        : "Geen facturen of offertes gevonden.";
    }
    list.appendChild(empty);
    return;
  }
  items.forEach((inv) => {
    const row = document.createElement("div");
    row.className = "archive-row";
    row.innerHTML = `
      <div class="col-main">
        <span class="badge badge-type-${inv.type}">${inv.type === "factuur" ? "Factuur" : "Offerte"}</span>
        <span class="badge badge-${inv.status}">${STATUS_LABELS[inv.status]}</span>
        <div class="num">${inv.number} — ${inv.clientSnapshot.name || "(geen klantnaam)"}</div>
        <div class="sub">${fmtDate(inv.date)} · ${inv.type === "factuur" ? "vervaldatum" : "geldig tot"} ${fmtDate(inv.dueDate)}</div>
      </div>
      <div class="amount">${fmtMoney(inv.total, inv.currency)}</div>
    `;
    const statusSelect = document.createElement("select");
    statusSelect.className = "status-select";
    STATUS_OPTIONS[inv.type].forEach((s) => {
      const o = document.createElement("option");
      o.value = s;
      o.textContent = STATUS_LABELS[s];
      if (inv.status === s) o.selected = true;
      statusSelect.appendChild(o);
    });
    statusSelect.addEventListener("change", () =>
      withBusy(statusSelect, async () => {
        inv.status = statusSelect.value;
        inv.updatedAt = new Date().toISOString();
        try {
          await saveInvoices();
          toast("Status van " + inv.number + " gewijzigd naar " + STATUS_LABELS[inv.status] + ".");
          renderArchiveList(list);
        } catch (err) {
          toast("Opslaan mislukt: " + err.message);
        }
      })
    );
    row.appendChild(statusSelect);

    const actions = document.createElement("div");
    actions.className = "actions";

    const pdfBtn = document.createElement("button");
    pdfBtn.className = "btn btn-secondary btn-sm";
    pdfBtn.textContent = "PDF";
    pdfBtn.addEventListener("click", () => withBusy(pdfBtn, () => generatePDF(inv)));
    actions.appendChild(pdfBtn);

    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-ghost btn-sm";
    editBtn.textContent = "Bewerken";
    editBtn.addEventListener("click", () => {
      state.draft = { ...JSON.parse(JSON.stringify(inv)), __editing: true };
      state.view = "new";
      renderView();
    });
    actions.appendChild(editBtn);

    const dupBtn = document.createElement("button");
    dupBtn.className = "btn btn-ghost btn-sm";
    dupBtn.textContent = "Dupliceren";
    dupBtn.addEventListener("click", () => {
      const copy = JSON.parse(JSON.stringify(inv));
      delete copy.__editing;
      delete copy.__saved;
      delete copy.createdAt;
      delete copy.updatedAt;
      copy.id = crypto.randomUUID();
      copy.number = suggestNumber(copy.type);
      copy.status = "concept";
      copy.date = todayISO();
      copy.dueDate = addDays(copy.date, copy.paymentTermDays);
      state.draft = copy;
      state.view = "new";
      renderView();
      toast("Kopie gemaakt van " + inv.number + " — controleer en sla op.");
    });
    actions.appendChild(dupBtn);

    const delBtn = document.createElement("button");
    delBtn.className = "btn btn-danger btn-sm";
    delBtn.textContent = "Verwijderen";
    delBtn.addEventListener("click", () =>
      withBusy(delBtn, async () => {
        if (!confirm(`${inv.type === "factuur" ? "Factuur" : "Offerte"} ${inv.number} verwijderen? Dit kan niet ongedaan worden gemaakt.`)) return;
        state.invoices = state.invoices.filter((i) => i.id !== inv.id);
        try {
          await saveInvoices();
          toast(inv.number + " verwijderd.");
          renderArchiveList(list);
        } catch (err) {
          toast("Verwijderen mislukt: " + err.message);
          return;
        }
        const provider = activeProvider();
        if (provider && provider.deleteFile) {
          try {
            await provider.deleteFile(`${inv.number}.pdf`);
          } catch (err) {
            console.warn("PDF verwijderen uit cloudopslag mislukt", err);
            toast("Let op: de PDF van " + inv.number + " kon niet worden verwijderd uit " + provider.name + ".");
          }
        }
      })
    );
    actions.appendChild(delBtn);

    row.appendChild(actions);
    list.appendChild(row);
  });
}

// ==========================================================================
// VIEW: KLANTEN
// ==========================================================================
function renderClientsView() {
  const wrap = document.createElement("div");
  const title = document.createElement("div");
  title.className = "section-title";
  title.textContent = "Klanten";
  wrap.appendChild(title);

  const addBtn = document.createElement("button");
  addBtn.className = "btn btn-primary";
  addBtn.textContent = "+ Nieuwe klant";
  addBtn.style.marginBottom = "16px";
  addBtn.addEventListener("click", () => openClientModal(null));
  wrap.appendChild(addBtn);

  if (!state.clients.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "Nog geen klanten toegevoegd.";
    wrap.appendChild(empty);
    return wrap;
  }

  state.clients.forEach((c) => {
    const row = document.createElement("div");
    row.className = "list-row";
    row.innerHTML = `
      <div class="col-main">
        <div class="name">${c.name}</div>
        <div class="sub">${[c.addressLine, [c.postalCode, c.city].filter(Boolean).join(" "), c.country].filter(Boolean).join(", ")} ${c.vatNumber ? "· btw: " + c.vatNumber : ""}</div>
      </div>
    `;
    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-ghost btn-sm";
    editBtn.textContent = "Bewerken";
    editBtn.addEventListener("click", () => openClientModal(c));
    row.appendChild(editBtn);

    const delBtn = document.createElement("button");
    delBtn.className = "btn btn-danger btn-sm";
    delBtn.textContent = "Verwijderen";
    delBtn.addEventListener("click", () =>
      withBusy(delBtn, async () => {
        if (!confirm(`Klant "${c.name}" verwijderen?`)) return;
        state.clients = state.clients.filter((x) => x.id !== c.id);
        await saveClients();
        renderView();
      })
    );
    row.appendChild(delBtn);

    wrap.appendChild(row);
  });

  return wrap;
}

function openClientModal(client) {
  const isNew = !client;
  const draftClient = client
    ? { ...client }
    : { id: crypto.randomUUID(), name: "", addressLine: "", postalCode: "", city: "", country: "", vatNumber: "", email: "" };

  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.innerHTML = `<h2>${isNew ? "Nieuwe klant" : "Klant bewerken"}</h2>`;
  const closeModal = attachModalA11y(backdrop, modal);

  modal.appendChild(makeTextField("Naam", draftClient.name, (v) => (draftClient.name = v)));
  modal.appendChild(makeTextField("Straat en huisnummer", draftClient.addressLine, (v) => (draftClient.addressLine = v)));
  const clientAddrRow = document.createElement("div");
  clientAddrRow.className = "row";
  clientAddrRow.appendChild(makeTextField("Postcode", draftClient.postalCode, (v) => (draftClient.postalCode = v)));
  clientAddrRow.appendChild(makeTextField("Plaats", draftClient.city, (v) => (draftClient.city = v)));
  clientAddrRow.appendChild(makeTextField("Land", draftClient.country, (v) => (draftClient.country = v)));
  modal.appendChild(clientAddrRow);
  modal.appendChild(makeTextField("Btw-nummer", draftClient.vatNumber, (v) => (draftClient.vatNumber = v)));
  modal.appendChild(makeTextField("E-mail", draftClient.email, (v) => (draftClient.email = v)));

  const actions = document.createElement("div");
  actions.className = "row";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Opslaan";
  saveBtn.addEventListener("click", () =>
    withBusy(saveBtn, async () => {
      if (!draftClient.name.trim()) {
        toast("Vul een naam in.");
        return;
      }
      const nameTaken = state.clients.some(
        (c) => c.id !== draftClient.id && c.name.trim().toLowerCase() === draftClient.name.trim().toLowerCase()
      );
      if (nameTaken) {
        const ok = confirm(`Er bestaat al een klant met de naam "${draftClient.name}". Toch opslaan?`);
        if (!ok) return;
      }
      const idx = state.clients.findIndex((c) => c.id === draftClient.id);
      if (idx >= 0) state.clients[idx] = draftClient;
      else state.clients.push(draftClient);
      await saveClients();
      closeModal();
      renderView();
    })
  );
  actions.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-ghost";
  cancelBtn.textContent = "Annuleren";
  cancelBtn.addEventListener("click", () => closeModal());
  actions.appendChild(cancelBtn);

  modal.appendChild(actions);
  backdrop.appendChild(modal);
  backdrop.addEventListener("click", (e) => {
    if (e.target === backdrop) closeModal();
  });
  document.body.appendChild(backdrop);
}

function openNoteTemplateModal(template) {
  const isNew = !template;
  const draftTpl = template ? { ...template } : { id: crypto.randomUUID(), name: "", nl: "", en: "" };

  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.innerHTML = `<h2>${isNew ? "Nieuw notitiesjabloon" : "Sjabloon bewerken"}</h2>`;
  const closeModal = attachModalA11y(backdrop, modal);

  modal.appendChild(makeTextField("Naam van het sjabloon", draftTpl.name, (v) => (draftTpl.name = v)));
  modal.appendChild(makeTextAreaField("Tekst (Nederlands)", draftTpl.nl, (v) => (draftTpl.nl = v)));
  modal.appendChild(makeTextAreaField("Tekst (Engels)", draftTpl.en, (v) => (draftTpl.en = v)));

  const defaultField = document.createElement("div");
  defaultField.className = "field";
  const defaultLabel = document.createElement("label");
  defaultLabel.style.cssText = "display:flex;align-items:center;gap:8px;font-weight:600;color:var(--text);";
  const defaultCheckbox = document.createElement("input");
  defaultCheckbox.type = "checkbox";
  defaultCheckbox.checked = state.settings.defaultNoteTemplateId === draftTpl.id;
  defaultLabel.appendChild(defaultCheckbox);
  defaultLabel.appendChild(
    document.createTextNode("Gebruik als standaard (wordt automatisch ingevuld bij een nieuwe factuur/offerte)")
  );
  defaultField.appendChild(defaultLabel);
  modal.appendChild(defaultField);

  const actions = document.createElement("div");
  actions.className = "row";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Opslaan";
  saveBtn.addEventListener("click", () =>
    withBusy(saveBtn, async () => {
      if (!draftTpl.name.trim()) {
        toast("Vul een naam in voor het sjabloon.");
        return;
      }
      if (!state.settings.noteTemplates) state.settings.noteTemplates = [];
      const idx = state.settings.noteTemplates.findIndex((t) => t.id === draftTpl.id);
      if (idx >= 0) state.settings.noteTemplates[idx] = draftTpl;
      else state.settings.noteTemplates.push(draftTpl);
      if (defaultCheckbox.checked) {
        state.settings.defaultNoteTemplateId = draftTpl.id;
      } else if (state.settings.defaultNoteTemplateId === draftTpl.id) {
        state.settings.defaultNoteTemplateId = null;
      }
      await saveSettings();
      closeModal();
      renderView();
    })
  );
  actions.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-ghost";
  cancelBtn.textContent = "Annuleren";
  cancelBtn.addEventListener("click", () => closeModal());
  actions.appendChild(cancelBtn);

  modal.appendChild(actions);
  backdrop.appendChild(modal);
  backdrop.addEventListener("click", (e) => {
    if (e.target === backdrop) closeModal();
  });
  document.body.appendChild(backdrop);
}

// ==========================================================================
// VIEW: INSTELLINGEN
// ==========================================================================
function renderSettingsView() {
  const wrap = document.createElement("div");
  const title = document.createElement("div");
  title.className = "section-title";
  title.textContent = "Instellingen";
  wrap.appendChild(title);

  const s = state.settings;
  const card = document.createElement("div");
  card.className = "card";
  card.innerHTML = `<h2>Bedrijfsgegevens</h2>`;

  if (s.logoDataUrl) {
    const img = document.createElement("img");
    img.className = "logo-preview";
    img.src = s.logoDataUrl;
    card.appendChild(img);
  }
  const logoField = document.createElement("div");
  logoField.className = "field";
  const logoInputId = nextFieldId("field");
  logoField.innerHTML = `<label for="${logoInputId}">Logo</label>`;
  const logoInput = document.createElement("input");
  logoInput.id = logoInputId;
  logoInput.type = "file";
  logoInput.accept = "image/*";
  logoInput.addEventListener("change", async () => {
    const file = logoInput.files[0];
    if (!file) return;
    s.logoDataUrl = await resizeImageToDataUrl(file, 400);
    renderView();
  });
  logoField.appendChild(logoInput);
  card.appendChild(logoField);

  const colorField = document.createElement("div");
  colorField.className = "field color-field";
  const colorInputId = nextFieldId("field");
  colorField.innerHTML = `<label for="${colorInputId}">Accentkleur (gebruikt in de PDF-opmaak)</label>`;
  const colorInput = document.createElement("input");
  colorInput.id = colorInputId;
  colorInput.type = "color";
  colorInput.value = s.accentColor || "#1d4ed8";
  colorInput.addEventListener("input", () => (s.accentColor = colorInput.value));
  colorField.appendChild(colorInput);
  card.appendChild(colorField);

  card.appendChild(makeTextField("Bedrijfsnaam", s.companyName, (v) => (s.companyName = v)));
  card.appendChild(makeTextField("Straat en huisnummer", s.addressLine, (v) => (s.addressLine = v)));
  const rowAddr = document.createElement("div");
  rowAddr.className = "row";
  rowAddr.appendChild(makeTextField("Postcode", s.postalCode, (v) => (s.postalCode = v)));
  rowAddr.appendChild(makeTextField("Plaats", s.city, (v) => (s.city = v)));
  rowAddr.appendChild(makeTextField("Land", s.country, (v) => (s.country = v)));
  card.appendChild(rowAddr);
  const row = document.createElement("div");
  row.className = "row";
  row.appendChild(makeTextField("Btw-nummer", s.vatNumber, (v) => (s.vatNumber = v)));
  row.appendChild(makeTextField("KvK-nummer", s.kvkNumber, (v) => (s.kvkNumber = v)));
  card.appendChild(row);
  const rowBank = document.createElement("div");
  rowBank.className = "row";
  rowBank.appendChild(makeTextField("IBAN", s.iban, (v) => (s.iban = v)));
  rowBank.appendChild(makeTextField("Banknaam", s.bankName, (v) => (s.bankName = v)));
  rowBank.appendChild(makeTextField("BIC/SWIFT", s.bic, (v) => (s.bic = v)));
  card.appendChild(rowBank);
  const rowContact = document.createElement("div");
  rowContact.className = "row";
  rowContact.appendChild(makeTextField("Telefoonnummer", s.phone, (v) => (s.phone = v)));
  rowContact.appendChild(makeTextField("E-mailadres (bedrijf)", s.email, (v) => (s.email = v)));
  card.appendChild(rowContact);
  wrap.appendChild(card);

  const card2 = document.createElement("div");
  card2.className = "card";
  card2.innerHTML = `<h2>Standaardwaarden nieuwe factuur/offerte</h2>`;
  const rowDefaults = document.createElement("div");
  rowDefaults.className = "row";
  rowDefaults.appendChild(
    makeSelectField(
      "Standaard betalingstermijn",
      String(s.paymentTermDays),
      [["0", "Direct"], ["7", "7 dagen"], ["14", "14 dagen"], ["30", "30 dagen"]],
      (v) => (s.paymentTermDays = Number(v))
    )
  );
  rowDefaults.appendChild(
    makeSelectField(
      "Valuta",
      s.currency || "EUR",
      [
        ["EUR", "Euro (€)"],
        ["USD", "US dollar ($)"],
        ["GBP", "Brits pond (£)"],
        ["CHF", "Zwitserse frank (CHF)"],
        ["custom", "Andere (ISO-code)..."],
      ],
      (v) => {
        if (v === "custom") {
          const code = promptCurrencyCode(s.currency || "EUR");
          if (code !== null) s.currency = code;
        } else {
          s.currency = v;
        }
        renderView();
      }
    )
  );
  card2.appendChild(rowDefaults);
  const row2 = document.createElement("div");
  row2.className = "row";
  row2.appendChild(makeTextField("Volgend factuurnummer", suggestNumber("factuur"), () => {}, { readOnly: true }));
  row2.appendChild(makeTextField("Volgend offertenummer", suggestNumber("offerte"), () => {}, { readOnly: true }));
  card2.appendChild(row2);
  const hint = document.createElement("p");
  hint.style.cssText = "color:var(--text-muted);font-size:0.85rem;";
  hint.textContent = "Deze nummers worden automatisch opgehoogd zodra je een factuur/offerte met dat nummer opslaat. Je kunt het nummer bij het aanmaken altijd zelf aanpassen.";
  card2.appendChild(hint);
  wrap.appendChild(card2);

  const cardNotes = document.createElement("div");
  cardNotes.className = "card";
  cardNotes.innerHTML = `<h2>Notitiesjablonen</h2><p style="color:var(--text-muted);font-size:0.85rem;margin-top:-8px;">Kant-en-klare teksten voor het notitieveld op een factuur/offerte, in het Nederlands en Engels. Het standaardsjabloon wordt automatisch ingevuld bij een nieuwe factuur/offerte, in de taal die je daarvoor kiest.</p>`;
  const noteTemplates = s.noteTemplates || [];
  if (!noteTemplates.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "Nog geen sjablonen toegevoegd.";
    cardNotes.appendChild(empty);
  } else {
    noteTemplates.forEach((tpl) => {
      const isDefault = s.defaultNoteTemplateId === tpl.id;
      const row = document.createElement("div");
      row.className = "list-row";
      const preview = (tpl.nl || tpl.en || "").slice(0, 70);
      row.innerHTML = `
        <div class="col-main">
          <div class="name">${tpl.name}${isDefault ? ' <span class="badge badge-verzonden">Standaard</span>' : ""}</div>
          <div class="sub">${preview}${(tpl.nl || tpl.en || "").length > 70 ? "…" : ""}</div>
        </div>
      `;
      if (!isDefault) {
        const defBtn = document.createElement("button");
        defBtn.className = "btn btn-ghost btn-sm";
        defBtn.textContent = "Als standaard instellen";
        defBtn.addEventListener("click", () =>
          withBusy(defBtn, async () => {
            s.defaultNoteTemplateId = tpl.id;
            await saveSettings();
            toast("Standaardsjabloon ingesteld.");
            renderView();
          })
        );
        row.appendChild(defBtn);
      }
      const editBtn = document.createElement("button");
      editBtn.className = "btn btn-ghost btn-sm";
      editBtn.textContent = "Bewerken";
      editBtn.addEventListener("click", () => openNoteTemplateModal(tpl));
      row.appendChild(editBtn);

      const delBtn = document.createElement("button");
      delBtn.className = "btn btn-danger btn-sm";
      delBtn.textContent = "Verwijderen";
      delBtn.addEventListener("click", () =>
        withBusy(delBtn, async () => {
          if (!confirm(`Sjabloon "${tpl.name}" verwijderen?`)) return;
          s.noteTemplates = s.noteTemplates.filter((x) => x.id !== tpl.id);
          if (s.defaultNoteTemplateId === tpl.id) s.defaultNoteTemplateId = null;
          await saveSettings();
          renderView();
        })
      );
      row.appendChild(delBtn);

      cardNotes.appendChild(row);
    });
  }
  const addTplBtn = document.createElement("button");
  addTplBtn.className = "btn btn-secondary btn-sm";
  addTplBtn.textContent = "+ Nieuw sjabloon";
  addTplBtn.style.marginTop = "6px";
  addTplBtn.addEventListener("click", () => openNoteTemplateModal(null));
  cardNotes.appendChild(addTplBtn);
  wrap.appendChild(cardNotes);

  const card3 = document.createElement("div");
  card3.className = "card";
  card3.innerHTML = `<h2>Opslag</h2><p style="color:var(--text-muted);font-size:0.9rem;">Je bent ingelogd via <strong>${state.identity.label}</strong>. Facturen, offertes en klanten worden bewaard in jouw eigen account en zijn niet zichtbaar voor andere gebruikers van deze app.</p>`;
  wrap.appendChild(card3);

  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Instellingen opslaan";
  saveBtn.addEventListener("click", () =>
    withBusy(saveBtn, async () => {
      try {
        await saveSettings();
        toast("Instellingen opgeslagen.");
      } catch (err) {
        toast("Opslaan mislukt: " + err.message);
      }
    })
  );
  wrap.appendChild(saveBtn);

  return wrap;
}

function resizeImageToDataUrl(file, maxWidth) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/png"));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

// ==========================================================================
// PDF GENERATIE
// ==========================================================================
const PDF_I18N = {
  nl: {
    title: { factuur: "FACTUUR", offerte: "OFFERTE" },
    to: "Aan:",
    number: { factuur: "Factuurnummer:", offerte: "Offertenummer:" },
    date: "Datum:",
    dueDate: { factuur: "Vervaldatum:", offerte: "Geldig tot:" },
    reference: "Referentie:",
    vatNumber: "Btw-nummer:",
    kvk: "KvK-nummer:",
    iban: "IBAN:",
    bank: "Bank:",
    bic: "BIC:",
    phone: "Tel:",
    email: "E-mail:",
    colDate: "Datum",
    colDesc: "Omschrijving",
    colQty: "Aantal",
    colPrice: "Prijs",
    colVat: "Btw",
    colTotal: "Totaal",
    subtotal: "Subtotaal excl. btw",
    vatGroup: (rate) => `Btw ${rate}%`,
    grandTotal: "Totaal incl. btw",
    paymentNote: (date, iban, number) => `Gelieve te betalen voor ${date} op IBAN ${iban} o.v.v. ${number}.`,
  },
  en: {
    title: { factuur: "INVOICE", offerte: "QUOTATION" },
    to: "To:",
    number: { factuur: "Invoice number:", offerte: "Quotation number:" },
    date: "Date:",
    dueDate: { factuur: "Due date:", offerte: "Valid until:" },
    reference: "Reference:",
    vatNumber: "VAT number:",
    kvk: "Chamber of Commerce no.:",
    iban: "IBAN:",
    bank: "Bank:",
    bic: "BIC:",
    phone: "Phone:",
    email: "Email:",
    colDate: "Date",
    colDesc: "Description",
    colQty: "Quantity",
    colPrice: "Price",
    colVat: "VAT",
    colTotal: "Total",
    subtotal: "Subtotal excl. VAT",
    vatGroup: (rate) => `VAT ${rate}%`,
    grandTotal: "Total incl. VAT",
    paymentNote: (date, iban, number) => `Please pay before ${date} to IBAN ${iban}, quoting ${number}.`,
  },
};

function pdfT(lang) {
  return PDF_I18N[lang] || PDF_I18N.nl;
}

async function generatePDF(inv) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const s = state.settings;
  const t = pdfT(inv.language);
  const marginX = 40;
  let y = 50;

  // Eén gedeelde kolom-X voor ALLE label:waarde-blokken op deze pagina
  // (bedrijfsgegevens hier, en de factuurgegevens verderop) - zo vormen ze
  // samen visueel één nette, doorlopende kolom, met steeds dezelfde
  // scheiding tussen "onderdeel" (label) en de ingevulde waarde. Ook dezelfde
  // fontSize overal (9.5) voor adresgegevens, bedrijfsgegevens én
  // factuurgegevens, zodat het er als één samenhangend geheel uitziet.
  const infoFontSize = 9.5;
  const infoValueX = 445;

  // Bereken eerst hoe hoog het bedrijfsgegevens-blok (rechtsboven) wordt,
  // vóórdat het logo getekend wordt: het logo schaalt mee tot ongeveer
  // diezelfde hoogte, zodat er geen klein logootje met veel witruimte
  // eronder ontstaat.
  const companyHeaderLines = [
    s.companyName,
    s.addressLine,
    [s.postalCode, s.city].filter(Boolean).join(" "),
    s.country,
  ].filter(Boolean);
  const companyFieldPairs = [
    s.vatNumber ? [t.vatNumber, s.vatNumber] : null,
    s.kvkNumber ? [t.kvk, s.kvkNumber] : null,
    s.iban ? [t.iban, s.iban] : null,
    s.bankName ? [t.bank, s.bankName] : null,
    s.bic ? [t.bic, s.bic] : null,
    s.phone ? [t.phone, s.phone] : null,
    s.email ? [t.email, s.email] : null,
  ].filter(Boolean);
  let companyBlockHeight = companyHeaderLines.length * 13;
  if (companyFieldPairs.length) {
    companyBlockHeight += companyHeaderLines.length ? 8 : 0;
    companyBlockHeight += companyFieldPairs.length * 13;
  }
  const companyBottom = y + companyBlockHeight;

  let logoBottom = y;
  if (s.logoDataUrl) {
    try {
      const props = doc.getImageProperties(s.logoDataUrl);
      const minH = 46;
      const maxH = 110;
      const maxW = 210;
      let h = Math.max(minH, Math.min(maxH, companyBlockHeight || minH));
      let w = (props.width / props.height) * h;
      if (w > maxW) {
        w = maxW;
        h = (props.height / props.width) * w;
      }
      doc.addImage(s.logoDataUrl, "PNG", marginX, y, w, h);
      logoBottom = y + h;
    } catch (e) {
      console.warn("Logo kon niet worden toegevoegd aan PDF", e);
    }
  }

  // Bedrijfsnaam + adres (links uitgelijnd op infoValueX) en de losse
  // bedrijfsgegevens eronder, in dezelfde twee-koloms indeling als de
  // factuurgegevens verderop.
  doc.setFontSize(infoFontSize);
  doc.setTextColor(60);
  doc.text(companyHeaderLines, infoValueX, y);
  if (companyFieldPairs.length) {
    let fieldY = y + companyHeaderLines.length * 13 + (companyHeaderLines.length ? 8 : 0);
    companyFieldPairs.forEach((pair) => {
      doc.setTextColor(120);
      doc.text(pair[0], infoValueX - 6, fieldY, { align: "right" });
      doc.setTextColor(50);
      doc.text(pair[1], infoValueX, fieldY);
      fieldY += 13;
    });
  }

  // Titel + lijn komen nu ONDER zowel het logo als de eigen bedrijfsgegevens
  // te staan (over de volle breedte) - dat maakt de knip tussen het
  // briefhoofd en de rest van de factuur/offerte duidelijk zichtbaar.
  // Daaronder komt pas de indeling "Aan:" (links) / factuurgegevens (rechts).
  y = Math.max(logoBottom, companyBottom) + 26;
  doc.setFontSize(20);
  doc.setTextColor(20);
  doc.text(t.title[inv.type], marginX, y);
  y += 10;
  doc.setDrawColor(220);
  doc.line(marginX, y, 555, y);

  y += 28;
  doc.setFontSize(11);
  doc.setTextColor(40);
  doc.text(t.to, marginX, y);
  doc.setFontSize(infoFontSize);
  const clientAddressLines = [
    inv.clientSnapshot.addressLine,
    [inv.clientSnapshot.postalCode, inv.clientSnapshot.city].filter(Boolean).join(" "),
    inv.clientSnapshot.country,
  ].filter(Boolean);
  const clientLines = [
    inv.clientSnapshot.name,
    ...clientAddressLines,
    inv.clientSnapshot.vatNumber ? t.vatNumber + " " + inv.clientSnapshot.vatNumber : "",
  ].filter(Boolean);
  doc.text(clientLines, marginX, y + 16);
  const clientBlockHeight = 16 + clientLines.length * 13;

  // Zelfde twee-koloms indeling (en dezelfde infoValueX) als het
  // bedrijfsgegevens-blok hierboven, zodat beide blokken één doorlopende
  // kolom vormen. De waarde mag wrappen naar een volgende regel als hij niet
  // past - zo krijgt bijvoorbeeld een lange referentie altijd genoeg ruimte
  // i.p.v. krap tegen de kantlijn te komen.
  const metaValueWidth = 555 - infoValueX;
  doc.setFontSize(infoFontSize);
  const metaLines = [
    [t.number[inv.type], inv.number],
    [t.date, fmtDate(inv.date)],
    [t.dueDate[inv.type], fmtDate(inv.dueDate)],
  ];
  if (inv.reference) metaLines.push([t.reference, inv.reference]);
  let metaY = y;
  metaLines.forEach((pair) => {
    const valueLines = doc.splitTextToSize(String(pair[1] ?? ""), metaValueWidth);
    doc.setTextColor(120);
    doc.text(pair[0], infoValueX - 6, metaY, { align: "right" });
    doc.setTextColor(20);
    doc.text(valueLines, infoValueX, metaY);
    metaY += Math.max(1, valueLines.length) * 16;
  });
  const metaBlockHeight = metaY - y;

  y += Math.max(clientBlockHeight, metaBlockHeight, 70) + 14;

  const rows = inv.items.map((it) => [
    itemDateLabel(it),
    it.desc || "",
    String(it.qty) + (unitDisplayLabel(it, inv.language) ? " " + unitDisplayLabel(it, inv.language) : ""),
    fmtMoney(Number(it.price), inv.currency),
    Number(it.vatRate) + "%",
    fmtMoney((Number(it.qty) || 0) * (Number(it.price) || 0), inv.currency),
  ]);

  doc.autoTable({
    startY: y,
    head: [[t.colDate, t.colDesc, t.colQty, t.colPrice, t.colVat, t.colTotal]],
    body: rows,
    margin: { left: marginX, right: 40 },
    styles: { fontSize: 9.5, cellPadding: 6 },
    headStyles: { fillColor: hexToRgb(s.accentColor), textColor: 255 },
    columnStyles: {
      0: { cellWidth: 85 },
      2: { halign: "right", cellWidth: 65 },
      3: { halign: "right", cellWidth: 75 },
      4: { halign: "right", cellWidth: 45 },
      5: { halign: "right", cellWidth: 75 },
    },
  });

  let finalY = doc.lastAutoTable.finalY + 20;
  const totals = calcTotals(inv.items);

  // Paginabreek-veiligheid: totalen/notities/betaalregel worden met plain
  // doc.text() getekend (i.t.t. de regel-tabel hierboven, die autoTable's
  // eigen paginering al afhandelt) - dus hier zelf controleren of het nog
  // op de huidige pagina past, anders een nieuwe pagina beginnen.
  const pageHeight = doc.internal.pageSize.getHeight();
  const bottomMargin = 50;
  function ensureSpace(neededHeight) {
    if (finalY + neededHeight > pageHeight - bottomMargin) {
      doc.addPage();
      finalY = 50;
    }
  }

  const groupCount = Object.keys(totals.groups).length;
  ensureSpace(16 * (1 + groupCount) + 56);

  const totalsX = 380;
  doc.setFontSize(10);
  doc.setTextColor(60);
  doc.text(t.subtotal, totalsX, finalY);
  doc.text(fmtMoney(totals.subtotal, inv.currency), 555, finalY, { align: "right" });
  finalY += 16;
  Object.keys(totals.groups)
    .sort()
    .forEach((rate) => {
      doc.text(t.vatGroup(rate), totalsX, finalY);
      doc.text(fmtMoney(totals.groups[rate].vat, inv.currency), 555, finalY, { align: "right" });
      finalY += 16;
    });
  doc.setDrawColor(20);
  doc.line(totalsX, finalY - 4, 555, finalY - 4);
  doc.setFontSize(12);
  doc.setTextColor(20);
  doc.text(t.grandTotal, totalsX, finalY + 12);
  doc.text(fmtMoney(totals.grandTotal, inv.currency), 555, finalY + 12, { align: "right" });
  finalY += 40;

  if (inv.notes) {
    doc.setFontSize(9.5);
    doc.setTextColor(90);
    const split = doc.splitTextToSize(inv.notes, 515);
    ensureSpace(split.length * 12 + 10);
    doc.text(split, marginX, finalY);
    finalY += split.length * 12 + 10;
  }

  if (inv.type === "factuur" && s.iban) {
    ensureSpace(20);
    doc.setFontSize(9.5);
    doc.setTextColor(90);
    doc.text(t.paymentNote(fmtDate(inv.dueDate), s.iban, inv.number), marginX, finalY);
  }

  doc.save(`${inv.number}.pdf`);

  // Ook een kopie van de PDF opslaan in dezelfde cloudmap als de gegevens,
  // zodat je 'm ook terugvindt naast settings.json/invoices.json. Dit mag
  // de download zelf nooit blokkeren, dus alleen een toast bij mislukking.
  const provider = activeProvider();
  if (provider && provider.writeBinary) {
    try {
      const blob = doc.output("blob");
      await provider.writeBinary(`${inv.number}.pdf`, blob);
    } catch (err) {
      console.warn("PDF uploaden naar cloudopslag mislukt", err);
      toast("PDF is gedownload, maar kon niet naar " + provider.name + " worden geüpload: " + err.message);
    }
  }
}

// ==========================================================================
// INIT
// ==========================================================================
window.addEventListener("DOMContentLoaded", async () => {
  setupNav();
  setupLoginButtons();
  try {
    googleProvider.init();
  } catch (e) {
    console.warn("Google-configuratie kon niet worden geïnitialiseerd", e);
  }
  try {
    microsoftProvider.init();
  } catch (e) {
    console.warn("Microsoft-configuratie kon niet worden geïnitialiseerd", e);
  }

  try {
    const restored = await tryRestoreSession();
    if (restored) {
      await afterLogin();
    }
  } catch (e) {
    console.warn("Sessie herstellen mislukt, toon inlogscherm", e);
  }
});
