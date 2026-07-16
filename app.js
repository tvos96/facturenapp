import { GOOGLE_CLIENT_ID, DRIVE_SCOPE, APP_FOLDER_NAME, DEFAULT_VAT_RATES } from "./config.js";

// ==========================================================================
// STATE
// ==========================================================================
const state = {
  token: null,
  tokenExpiresAt: 0,
  folderId: null,
  fileIds: { settings: null, clients: null, invoices: null },
  settings: null,
  clients: [],
  invoices: [],
  view: "new",
  draft: null,          // huidige factuur/offerte in bewerking
  archiveFilter: { type: "alle", status: "alle", q: "" },
  saving: false,
};

const DEFAULT_SETTINGS = {
  companyName: "",
  address: "",
  vatNumber: "",
  iban: "",
  logoDataUrl: "",
  paymentTermDays: 14,
  nextFactuurNummer: 1,
  nextOfferteNummer: 1,
  numberPrefix: { factuur: "F", offerte: "O" },
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

// ==========================================================================
// GOOGLE AUTH
// ==========================================================================
let tokenClient = null;

function initAuth() {
  if (typeof google === "undefined" || !google.accounts) {
    showLoginError("Kon Google Identity Services niet laden. Controleer je internetverbinding en herlaad de pagina.");
    document.getElementById("login-btn").disabled = true;
    return;
  }
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: async (resp) => {
      if (resp.error) {
        showLoginError("Inloggen mislukt: " + resp.error);
        return;
      }
      state.token = resp.access_token;
      state.tokenExpiresAt = Date.now() + (resp.expires_in - 60) * 1000;
      sessionStorage.setItem("gd_token", state.token);
      sessionStorage.setItem("gd_token_exp", String(state.tokenExpiresAt));
      await afterLogin();
    },
  });

  document.getElementById("login-btn").addEventListener("click", () => {
    showLoginError("");
    if (GOOGLE_CLIENT_ID.includes("VUL_HIER")) {
      showLoginError("Er is nog geen Google Client ID ingevuld in config.js. Zie README.md.");
      return;
    }
    tokenClient.requestAccessToken({ prompt: "" });
  });

  document.getElementById("logout-btn").addEventListener("click", () => {
    if (state.token) {
      google.accounts.oauth2.revoke(state.token, () => {});
    }
    sessionStorage.removeItem("gd_token");
    sessionStorage.removeItem("gd_token_exp");
    state.token = null;
    location.reload();
  });

  // Probeer een bewaarde sessie te hergebruiken (binnen dezelfde browsertab)
  const savedToken = sessionStorage.getItem("gd_token");
  const savedExp = Number(sessionStorage.getItem("gd_token_exp") || 0);
  if (savedToken && savedExp > Date.now()) {
    state.token = savedToken;
    state.tokenExpiresAt = savedExp;
    afterLogin();
  }
}

function showLoginError(msg) {
  document.getElementById("login-error").textContent = msg;
}

async function afterLogin() {
  document.getElementById("login-screen").classList.add("hidden");
  document.getElementById("app").classList.remove("hidden");
  setSyncStatus("Synchroniseren...");
  try {
    await ensureAppFolder();
    await loadAllData();
    setSyncStatus("");
    renderView();
  } catch (err) {
    console.error(err);
    setSyncStatus("");
    toast("Fout bij laden van Google Drive: " + err.message);
  }
}

function setSyncStatus(txt) {
  document.getElementById("sync-status").textContent = txt;
}

// ==========================================================================
// GOOGLE DRIVE HELPERS
// ==========================================================================
async function driveFetch(url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: "Bearer " + state.token,
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Drive API fout (${res.status}): ${text.slice(0, 200)}`);
  }
  return res;
}

async function driveFindByName(name, parentId) {
  const q = encodeURIComponent(
    `name='${name.replace(/'/g, "\\'")}' and trashed=false` +
      (parentId ? ` and '${parentId}' in parents` : "")
  );
  const res = await driveFetch(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&spaces=drive`
  );
  const data = await res.json();
  return data.files && data.files.length ? data.files[0] : null;
}

async function ensureAppFolder() {
  const cached = sessionStorage.getItem("gd_folder_id");
  if (cached) {
    state.folderId = cached;
    return;
  }
  let folder = await driveFindByName(APP_FOLDER_NAME, null);
  if (!folder) {
    const res = await driveFetch("https://www.googleapis.com/drive/v3/files", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name: APP_FOLDER_NAME,
        mimeType: "application/vnd.google-apps.folder",
      }),
    });
    folder = await res.json();
  }
  state.folderId = folder.id;
  sessionStorage.setItem("gd_folder_id", folder.id);
}

async function driveReadJSON(fileName, fallback) {
  const file = await driveFindByName(fileName, state.folderId);
  if (!file) return { fileId: null, data: fallback };
  const res = await driveFetch(
    `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`
  );
  const data = await res.json();
  return { fileId: file.id, data };
}

async function driveWriteJSON(fileName, fileId, data) {
  const content = JSON.stringify(data, null, 2);
  if (fileId) {
    await driveFetch(
      `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: content,
      }
    );
    return fileId;
  } else {
    const boundary = "-------facturenapp" + Date.now();
    const metadata = { name: fileName, parents: [state.folderId] };
    const body =
      `--${boundary}\r\n` +
      `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
      `${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\n` +
      `Content-Type: application/json\r\n\r\n` +
      `${content}\r\n` +
      `--${boundary}--`;
    const res = await driveFetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
      {
        method: "POST",
        headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
        body,
      }
    );
    const created = await res.json();
    return created.id;
  }
}

// ==========================================================================
// STORE (settings / clients / invoices)
// ==========================================================================
async function loadAllData() {
  const [settingsRes, clientsRes, invoicesRes] = await Promise.all([
    driveReadJSON("settings.json", DEFAULT_SETTINGS),
    driveReadJSON("clients.json", []),
    driveReadJSON("invoices.json", []),
  ]);
  state.settings = { ...DEFAULT_SETTINGS, ...settingsRes.data };
  state.fileIds.settings = settingsRes.fileId;
  state.clients = clientsRes.data || [];
  state.fileIds.clients = clientsRes.fileId;
  state.invoices = invoicesRes.data || [];
  state.fileIds.invoices = invoicesRes.fileId;
}

async function saveSettings() {
  setSyncStatus("Opslaan...");
  state.fileIds.settings = await driveWriteJSON("settings.json", state.fileIds.settings, state.settings);
  setSyncStatus("");
}

async function saveClients() {
  setSyncStatus("Opslaan...");
  state.fileIds.clients = await driveWriteJSON("clients.json", state.fileIds.clients, state.clients);
  setSyncStatus("");
}

async function saveInvoices() {
  setSyncStatus("Opslaan...");
  state.fileIds.invoices = await driveWriteJSON("invoices.json", state.fileIds.invoices, state.invoices);
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
  const groups = {}; // vatRate -> { subtotal, vat }
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

function fmtMoney(n) {
  return new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR" }).format(n || 0);
}

function fmtDate(iso) {
  if (!iso) return "";
  return new Date(iso + "T00:00:00").toLocaleDateString("nl-NL");
}

function emptyDraft(type = "factuur") {
  const date = todayISO();
  return {
    id: crypto.randomUUID(),
    type,
    number: suggestNumber(type),
    status: "concept",
    date,
    paymentTermDays: state.settings.paymentTermDays,
    dueDate: addDays(date, state.settings.paymentTermDays),
    clientId: null,
    clientSnapshot: { name: "", address: "", vatNumber: "", email: "" },
    items: [{ desc: "", qty: 1, price: 0, vatRate: DEFAULT_VAT_RATES[0] }],
    notes: "",
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
// NAVIGATION
// ==========================================================================
function setupNav() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const view = btn.dataset.view;
      if (view === "new" && (!state.draft || state.draft.__saved)) {
        state.draft = emptyDraft();
      }
      state.view = view;
      renderView();
    });
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

  // Type toggle
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

  // Klant kiezen
  const clientField = document.createElement("div");
  clientField.className = "field";
  clientField.innerHTML = `<label>Klant</label>`;
  const clientSelect = document.createElement("select");
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
      ? { name: c.name, address: c.address, vatNumber: c.vatNumber, email: c.email }
      : { name: "", address: "", vatNumber: "", email: "" };
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
  row2.appendChild(makeTextAreaField("Adres", draft.clientSnapshot.address, (v) => (draft.clientSnapshot.address = v)));
  row2.appendChild(makeTextField("Btw-nummer klant", draft.clientSnapshot.vatNumber, (v) => (draft.clientSnapshot.vatNumber = v)));
  card.appendChild(row2);

  wrap.appendChild(card);

  // Nummer, datum, termijn, status
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
      draft.dueDate = addDays(v, draft.paymentTermDays);
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
          const n = Number(prompt("Aantal dagen betalingstermijn:", draft.paymentTermDays) || draft.paymentTermDays);
          draft.paymentTermDays = n;
        } else {
          draft.paymentTermDays = Number(v);
        }
        draft.dueDate = addDays(draft.date, draft.paymentTermDays);
        renderView();
      }
    )
  );
  row4.appendChild(
    makeDateField(draft.type === "factuur" ? "Vervaldatum" : "Geldig tot", draft.dueDate, (v) => (draft.dueDate = v))
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
  wrap.appendChild(card2);

  // Regels
  const card3 = document.createElement("div");
  card3.className = "card";
  card3.innerHTML = `<h2>Regels</h2>`;
  card3.appendChild(renderItemsTable(draft));

  const addBtn = document.createElement("button");
  addBtn.className = "btn btn-secondary btn-sm";
  addBtn.textContent = "+ Regel toevoegen";
  addBtn.addEventListener("click", () => {
    draft.items.push({ desc: "", qty: 1, price: 0, vatRate: DEFAULT_VAT_RATES[0] });
    renderView();
  });
  card3.appendChild(addBtn);

  const totals = calcTotals(draft.items);
  const totalsBox = document.createElement("div");
  totalsBox.className = "totals-box";
  let totalsHtml = `<div class="totals-row"><span>Subtotaal excl. btw</span><span>${fmtMoney(totals.subtotal)}</span></div>`;
  Object.keys(totals.groups)
    .sort()
    .forEach((rate) => {
      totalsHtml += `<div class="totals-row"><span>Btw ${rate}%</span><span>${fmtMoney(totals.groups[rate].vat)}</span></div>`;
    });
  totalsHtml += `<div class="totals-row grand"><span>Totaal incl. btw</span><span>${fmtMoney(totals.grandTotal)}</span></div>`;
  totalsBox.innerHTML = totalsHtml;
  card3.appendChild(totalsBox);
  wrap.appendChild(card3);

  // Notities
  const card4 = document.createElement("div");
  card4.className = "card";
  card4.appendChild(makeTextAreaField("Notities (optioneel, komt op de " + draft.type + ")", draft.notes, (v) => (draft.notes = v)));
  wrap.appendChild(card4);

  // Acties
  const actions = document.createElement("div");
  actions.className = "row";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = draft.__editing ? "Wijzigingen opslaan" : "Opslaan in archief";
  saveBtn.addEventListener("click", () => saveDraft(draft));
  actions.appendChild(saveBtn);

  const pdfBtn = document.createElement("button");
  pdfBtn.className = "btn btn-secondary";
  pdfBtn.textContent = "Download PDF";
  pdfBtn.addEventListener("click", () => generatePDF(draft));
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
    <th>Omschrijving</th><th class="col-qty">Aantal</th><th class="col-price">Prijs</th>
    <th class="col-vat">Btw</th><th class="col-total">Totaal</th><th class="col-remove"></th>
  </tr></thead>`;
  const tbody = document.createElement("tbody");
  draft.items.forEach((item, idx) => {
    const tr = document.createElement("tr");

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
        const n = Number(prompt("Aangepast btw-percentage:", item.vatRate) ?? item.vatRate);
        item.vatRate = n;
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
    tdTotal.textContent = fmtMoney((Number(item.qty) || 0) * (Number(item.price) || 0));
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
  return table;
}

function updateLineTotal(idx) {
  const draft = state.draft;
  const item = draft.items[idx];
  const el = document.getElementById("line-total-" + idx);
  if (el) el.textContent = fmtMoney((Number(item.qty) || 0) * (Number(item.price) || 0));
  // Totalen-blok live bijwerken zonder hele view te her-renderen (voorkomt focus-verlies)
  const totals = calcTotals(draft.items);
  const totalsBox = document.querySelector(".totals-box");
  if (totalsBox) {
    let html = `<div class="totals-row"><span>Subtotaal excl. btw</span><span>${fmtMoney(totals.subtotal)}</span></div>`;
    Object.keys(totals.groups)
      .sort()
      .forEach((rate) => {
        html += `<div class="totals-row"><span>Btw ${rate}%</span><span>${fmtMoney(totals.groups[rate].vat)}</span></div>`;
      });
    html += `<div class="totals-row grand"><span>Totaal incl. btw</span><span>${fmtMoney(totals.grandTotal)}</span></div>`;
    totalsBox.innerHTML = html;
  }
}

function makeTextField(label, value, onChange, opts = {}) {
  const div = document.createElement("div");
  div.className = "field";
  const l = document.createElement("label");
  l.textContent = label;
  const input = document.createElement("input");
  input.type = "text";
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
  const l = document.createElement("label");
  l.textContent = label;
  const ta = document.createElement("textarea");
  ta.value = value || "";
  ta.addEventListener("input", () => onChange(ta.value));
  div.appendChild(l);
  div.appendChild(ta);
  return div;
}

function makeDateField(label, value, onChange) {
  const div = document.createElement("div");
  div.className = "field";
  const l = document.createElement("label");
  l.textContent = label;
  const input = document.createElement("input");
  input.type = "date";
  input.value = value || "";
  input.addEventListener("change", () => onChange(input.value));
  div.appendChild(l);
  div.appendChild(input);
  return div;
}

function makeSelectField(label, value, options, onChange) {
  const div = document.createElement("div");
  div.className = "field";
  const l = document.createElement("label");
  l.textContent = label;
  const select = document.createElement("select");
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
      return (
        inv.number.toLowerCase().includes(q) ||
        (inv.clientSnapshot.name || "").toLowerCase().includes(q)
      );
    })
    .sort((a, b) => (b.date > a.date ? 1 : -1));
}

function renderArchiveList(list) {
  const items = getFilteredInvoices();
  list.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "Geen facturen of offertes gevonden.";
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
      <div class="amount">${fmtMoney(inv.total)}</div>
    `;
    const actions = document.createElement("div");
    actions.className = "actions";

    const pdfBtn = document.createElement("button");
    pdfBtn.className = "btn btn-secondary btn-sm";
    pdfBtn.textContent = "PDF";
    pdfBtn.addEventListener("click", () => generatePDF(inv));
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
        <div class="sub">${c.address || ""} ${c.vatNumber ? "· btw: " + c.vatNumber : ""}</div>
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
    delBtn.addEventListener("click", async () => {
      if (!confirm(`Klant "${c.name}" verwijderen?`)) return;
      state.clients = state.clients.filter((x) => x.id !== c.id);
      await saveClients();
      renderView();
    });
    row.appendChild(delBtn);

    wrap.appendChild(row);
  });

  return wrap;
}

function openClientModal(client) {
  const isNew = !client;
  const draftClient = client ? { ...client } : { id: crypto.randomUUID(), name: "", address: "", vatNumber: "", email: "" };

  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.innerHTML = `<h2>${isNew ? "Nieuwe klant" : "Klant bewerken"}</h2>`;

  modal.appendChild(makeTextField("Naam", draftClient.name, (v) => (draftClient.name = v)));
  modal.appendChild(makeTextAreaField("Adres", draftClient.address, (v) => (draftClient.address = v)));
  modal.appendChild(makeTextField("Btw-nummer", draftClient.vatNumber, (v) => (draftClient.vatNumber = v)));
  modal.appendChild(makeTextField("E-mail", draftClient.email, (v) => (draftClient.email = v)));

  const actions = document.createElement("div");
  actions.className = "row";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Opslaan";
  saveBtn.addEventListener("click", async () => {
    if (!draftClient.name.trim()) {
      toast("Vul een naam in.");
      return;
    }
    const idx = state.clients.findIndex((c) => c.id === draftClient.id);
    if (idx >= 0) state.clients[idx] = draftClient;
    else state.clients.push(draftClient);
    await saveClients();
    backdrop.remove();
    renderView();
  });
  actions.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-ghost";
  cancelBtn.textContent = "Annuleren";
  cancelBtn.addEventListener("click", () => backdrop.remove());
  actions.appendChild(cancelBtn);

  modal.appendChild(actions);
  backdrop.appendChild(modal);
  backdrop.addEventListener("click", (e) => {
    if (e.target === backdrop) backdrop.remove();
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
  logoField.innerHTML = `<label>Logo</label>`;
  const logoInput = document.createElement("input");
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

  card.appendChild(makeTextField("Bedrijfsnaam", s.companyName, (v) => (s.companyName = v)));
  card.appendChild(makeTextAreaField("Adres", s.address, (v) => (s.address = v)));
  const row = document.createElement("div");
  row.className = "row";
  row.appendChild(makeTextField("Btw-nummer", s.vatNumber, (v) => (s.vatNumber = v)));
  row.appendChild(makeTextField("IBAN", s.iban, (v) => (s.iban = v)));
  card.appendChild(row);
  wrap.appendChild(card);

  const card2 = document.createElement("div");
  card2.className = "card";
  card2.innerHTML = `<h2>Standaardwaarden nieuwe factuur/offerte</h2>`;
  card2.appendChild(
    makeSelectField(
      "Standaard betalingstermijn",
      String(s.paymentTermDays),
      [["0", "Direct"], ["7", "7 dagen"], ["14", "14 dagen"], ["30", "30 dagen"]],
      (v) => (s.paymentTermDays = Number(v))
    )
  );
  const row2 = document.createElement("div");
  row2.className = "row";
  row2.appendChild(makeTextField("Volgend factuurnummer", suggestNumber("factuur"), () => {}, { readOnly: true }));
  row2.appendChild(makeTextField("Volgend offertenummer", suggestNumber("offerte"), () => {}, { readOnly: true }));
  card2.appendChild(row2);
  const hint = document.createElement("p");
  hint.style.cssText = "color:#667085;font-size:0.85rem;";
  hint.textContent = "Deze nummers worden automatisch opgehoogd zodra je een factuur/offerte met dat nummer opslaat. Je kunt het nummer bij het aanmaken altijd zelf aanpassen.";
  card2.appendChild(hint);
  wrap.appendChild(card2);

  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Instellingen opslaan";
  saveBtn.addEventListener("click", async () => {
    try {
      await saveSettings();
      toast("Instellingen opgeslagen.");
    } catch (err) {
      toast("Opslaan mislukt: " + err.message);
    }
  });
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
function generatePDF(inv) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const s = state.settings;
  const marginX = 40;
  let y = 50;

  if (s.logoDataUrl) {
    try {
      const props = doc.getImageProperties(s.logoDataUrl);
      const w = 90;
      const h = (props.height / props.width) * w;
      doc.addImage(s.logoDataUrl, "PNG", marginX, y, w, h);
    } catch (e) {
      console.warn("Logo kon niet worden toegevoegd aan PDF", e);
    }
  }

  doc.setFontSize(10);
  doc.setTextColor(60);
  const companyLines = [s.companyName, s.address, s.vatNumber ? "Btw-nummer: " + s.vatNumber : "", s.iban ? "IBAN: " + s.iban : ""].filter(Boolean);
  doc.text(companyLines, 555, y, { align: "right" });

  y += 70;
  doc.setFontSize(20);
  doc.setTextColor(20);
  doc.text(inv.type === "factuur" ? "FACTUUR" : "OFFERTE", marginX, y);

  y += 10;
  doc.setDrawColor(220);
  doc.line(marginX, y, 555, y);

  y += 28;
  doc.setFontSize(11);
  doc.setTextColor(40);
  doc.text("Aan:", marginX, y);
  doc.setFontSize(10);
  const clientLines = [
    inv.clientSnapshot.name,
    inv.clientSnapshot.address,
    inv.clientSnapshot.vatNumber ? "Btw-nummer: " + inv.clientSnapshot.vatNumber : "",
  ].filter(Boolean);
  doc.text(clientLines, marginX, y + 16);

  const metaX = 350;
  doc.setFontSize(10);
  const metaLines = [
    [inv.type === "factuur" ? "Factuurnummer:" : "Offertenummer:", inv.number],
    ["Datum:", fmtDate(inv.date)],
    [inv.type === "factuur" ? "Vervaldatum:" : "Geldig tot:", fmtDate(inv.dueDate)],
  ];
  metaLines.forEach((pair, i) => {
    doc.setTextColor(120);
    doc.text(pair[0], metaX, y + i * 16);
    doc.setTextColor(20);
    doc.text(pair[1], metaX + 100, y + i * 16);
  });

  y += 70;

  const rows = inv.items.map((it) => [
    it.desc || "",
    String(it.qty),
    fmtMoney(Number(it.price)),
    Number(it.vatRate) + "%",
    fmtMoney((Number(it.qty) || 0) * (Number(it.price) || 0)),
  ]);

  doc.autoTable({
    startY: y,
    head: [["Omschrijving", "Aantal", "Prijs", "Btw", "Totaal"]],
    body: rows,
    margin: { left: marginX, right: 40 },
    styles: { fontSize: 9.5, cellPadding: 6 },
    headStyles: { fillColor: [29, 78, 216], textColor: 255 },
    columnStyles: {
      1: { halign: "right", cellWidth: 55 },
      2: { halign: "right", cellWidth: 75 },
      3: { halign: "right", cellWidth: 45 },
      4: { halign: "right", cellWidth: 75 },
    },
  });

  let finalY = doc.lastAutoTable.finalY + 20;
  const totals = calcTotals(inv.items);
  const totalsX = 380;
  doc.setFontSize(10);
  doc.setTextColor(60);
  doc.text("Subtotaal excl. btw", totalsX, finalY);
  doc.text(fmtMoney(totals.subtotal), 555, finalY, { align: "right" });
  finalY += 16;
  Object.keys(totals.groups)
    .sort()
    .forEach((rate) => {
      doc.text(`Btw ${rate}%`, totalsX, finalY);
      doc.text(fmtMoney(totals.groups[rate].vat), 555, finalY, { align: "right" });
      finalY += 16;
    });
  doc.setDrawColor(20);
  doc.line(totalsX, finalY - 4, 555, finalY - 4);
  doc.setFontSize(12);
  doc.setTextColor(20);
  doc.text("Totaal incl. btw", totalsX, finalY + 12);
  doc.text(fmtMoney(totals.grandTotal), 555, finalY + 12, { align: "right" });
  finalY += 40;

  if (inv.notes) {
    doc.setFontSize(9.5);
    doc.setTextColor(90);
    const split = doc.splitTextToSize(inv.notes, 515);
    doc.text(split, marginX, finalY);
    finalY += split.length * 12 + 10;
  }

  if (inv.type === "factuur" && s.iban) {
    doc.setFontSize(9.5);
    doc.setTextColor(90);
    doc.text(`Gelieve te betalen voor ${fmtDate(inv.dueDate)} op IBAN ${s.iban} o.v.v. ${inv.number}.`, marginX, finalY);
  }

  doc.save(`${inv.number}.pdf`);
}

// ==========================================================================
// INIT
// ==========================================================================
window.addEventListener("DOMContentLoaded", () => {
  setupNav();
  initAuth();
  // state.draft wordt pas aangemaakt nadat instellingen zijn geladen
  // (zie afterLogin -> renderView / renderNewView), omdat de nummering
  // en standaard betalingstermijn afhankelijk zijn van state.settings.
});
