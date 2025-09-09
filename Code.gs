/***** Donation Web App – Server (Code.gs) *****/

/* ============ CONFIG ============ */
// REQUIRED: paste your Google Sheet ID between the quotes:
// (Find it in the URL of your sheet: https://docs.google.com/spreadsheets/d/<THIS_PART>/edit)
const SHEET_ID   = "YOUR_SHEET_ID_HERE";

const SHEET_LOG  = "Donations_Log";
const SHEET_CHAR = "Charities";
const SHEET_GUIDE= "ValueGuide_Custom";
/* =================================*/

function _ss() {
  // Open by ID so this works even in a standalone Web App
  return SpreadsheetApp.openById(SHEET_ID);
}

// Serve UI
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Donations")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Helpers ---------- */
function normalize_(s){ return (s||"").toString().trim().toLowerCase(); }

function listCharities() {
  try {
    const sh = _ss().getSheetByName(SHEET_CHAR);
    if (!sh) return {ok:false,error:`Sheet "${SHEET_CHAR}" not found`, data:[]};
    const vals = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), 2).getValues();
    const data = vals.filter(r => r[0]).map(([name, addr]) => ({name, address: addr || ""}));
    return {ok:true, data};
  } catch (e) {
    return {ok:false, error: String(e), data:[]};
  }
}

function listItems() {
  try {
    const sh = _ss().getSheetByName(SHEET_GUIDE);
    if (!sh) return {ok:false,error:`Sheet "${SHEET_GUIDE}" not found`, data:[]};
    const vals = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), 4).getValues();
    const data = vals
      .filter(r => r[1])
      .map(([cat, item, low, high]) => ({
        item, low: Number(low || 0), high: Number(high || 0), category: cat
      }));
    return {ok:true, data};
  } catch (e) {
    return {ok:false, error: String(e), data:[]};
  }
}

// Smart lookup: exact → startsWith → contains; also match after "Category: "
function lookupItem(name) {
  try {
    const q = normalize_(name);
    if (!q) return {ok:true, low:null, high:null};

    const sh = _ss().getSheetByName(SHEET_GUIDE);
    if (!sh) return {ok:false, error:`Sheet "${SHEET_GUIDE}" not found`, low:null, high:null};

    const vals = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), 4).getValues();
    let exact=null, starts=null, contains=null;

    for (const [cat, item, low, high] of vals) {
      const full = normalize_(item);
      const bare = normalize_(item.replace(/^[^:]+:\s*/, ""));
      const lo = Number(low||0), hi = Number(high||0);

      if (full === q || bare === q) { exact = {low:lo, high:hi}; break; }
      if (!starts && (full.startsWith(q) || bare.startsWith(q))) starts = {low:lo, high:hi};
      if (!contains && (full.includes(q) || bare.includes(q))) contains = {low:lo, high:hi};
    }
    const pick = exact || starts || contains || {low:null, high:null};
    return {ok:true, ...pick};
  } catch (e) {
    return {ok:false, error: String(e), low:null, high:null};
  }
}

// Compute FMV (midpoint × factor, clamped)
function computeFMV(low, high, condition) {
  if (!(low>0) || !(high>0)) return null;
  const mid = (low + high)/2;
  const f =
    condition === "Poor"      ? 0.8 :
    condition === "Fair"      ? 0.9 :
    condition === "Excellent" ? 1.1 : 1.0; // default Good/blank = 1.0
  const raw = mid * f;
  return Math.max(low, Math.min(high, raw));
}

// Ensure charity exists (adds if new). Returns {name, address}
function ensureCharity(name, address) {
  if (!name) return {name:"", address:""};
  const sh = _ss().getSheetByName(SHEET_CHAR);
  if (!sh) throw new Error(`Sheet "${SHEET_CHAR}" missing.`);
  const vals = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), 2).getValues();
  for (const [n, a] of vals) {
    if (normalize_(n) === normalize_(name)) return {name:n, address:a || ""};
  }
  sh.appendRow([name, address || ""]);
  return {name, address: address || ""};
}

// Format "MMMM d, yyyy"
function fmtDate(dt) {
  const d = typeof dt === "string" ? new Date(dt) : dt;
  if (Object.prototype.toString.call(d) !== "[object Date]" || isNaN(d)) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMMM d, yyyy");
}

/* ---------- Submit endpoints ---------- */

// payload: {org, newOrg, newAddr, dateISO, amount, method}
function submitCash(payload) {
  const sh = _ss().getSheetByName(SHEET_LOG);
  if (!sh) throw new Error(`Sheet "${SHEET_LOG}" missing.`);

  const chosen = (payload.newOrg && payload.newOrg.trim())
    ? ensureCharity(payload.newOrg, payload.newAddr)
    : ensureCharity(payload.org, "");

  const row = [
    chosen.name,
    chosen.address,
    fmtDate(payload.dateISO),
    "Money",
    "Cash Donation",
    "",                      // Description blank to match log schema
    Number(payload.amount||0)
  ];
  sh.appendRow(row);
  return {ok:true};
}

// payload: {org, newOrg, newAddr, dateISO, lines:[{item, condition, qty, override, note}]}
function submitItems(payload) {
  const sh = _ss().getSheetByName(SHEET_LOG);
  if (!sh) throw new Error(`Sheet "${SHEET_LOG}" missing.`);

  const chosen = (payload.newOrg && payload.newOrg.trim())
    ? ensureCharity(payload.newOrg, payload.newAddr)
    : ensureCharity(payload.org, "");

  const dateOut = fmtDate(payload.dateISO);
  let count = 0;

  (payload.lines || []).forEach(line => {
    if (!line.item) return;

    const res = lookupItem(line.item);
    if (res.ok === false) throw new Error(res.error || "Lookup failed.");
    const {low, high} = res;

    const fmv = (line.override && Number(line.override) > 0)
      ? Number(line.override)
      : computeFMV(low, high, line.condition || "");

    const qty   = Number(line.qty || 1) || 1;
    const total = fmv ? fmv * qty : 0;
    if (!(total > 0)) return;

    const desc = line.item +
      (line.condition ? ` (${line.condition})` : "") +
      (qty !== 1 ? ` x${qty}` : "") +
      (line.note ? ` – ${line.note}` : "");

    sh.appendRow([chosen.name, chosen.address, dateOut, "Item", "Non-Cash Donation", desc, total]);
    count++;
  });

  return {ok:true, lines:count};
}
