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
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Helpers ---------- */
function normalize_(s){ return (s||"").toString().trim().toLowerCase(); }

function listCharities() {
  try {
    const sh = _ss().getSheetByName(SHEET_CHAR);
    if (!sh) return {ok:false,error:`Sheet "${SHEET_CHAR}" not found`, data:[]};
    const lr = sh.getLastRow();
    const vals = lr < 2 ? [] : sh.getRange(2,1, lr-1, 2).getValues();
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
    const lr = sh.getLastRow();
    const vals = lr < 2 ? [] : sh.getRange(2,1, lr-1, 4).getValues();
    const data = vals
      .filter(r => r[1])
      .map(([cat, item, low, high]) => ({
        item, low: Number(low || 0), high: Number(high || 0), category: cat
      }));
    return { ok: true, rawData: data };
  } catch (e) {
    return {ok:false, error: String(e), data:[]};
  }
}

// Ensure charity exists (adds if new). Returns {name, address}
function ensureCharity(name, address) {
  if (!name) return {name:"", address:""};
  const sh = _ss().getSheetByName(SHEET_CHAR);
  if (!sh) throw new Error(`Sheet "${SHEET_CHAR}" missing.`);
  const lr = sh.getLastRow();
  const vals = lr < 2 ? [] : sh.getRange(2,1, lr-1, 2).getValues();
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
    fmtDate(payload.dateISO),
    "Money",
    payload.method,
    "",                      // Description blank to match log schema
    Number(payload.amount||0), 
    payload.note || ""
  ];
  sh.appendRow(row);
  return {ok:true};
}

// payload: {org, newOrg, newAddr, dateISO, lines:[{item, condition, qty, override, note}]}
function submitItems(payload) {
  try {
    const ss = _ss();
    const sheet = ss.getSheetByName(SHEET_LOG); // Replace with your actual donations sheet constant
        
    // Determine the Organization Name
    const chosen = (payload.newOrg && payload.newOrg.trim())
    ? ensureCharity(payload.newOrg, payload.newAddr)
    : ensureCharity(payload.org, "");

    const dateOut = fmtDate(payload.dateISO);
    
    // Prepare the rows to be appended
    const rows = payload.lines.map(line => [
      chosen.name,
      dateOut,
      line.item,
      line.condition,
      line.qty,
      line.qty * (line.override || line.fmvAtTime), // Use override if exists, otherwise the FMV found
      line.note || ""
    ]);

    // Append rows to the sheet
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { ok: true, lines: rows.length };
  } catch (e) {
    throw new Error("Failed to submit items: " + e.toString());
  }
}