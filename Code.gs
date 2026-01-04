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

function buildPrintout() {
  const ss = _ss();
  const dataSheet = ss.getSheetByName(SHEET_LOG);
  if (!dataSheet) throw new Error(`Sheet "${SHEET_LOG}" not found.`);

  let rpt = ss.getSheetByName('Printout');
  if (!rpt) rpt = ss.insertSheet('Printout');
  rpt.clear();
  rpt.setHiddenGridlines(true);

  rpt.setColumnWidths(1, 1, 28);
  rpt.setColumnWidths(2, 1, 36);
  rpt.setColumnWidths(3, 1, 65);
  rpt.setColumnWidths(4, 1, 16);

  const year = new Date().getFullYear();
  const name = ss.getOwner() ? ss.getOwner().getEmail() : 'YOUR NAME';

  rpt.getRange('A1:D1').merge().setValue(`${year} Tax Year`).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  rpt.getRange('A2:D2').merge().setValue('YTD Charitable Deductions').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  rpt.getRange('A3:D3').merge().setValue(name).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

  let row = 5;

  const values = dataSheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));

  function isCash(r){ return String(r[idx['IRS Donation Type Classification']] || '').toLowerCase().includes('cash'); }
  function isNonCash(r){ return String(r[idx['IRS Donation Type Classification']] || '').toLowerCase().includes('non'); }

  function fmtDateForReport(d){
    if (!(d instanceof Date)) return d;
    return Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }

  function writeSection(title, filterFn) {
    rpt.getRange(row,1,1,4).merge().setValue(title).setFontWeight('bold');
    row++;

    rpt.getRange(row,1).setValue('Charity Name / Donated Date').setFontWeight('bold');
    rpt.getRange(row,2).setValue('Charity Address').setFontWeight('bold');
    rpt.getRange(row,3).setValue('Donation Description').setFontWeight('bold');
    rpt.getRange(row,4).setValue('Donation Amount').setFontWeight('bold').setHorizontalAlignment('right');
    row++;

    const rows = values.filter(filterFn).sort((a,b)=>{
      const ca = String(a[idx['Charity']]||'');
      const cb = String(b[idx['Charity']]||'');
      if (ca !== cb) return ca.localeCompare(cb);
      return new Date(a[idx['Date']]) - new Date(b[idx['Date']]);
    });

    let total = 0;

    rows.forEach(r=>{
      const charity = r[idx['Charity']] || '';
      const addr = r[idx['Charity Address']] || '';
      const date = fmtDateForReport(r[idx['Date']]);
      const desc = r[idx['Description']] || '';
      const amt = Number(r[idx['Donation Value in $']] || 0);

      rpt.getRange(row,1).setValue(`${charity}\n${date}`).setWrap(true);
      rpt.getRange(row,2).setValue(addr).setWrap(true);
      rpt.getRange(row,3).setValue(desc).setWrap(true);
      rpt.getRange(row,4).setValue(amt).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');

      total += amt;
      row++;
    });

    rpt.getRange(row,3).setValue('Subtotal :').setFontWeight('bold').setHorizontalAlignment('right');
    rpt.getRange(row,4).setValue(total).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
    row++;

    return total;
  }

  const nonCashTotal = writeSection('Non-Cash Donations', isNonCash);
  rpt.getRange(row,3).setValue('Total Non-Cash Donations:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(nonCashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
  row += 2;

  const cashTotal = writeSection('Cash Donations', isCash);
  rpt.getRange(row,3).setValue('Total Cash Donations:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(cashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
  row += 2;

  rpt.getRange(row,3).setValue('Grand Total:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(nonCashTotal + cashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');

  rpt.getRange(6,1,row,4).setVerticalAlignment('top');
  rpt.getDataRange().setFontFamily('Arial').setFontSize(10);
}
