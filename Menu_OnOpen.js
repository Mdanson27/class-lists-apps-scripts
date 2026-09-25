/***** ===== CLASS LISTS — TERM AUTOMATION (MENUS / onOpen) =====
 * Single source of truth for menus.
 *
 * UPDATE (as you requested):
 * ✅ Adds a helper to ensure the "Leavers – Debtors" sheet is properly organized:
 *    - Grey header row
 *    - Bold header text
 *    - Frozen header
 *    - Filter on header
 *    - Useful column widths
 *    - Money columns formatted with #,##0
 *
 * NOTE:
 * - This does NOT move students. It only ensures the Debtors sheet is neat/ready.
 * - Moving/restoring is still done by: recordLeaverQuick / restoreLeaverToClass / markLeaverSettled
 ************************************************************/

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // ===== Main menu =====
  ui.createMenu("CLASS LISTS — TERM")
    .addItem('Ensure "LAST TERM BALANCE" + Balance formulas', "ensureStructureAll")
    .addItem("Rebuild all Defaulters sheets", "rebuildAllDefaulters_QUERY")
    .addSeparator()
    .addItem("Fix totals (Expected/Paid/Balance)", "fixTotalsAllSheets")
    .addItem("Fix totals — use fixed positions", "fixTotals_UseFixedPositions")
    .addSeparator()
    .addItem("Refresh SUMMARIES (Totals + Debtors + Headcounts)", "refreshSummaries")
    .addSeparator()
    .addItem("Start a New Term… (Archive + Carry balances)", "startNewTerm")
    .addItem("Start a New Year… (Promote + New Term)", "startNewYear")
    .addSeparator()
    .addItem('Prepare / Format "Leavers – Debtors" sheet', "prepareLeaversDebtorsSheet_")
    .addItem("Record Leaver (→ Debtors)…", "recordLeaverQuick")
    .addItem("Restore Debtor to Class…", "restoreLeaverToClass")
    .addItem("Mark Debtor Settled…", "markLeaverSettled")
    .addToUi();

  // ===== Admin menu =====
  ui.createMenu("ADMIN")
    .addItem("Lock + hide protected sheets (apply now)", "applyAdminLocks_")
    .addItem("Unhide protected sheets (admins only)", "unhideForAdmins_")
    .addSeparator()
    .addItem("Fix green below TOTALS (once)", "fixTotalsFormattingAll_")
    .addToUi();

  // Best-effort: enforce hidden state on open (does not edit data)
  try {
    autoHideForNonAdmins_();
  } catch (_) {}

  // Best-effort: keep Leavers – Debtors tidy (safe formatting only)
  try {
    prepareLeaversDebtorsSheet_();
  } catch (_) {}
}

/** Optional: ensures menus appear after installing the script */
function onInstall(e) {
  onOpen(e);
}

/**
 * Ensures "Leavers – Debtors" exists + formats header nicely.
 * SAFE: only touches that sheet + only formatting/header row.
 */
function prepareLeaversDebtorsSheet_() {
  // Ensure sheet exists + header exists (your Leavers_Debtors.gs function)
  const sh = ensureLeaversSheet_();
  if (!sh) return;

  const lastCol = Math.max(sh.getLastColumn(), CL.LEAVERS.HEAD.length);

  // Ensure header row exactly matches config (keeps everything consistent)
  sh.getRange(1, 1, 1, CL.LEAVERS.HEAD.length).setValues([CL.LEAVERS.HEAD]);

  // Freeze header
  try { sh.setFrozenRows(1); } catch (_) {}

  // Header style: grey + bold + centered
  const header = sh.getRange(1, 1, 1, lastCol);
  header
    .setBackground("#D9D9D9")
    .setFontWeight("bold")
    .setFontColor("#000000")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Add filter on header row (only if none)
  try {
    if (!sh.getFilter()) header.createFilter();
  } catch (_) {}

  // Column widths (nice defaults; adjust anytime)
  try {
    // Core class columns (A..L)
    sh.setColumnWidth(1, 60);   // NO.
    sh.setColumnWidth(2, 130);  // Date of Admission
    sh.setColumnWidth(3, 220);  // Student Name
    sh.setColumnWidth(4, 60);   // SEX
    sh.setColumnWidth(5, 130);  // NATIONALITY
    sh.setColumnWidth(6, 140);  // DAY/BOARDING
    sh.setColumnWidth(7, 140);  // LAST TERM BALANCE
    sh.setColumnWidth(8, 110);  // FULL FEE
    sh.setColumnWidth(9, 110);  // P (Paid)
    sh.setColumnWidth(10, 120); // BALANCE
    sh.setColumnWidth(11, 110); // CAUTION
    sh.setColumnWidth(12, 110); // TEL
  } catch (_) {}

  // Money number formats on LTB / FULL FEE / PAID / BALANCE (if those headers exist)
  try {
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const colByName = (name) => {
      const i = headers.findIndex(h => h.toLowerCase() === String(name).toLowerCase());
      return i === -1 ? null : i + 1;
    };

    const cLTB  = colByName(CL.H.LTB);
    const cFEE  = colByName(CL.H.FEE);
    const cPAID = colByName(CL.H.PAID);
    const cBAL  = colByName(CL.H.BAL);

    const lastRow = sh.getMaxRows();
    const fmt = "#,##0";

    if (cLTB)  sh.getRange(2, cLTB,  lastRow - 1, 1).setNumberFormat(fmt);
    if (cFEE)  sh.getRange(2, cFEE,  lastRow - 1, 1).setNumberFormat(fmt);
    if (cPAID) sh.getRange(2, cPAID, lastRow - 1, 1).setNumberFormat(fmt);
    if (cBAL)  sh.getRange(2, cBAL,  lastRow - 1, 1).setNumberFormat(fmt);
  } catch (_) {}
}
