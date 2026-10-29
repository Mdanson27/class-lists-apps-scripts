/***** ===== CLASS LISTS — TERM AUTOMATION (CLASS LISTS / STRUCTURE) =====
 * v1.7-structure-hardened (NO-GREEN + SAFE student-end detection + NO formula spread)
 *
 * PERMANENT FIXES ADDED IN THIS UPDATE:
 * ✅ BALANCE formula is applied ONLY to REAL student rows (rows 2 → last student row), NOT “row before TOTALS”.
 * ✅ Any BALANCE formulas that previously spilled into blank rows ABOVE TOTALS are cleaned safely.
 * ✅ We NEVER insert rows up to TOTALS row (no “gap fill” that creates more rows to fill formulas into).
 * ✅ Total/format rules remain: NO conditional formatting on class sheets; totals row stays neutral grey.
 *
 * SAFE EDITOR-RUNNABLE HELPERS ADDED:
 * - SAFE_StructureFix_OneSheetByName(sheetName)
 * - SAFE_StructureFix_AllClassSheets()
 * - SAFE_StructureFix_SeniorSheetsOnly()   (targets sheets starting with "Senior")
 *
 * DEPENDS ON:
 * - Config.gs (CL)
 * - Utils.gs: ss_(), toast_(), getClassSheets_(), getHeaderIndexMap_(), getStudentRegion_(), findTotalsRowWide_()
 *   (Optional) getHeaderIndexMapUpper_(), isDefaultersSheetName_(), _flatUpper_()
 ************************************************************/

/* =========================
 * 0) SMALL SAFE HELPERS
 * ========================= */

function CLS_isDefaultersSheetName_(name) {
  try { if (typeof isDefaultersSheetName_ === "function") return isDefaultersSheetName_(name); } catch (_) {}
  const n = String(name || "").trim().toLowerCase();
  const suffix = String(CL?.DEFS_SUFFIX || " Defaulters").trim().toLowerCase();
  if (suffix && n.endsWith(suffix)) return true;
  return /defauldt?efaulters$/i.test(name || "") || /defaulters$/i.test(name || "");
}

/** Best-effort header column getter supporting optional CL.H_ALIASES + upper-map */
function CLS_col_(sh, canonicalHeader) {
  const exact = getHeaderIndexMap_(sh) || {};
  if (exact[canonicalHeader]) return exact[canonicalHeader];

  const aliases = (CL?.H_ALIASES && typeof CL.H_ALIASES === "object") ? CL.H_ALIASES : null;

  let upper = null;
  try { if (typeof getHeaderIndexMapUpper_ === "function") upper = getHeaderIndexMapUpper_(sh); } catch (_) {}

  if (aliases) {
    const key = Object.keys(CL.H).find(k => CL.H[k] === canonicalHeader);
    const list = key && Array.isArray(aliases[key]) ? aliases[key] : null;
    if (list && list.length) {
      for (const h of list) {
        if (exact[h]) return exact[h];
        try {
          if (upper && typeof _flatUpper_ === "function" && upper[_flatUpper_(h)]) return upper[_flatUpper_(h)];
        } catch (_) {}
      }
    }
  }

  try {
    if (upper && typeof _flatUpper_ === "function" && upper[_flatUpper_(canonicalHeader)]) return upper[_flatUpper_(canonicalHeader)];
  } catch (_) {}

  return null;
}

/** Remove ALL conditional formatting rules from a class sheet (defaulters keep theirs). */
function CLS_removeConditionalFormattingFromClassSheet_(sh) {
  try {
    if (!sh) return;
    if (CLS_isDefaultersSheetName_(sh.getName())) return;
    const rules = sh.getConditionalFormatRules();
    if (rules && rules.length) sh.setConditionalFormatRules([]);
  } catch (_) {}
}

/** Normalize display text safely */
function CLS_norm_(v) {
  return String(v == null ? "" : v).trim().replace(/\s+/g, " ");
}

/**
 * Find the LAST REAL student row (so we don't fill formulas into empty space above TOTALS).
 * Strategy:
 * - Look in a "student key" column that should be filled for real students:
 *   Prefer NAME; else NO.; else column A.
 * - Scan bottom-up within rows 2..(totalsRow-1 OR lastRow) and return the last row whose key cell is non-empty.
 */
function CLS_findLastStudentRow_(sh, totalsRow, cNAME, cNO) {
  const lastRow = Math.max(1, sh.getLastRow());
  const hardEnd = totalsRow ? Math.max(2, totalsRow - 1) : Math.max(2, lastRow);
  if (hardEnd < 2) return 1;

  const keyCol = cNAME || cNO || 1; // NAME preferred, then NO., else col A
  const numRows = hardEnd - 1;

  // Use display values to catch numbers/text consistently and avoid formula objects.
  const vals = sh.getRange(2, keyCol, numRows, 1).getDisplayValues().flat();

  for (let i = vals.length - 1; i >= 0; i--) {
    const t = CLS_norm_(vals[i]);
    if (t) return 2 + i; // absolute row
  }
  return 1; // no student rows found
}

/**
 * Clear spilled BALANCE formulas in blank rows between (lastStudentRow+1 .. totalsRow-1),
 * but ONLY where the row is truly blank in student key columns (NAME/NO).
 */
function CLS_clearSpilledBalanceFormulas_(sh, cBAL, totalsRow, lastStudentRow, cNAME, cNO) {
  if (!totalsRow || totalsRow <= 3) return;
  if (!cBAL) return;

  const start = Math.max(2, (lastStudentRow || 1) + 1);
  const end = Math.max(1, totalsRow - 1);
  if (start > end) return;

  const rows = end - start + 1;

  const keyCol1 = cNAME || 1;
  const keyCol2 = cNO || null;

  const key1 = sh.getRange(start, keyCol1, rows, 1).getDisplayValues().flat();
  const key2 = keyCol2 ? sh.getRange(start, keyCol2, rows, 1).getDisplayValues().flat() : null;

  const balFormulas = sh.getRange(start, cBAL, rows, 1).getFormulas();

  // Build list of rows to clear (only rows that look blank AND have a formula)
  const toClearA1 = [];
  for (let i = 0; i < rows; i++) {
    const k1 = CLS_norm_(key1[i]);
    const k2 = key2 ? CLS_norm_(key2[i]) : "";
    const hasKey = !!(k1 || k2);

    const hasFormula = !!(balFormulas[i] && balFormulas[i][0]);
    if (!hasKey && hasFormula) {
      const r = start + i;
      toClearA1.push(`R${r}C${cBAL}`);
    }
  }

  if (!toClearA1.length) return;

  // Clear each safely (small count usually). Avoid clearing values in non-empty rows.
  toClearA1.forEach(a1 => {
    try {
      sh.getRange(a1).clearContent();
    } catch (_) {}
  });
}

/* =========================
 * 1) STRUCTURE ENSURE
 * ========================= */

function ensureStructureAll() {
  const sheets = getClassSheets_();
  const errors = [];

  sheets.forEach(sh => {
    try {
      ensureStructureOne_(sh);
    } catch (e) {
      errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    toast_(`Structure ensured with ${errors.length} issue(s). Check logs.`, "Ensure Structure", 8);
    Logger.log("ensureStructureAll() issues:\n" + errors.join("\n"));
    return;
  }

  toast_("Structure ensured for all class sheets.", "Done", 5);
}

/**
 * Ensures:
 * - LTB column exists immediately before FULL FEE
 * - BALANCE formula exists in BALANCE col for REAL student rows ONLY
 *
 * Safety:
 * - Never writes below TOTALS
 * - Never fills formulas into blank space above TOTALS
 * - If TOTALS missing, uses last used content area
 * - Does NOT copy template formatting (prevents “green spread”)
 */
function ensureStructureOne_(sh) {
  if (!sh) throw new Error("No sheet provided.");

  // Never enforce structure on defaulters sheets
  if (CLS_isDefaultersSheetName_(sh.getName())) return;

  // Remove conditional formatting from class sheets
  CLS_removeConditionalFormattingFromClassSheet_(sh);

  const cFEE0 = CLS_col_(sh, CL.H.FEE);
  if (!cFEE0) throw new Error(`Missing header "${CL.H.FEE}". (Check spelling/case in row 1)`);

  const cLTB0 = CLS_col_(sh, CL.H.LTB);

  // Insert LTB immediately before FULL FEE if missing
  if (!cLTB0) {
    sh.insertColumnBefore(cFEE0);
    sh.getRange(1, cFEE0).setValue(CL.H.LTB).setFontWeight("bold");
  }

  // Re-evaluate columns after insertion
  const cLTB  = CLS_col_(sh, CL.H.LTB);
  const cFEE  = CLS_col_(sh, CL.H.FEE);
  const cPAID = CLS_col_(sh, CL.H.PAID);
  const cBAL  = CLS_col_(sh, CL.H.BAL);
  const cNAME = CLS_col_(sh, CL.H.NAME); // used to detect last student
  const cNO   = CLS_col_(sh, CL.H.NO);   // fallback student key

  if (!cLTB || !cFEE || !cPAID || !cBAL) {
    throw new Error(`Missing one of: "${CL.H.LTB}", "${CL.H.FEE}", "${CL.H.PAID}", "${CL.H.BAL}".`);
  }

  // Find totals row (if present)
  const region = (typeof getStudentRegion_ === "function") ? getStudentRegion_(sh) : null;
  const totalsRow =
    (region && region.totalsRow) ? region.totalsRow :
    (typeof findTotalsRowWide_ === "function" ? findTotalsRowWide_(sh) : null);

  // Determine LAST REAL student row (prevents formula spread into blank rows)
  const lastStudentRow = CLS_findLastStudentRow_(sh, totalsRow, cNAME, cNO);

  // If no students, nothing to fill
  if (!lastStudentRow || lastStudentRow < 2) {
    // Still clean any spilled BAL formulas above TOTALS if they exist
    if (totalsRow) CLS_clearSpilledBalanceFormulas_(sh, cBAL, totalsRow, 1, cNAME, cNO);
    return;
  }

  // Compute safe end row: never beyond totalsRow-1
  const hardEnd = totalsRow ? Math.max(2, totalsRow - 1) : Math.max(2, sh.getLastRow());
  const endRow = Math.min(lastStudentRow, hardEnd);
  if (endRow < 2) return;

  // BAL = (LTB + FEE) - PAID, blank if both LTB and FEE are 0/blank
  const dLTB  = cLTB  - cBAL;
  const dFEE  = cFEE  - cBAL;
  const dPAID = cPAID - cBAL;

  const fR1C1 =
    `=IF((N(RC[${dLTB}])+N(RC[${dFEE}]))=0,"", N(RC[${dLTB}])+N(RC[${dFEE}]) - N(RC[${dPAID}]))`;

  // Apply ONLY to rows 2..endRow (REAL students only)
  sh.getRange(2, cBAL, endRow - 1, 1).setFormulaR1C1(fR1C1);

  // Clean any spilled BALANCE formulas in blank rows up to TOTALS
  if (totalsRow) CLS_clearSpilledBalanceFormulas_(sh, cBAL, totalsRow, endRow, cNAME, cNO);
}

/* =========================
 * 2) SAFE CLEANUP HELPERS
 * ========================= */

function trimBelowTotals_One(shOrName) {
  const ss = ss_();
  const sh = (typeof shOrName === "string") ? ss.getSheetByName(shOrName) : shOrName;
  if (!sh) return;

  if (CLS_isDefaultersSheetName_(sh.getName())) return;

  const tRow = (typeof findTotalsRowWide_ === "function") ? findTotalsRowWide_(sh) : null;
  if (!tRow) return;

  // NEVER delete physical rows from class sheets.
  // Rows below TOTALS may be reused later; leave the grid structure intact.
  const extra = sh.getMaxRows() - tRow;
  if (extra > 0) {
    const lastCol = Math.max(1, sh.getLastColumn());
    sh.getRange(tRow + 1, 1, extra, lastCol).clearContent();
  }
}

function trimBelowTotals_All() {
  getClassSheets_().forEach(sh => {
    try { trimBelowTotals_One(sh); } catch (e) {}
  });
  toast_("Cleaned rows below TOTALS on all class sheets.", "Clean", 6);
}

/**
 * Ensure sheet has at least neededLastRow rows.
 * IMPORTANT: we keep this, but it is NO LONGER used to "reach" totalsRow-1 in ensureStructureOne_.
 * (We do not want row-insertion to create more blank rows for formulas.)
 */
function ensureRows_(sh, neededLastRow) {
  const needed = Number(neededLastRow) || 0;
  if (needed <= 0) return;

  const currentMax = sh.getMaxRows();
  const gap = needed - currentMax;
  if (gap <= 0) return;

  sh.insertRowsAfter(currentMax, gap);

  try {
    const lastCol = Math.max(1, sh.getLastColumn());
    sh.getRange(currentMax + 1, 1, gap, lastCol).clearFormat();
  } catch (_) {}
}

/* =========================
 * 2B) SAFE EDITOR-RUNNABLE REPAIR HELPERS (NEW)
 * ========================= */

/**
 * Run from editor to safely fix ONE sheet by name.
 * - Rebuilds structure (LTB if missing + Balance formulas only on real students)
 * - Removes any spilled balance formulas in blank rows above totals
 */
function SAFE_StructureFix_OneSheetByName(sheetName) {
  const ss = ss_();
  const sh = ss.getSheetByName(String(sheetName || ""));
  if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);
  ensureStructureOne_(sh);
  toast_(`Fixed structure safely: ${sh.getName()}`, "SAFE Fix", 6);
}

/** Run from editor to safely fix ALL class sheets */
function SAFE_StructureFix_AllClassSheets() {
  const sheets = getClassSheets_();
  const errs = [];
  sheets.forEach(sh => {
    try { ensureStructureOne_(sh); }
    catch (e) { errs.push(`${sh.getName()}: ${e && e.message ? e.message : e}`); }
  });
  if (errs.length) {
    Logger.log("SAFE_StructureFix_AllClassSheets() issues:\n" + errs.join("\n"));
    toast_(`SAFE fix finished with ${errs.length} issue(s).`, "SAFE Fix", 8);
    return;
  }
  toast_("SAFE fix complete for all class sheets.", "SAFE Fix", 6);
}

/** Run from editor to fix only sheets whose names start with "Senior" (common problem area) */
function SAFE_StructureFix_SeniorSheetsOnly() {
  const ss = ss_();
  const suf = String(CL?.CLASS_SUFFIX || " list").toLowerCase();
  const sheets = ss.getSheets().filter(sh => {
    const n = sh.getName();
    if (CLS_isDefaultersSheetName_(n)) return false;
    const nl = String(n).toLowerCase();
    return nl.startsWith("senior") && nl.endsWith(suf);
  });

  const errs = [];
  sheets.forEach(sh => {
    try { ensureStructureOne_(sh); }
    catch (e) { errs.push(`${sh.getName()}: ${e && e.message ? e.message : e}`); }
  });

  if (errs.length) {
    Logger.log("SAFE_StructureFix_SeniorSheetsOnly() issues:\n" + errs.join("\n"));
    toast_(`Senior SAFE fix finished with ${errs.length} issue(s).`, "SAFE Fix", 8);
    return;
  }
  toast_("Senior SAFE fix complete.", "SAFE Fix", 6);
}

/* =========================
 * 3) FORMAT / ORGANIZE (VISUAL ONLY)
 * ========================= */

function formatWorkbookLayout() {
  formatAllClassSheets_();
  formatAllDefaultersSheets_();
  toast_("Formatting applied to class lists + defaulters.", "Format", 6);
}

function formatAllClassSheets_() {
  const sheets = getClassSheets_();
  const errors = [];

  sheets.forEach(sh => {
    try {
      formatOneClassSheet_(sh);
    } catch (e) {
      errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    Logger.log("formatAllClassSheets_() issues:\n" + errors.join("\n"));
    toast_(`Formatted with ${errors.length} issue(s). Check logs.`, "Format", 8);
  }
}

function formatAllDefaultersSheets_() {
  const ss = ss_();
  const errors = [];

  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (!CLS_isDefaultersSheetName_(name)) return;

    try {
      formatOneDefaultersSheet_(sh);
    } catch (e) {
      errors.push(`${name}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    Logger.log("formatAllDefaultersSheets_() issues:\n" + errors.join("\n"));
    toast_(`Defaulters formatted with ${errors.length} issue(s). Check logs.`, "Format", 8);
  }
}

function formatOneClassSheet_(sh) {
  if (!sh) throw new Error("No sheet provided.");
  if (CLS_isDefaultersSheetName_(sh.getName())) return;

  try { ensureStructureOne_(sh); } catch (_) {}
  CLS_removeConditionalFormattingFromClassSheet_(sh);

  sh.setFrozenRows(1);

  const headerColEnd = _lastHeaderCol_(sh);

  const region = (typeof getStudentRegion_ === "function") ? getStudentRegion_(sh) : null;
  const tRow = (region && region.totalsRow) ? region.totalsRow : (typeof findTotalsRowWide_ === "function" ? findTotalsRowWide_(sh) : null);

  const endRow = tRow ? tRow : Math.max(2, _lastContentRowInCols_(sh, headerColEnd));

  _applyHeaderStyle_(sh, headerColEnd);
  _applyColumnWidths_(sh);
  _applyTextAlignments_(sh, endRow);
  _applyMoneyFormats_(sh, endRow);
  _applyRowHeights_(sh, 2, endRow);
  _applyTableBorders_(sh, 1, 1, endRow, headerColEnd);

  if (tRow) _applyTotalsRowStyle_(sh, tRow, headerColEnd);
}

function formatOneDefaultersSheet_(sh) {
  if (!sh) throw new Error("No sheet provided.");

  sh.setFrozenRows(1);

  const headerColEnd = _lastHeaderCol_(sh);
  const tRow = (typeof findTotalsRowWide_ === "function") ? findTotalsRowWide_(sh) : null;
  const endRow = tRow ? tRow : Math.max(2, _lastContentRowInCols_(sh, headerColEnd));

  try {
    sh.getRange(1, 1, Math.max(1, endRow), Math.max(1, headerColEnd)).clearFormat();
  } catch (_) {}

  _applyHeaderStyle_(sh, headerColEnd);
  _applyColumnWidths_(sh);
  _applyTextAlignments_(sh, endRow);
  _applyMoneyFormats_(sh, endRow);
  _applyRowHeights_(sh, 2, endRow);

  _applyTableBorders_(sh, 1, 1, endRow, headerColEnd);
  if (tRow) _applyTotalsRowStyle_(sh, tRow, headerColEnd);
}

/* =========================
 * Internal formatting helpers
 * ========================= */

function _applyHeaderStyle_(sh, headerColEnd) {
  const header = sh.getRange(1, 1, 1, headerColEnd);

  header
    .setBackground("#E6E6E6")
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  try { sh.setRowHeight(1, 34); } catch (_) {}

  header.setBorder(null, null, true, null, null, null, "#666666", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function _applyTotalsRowStyle_(sh, totalsRow, colEnd) {
  const rng = sh.getRange(totalsRow, 1, 1, colEnd);

  rng
    .setBackground("#D9D9D9")
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  rng.setBorder(true, null, null, null, null, null, "#666666", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  try { sh.setRowHeight(totalsRow, 24); } catch (_) {}
}

function _applyTableBorders_(sh, row, col, numRows, numCols) {
  if (numRows <= 0 || numCols <= 0) return;
  const rng = sh.getRange(row, col, numRows, numCols);

  rng.setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);
  _setOuterBorder_(rng, "#555555");
}

function _setOuterBorder_(range, color) {
  const sheet = range.getSheet();
  const r = range.getRow();
  const c = range.getColumn();
  const nr = range.getNumRows();
  const nc = range.getNumColumns();

  sheet.getRange(r, c, 1, nc)
    .setBorder(true, null, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(r + nr - 1, c, 1, nc)
    .setBorder(null, null, true, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(r, c, nr, 1)
    .setBorder(null, true, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(r, c + nc - 1, nr, 1)
    .setBorder(null, null, null, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function _applyMoneyFormats_(sh, endRow) {
  const map = getHeaderIndexMap_(sh);

  const moneyHeaders = [CL.H.LTB, CL.H.FEE, CL.H.PAID, CL.H.BAL, CL.H.CAUT];
  const cols = moneyHeaders.map(h => map[h]).filter(Boolean);
  if (!cols.length) return;

  const last = Math.max(2, Number(endRow || sh.getLastRow()) || 2);
  const numRows = Math.max(1, last - 1);

  cols.forEach(col => {
    const rng = sh.getRange(2, col, numRows, 1);
    rng
      .setNumberFormat("#,##0")
      .setHorizontalAlignment("right")
      .setFontFamily("Arial")
      .setFontSize(10);
  });
}

function _applyTextAlignments_(sh, endRow) {
  const map = getHeaderIndexMap_(sh);
  const last = Math.max(2, Number(endRow || sh.getLastRow()) || 2);
  const numRows = Math.max(1, last - 1);

  const textHeaders = [CL.H.DATE, CL.H.NAME, CL.H.SEX, CL.H.NAT, CL.H.DB, CL.H.TEL, CL.H.REM, CL.H.NO];
  textHeaders.forEach(h => {
    const c = map[h];
    if (!c) return;
    sh.getRange(2, c, numRows, 1)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setFontFamily("Arial")
      .setFontSize(10);
  });
}

function _applyRowHeights_(sh, startRow, endRow) {
  const s = Math.max(2, startRow || 2);
  const e = Math.max(s, endRow || s);
  for (let r = s; r <= e; r++) {
    try { sh.setRowHeight(r, 22); } catch (_) {}
  }
}

function _applyColumnWidths_(sh) {
  const map = getHeaderIndexMap_(sh);

  const widths = {};
  widths[CL.H.NO]   = 50;
  widths[CL.H.DATE] = 120;
  widths[CL.H.NAME] = 260;
  widths[CL.H.SEX]  = 55;
  widths[CL.H.NAT]  = 130;
  widths[CL.H.DB]   = 110;

  widths[CL.H.LTB]  = 145;
  widths[CL.H.FEE]  = 120;
  widths[CL.H.PAID] = 120;
  widths[CL.H.BAL]  = 130;
  widths[CL.H.CAUT] = 110;

  widths[CL.H.TEL]  = 130;
  widths[CL.H.REM]  = 360;

  Object.keys(widths).forEach(h => {
    const c = map[h];
    if (!c) return;
    try { sh.setColumnWidth(c, widths[h]); } catch (_) {}
  });
}

function _lastHeaderCol_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol <= 1) return 1;

  const vals = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let end = 1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i] || "").trim() !== "") end = i + 1;
  }
  return end;
}

function _lastContentRowInCols_(sh, colEnd) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;

  const rows = Math.max(1, lastRow);
  const vals = sh.getRange(1, 1, rows, colEnd).getValues();

  let last = 1;
  for (let r = 0; r < vals.length; r++) {
    const rowHas = vals[r].some(v => String(v || "").trim() !== "");
    if (rowHas) last = r + 1;
  }
  return last;
}