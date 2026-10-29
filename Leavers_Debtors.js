/***** ===== LEAVERS / DEBTORS (capture + restore + sheet layout) =====
 * v2.1-fixed-totals-no-row-shift
 *
 * CRITICAL POLICY FOR YOUR “FIXED TOTALS ROW” SYSTEM:
 * ✅ We DO NOT delete student rows from class lists.
 * ✅ We DO NOT insert rows above TOTALS.
 * ✅ “Remove student” = clear the student’s row content (row stays).
 * ✅ “Restore student” = fill the first EMPTY student row above TOTALS.
 *
 * This guarantees your TOTALS rows stay EXACTLY at the fixed positions you provided.
 *
 * RUN FROM EDITOR:
 *  ✅ prepareLeaversDebtorsSheetSafe()        -> keeps ALL existing data, fixes layout + formatting
 *  ✅ rebuildLeaversDebtorsSheetFresh()       -> deletes sheet + recreates it (WIPES ALL DATA)
 *  ✅ rebuildLeaversDebtorsSheetClearOnly()   -> clears all data in-place + rebuilds headers/format (WIPES ALL DATA)
 ************************************************************/

/** ==================== SETTINGS ==================== */
/**
 * Restoration policy:
 * Restoring a student is a MOVE, never a copy. After the class write is
 * verified, the matching row is deleted from Leavers – Debtors.
 */
const LEAVERS_RESTORE_BEHAVIOR = "DELETE";

/**
 * FIXED TOTALS ROWS (AUTHORITATIVE)
 * These are the exact totals row positions you provided.
 * Sorting/renumbering (nightly) and row-safe operations should respect these.
 */
const CLASS_TOTALS_ROWS = {
  // Nursery
  "Baby class list": 36,
  "Middle class list": 42,
  "Top Class list": 47,

  // Primary
  "Primary 1 list": 45,
  "Primary 2 list": 51,
  "Primary 3 list": 58,
  "Primary 4 list": 48,
  "Primary 5 list": 49,
  "Primary 6 list": 50,
  "Primary 7 list": 64,

  // Secondary
  "Senior 1 list": 74,
  "Senior 2 list": 107,
  "Senior 3 list": 91,
  "Senior 4 list": 113,
  "Senior 5 list": 86,
  "Senior 6 list": 97,
};

/** ==================== SAFE LOCAL WRAPPERS (no hard dependency on Utils.gs) ==================== */
function _ss_() { return SpreadsheetApp.getActive(); }

function _toast_(msg, title = "Info", seconds = 5) {
  try { _ss_().toast(String(msg), String(title), Number(seconds) || 5); } catch (_) {}
}

function _norm_(v) {
  return String(v ?? "")
    .replace(/[\u00A0]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function _flatUpper_(v) {
  return _norm_(v).replace(/\s+/g, "").toUpperCase();
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

/** Resolve Leavers sheet name by canonical + aliases (prevents duplicates due to dash differences) */
function _resolveLeaversSheetName_() {
  const ss = _ss_();
  const canonical = String(CL?.LEAVERS?.SHEET || "Leavers – Debtors");
  if (ss.getSheetByName(canonical)) return canonical;

  const aliases = (CL?.SHEET_ALIASES && Array.isArray(CL.SHEET_ALIASES.LEAVERS))
    ? CL.SHEET_ALIASES.LEAVERS
    : ["Leavers – Debtors", "Leavers - Debtors", "Leavers–Debtors", "Leavers—Debtors"];

  for (const nm of aliases) {
    if (ss.getSheetByName(String(nm))) return String(nm);
  }
  return canonical;
}

/** Header index map (prefer Utils if available) */
function _getHeaderIndexMapLocal_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_norm_);
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (!h) continue;
    if (map[h] == null) map[h] = i + 1;
  }
  return map;
}
function _getHeaderIndexMap_(sh) {
  try { if (typeof getHeaderIndexMap_ === "function") return getHeaderIndexMap_(sh); } catch (_) {}
  return _getHeaderIndexMapLocal_(sh);
}

/** Find last header column (for safe “clear row content” boundaries) */
function _lastHeaderCol_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const vals = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let end = 1;
  for (let i = 0; i < vals.length; i++) {
    if (_norm_(vals[i])) end = i + 1;
  }
  return end;
}

/** Totals row for a class sheet.
 * SAFETY: detect the ACTUAL TOTALS label in the live sheet; never trust a stale fixed row.
 */
function _totalsRowForClass_(sh) {
  // Prefer the shared helper after its safety fix.
  try {
    if (typeof findTotalsRowWide_ === "function") {
      const found = findTotalsRowWide_(sh, 12);
      if (found) return found;
    }
  } catch (_) {}

  const last = Math.min(sh.getLastRow(), 4000);
  if (last < 2) return null;
  const width = Math.min(12, Math.max(1, sh.getLastColumn()));
  const vals = sh.getRange(1, 1, last, width).getValues();

  // Bottom-up so the real final TOTALS row wins if old labels exist higher up.
  for (let r = vals.length - 1; r >= 0; r--) {
    for (let c = 0; c < width; c++) {
      const flat = _flatUpper_(vals[r][c]);
      if (flat === "TOTALS" || flat.startsWith("TOTALS")) return r + 1;
    }
  }
  return null;
}

function _getStudentRegion_(sh) {
  const t = _totalsRowForClass_(sh);
  const endRow = t ? (t - 1) : Math.max(1, sh.getLastRow());
  return { headerRow: 1, startRow: 2, totalsRow: t || null, endRow };
}

function _nextSerialNo_(sh) {
  // If you later make NO. formula-based, you can ignore this.
  try { if (typeof nextSerialNo_ === "function") return nextSerialNo_(sh); } catch (_) {}
  const map = _getHeaderIndexMap_(sh);
  const cNO = map[CL.H.NO];
  if (!cNO) return "";
  const region = _getStudentRegion_(sh);
  const endRow = region.endRow;
  if (endRow < 2) return 1;

  const vals = sh.getRange(2, cNO, endRow - 1, 1).getValues();
  let maxN = 0;
  for (const [v] of vals) {
    const n = Number(v);
    if (!isNaN(n) && n > maxN) maxN = n;
  }
  return maxN + 1;
}

/** Ensure sheet has enough rows WITHOUT shifting totals row (insert only AFTER maxRows) */
function _ensureSheetHasAtLeastRows_(sh, neededRows) {
  const need = Number(neededRows) || 0;
  if (need <= 0) return;
  const max = sh.getMaxRows();
  const gap = need - max;
  if (gap > 0) sh.insertRowsAfter(max, gap);
}

function _ensureStructureOneSafe_(sh) {
  try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}
}
function _rebuildOneDefaultersSafe_(sheetName) {
  try { if (typeof rebuildOneDefaulters === "function") rebuildOneDefaulters(sheetName); } catch (_) {}
}

/** ==================== HEADERS ==================== */
const LEAVERS_CLASS_COLS = [
  CL.H.NO, CL.H.DATE, CL.H.NAME, CL.H.SEX, CL.H.NAT, CL.H.DB,
  CL.H.LTB, CL.H.FEE, CL.H.PAID, CL.H.BAL, CL.H.CAUT, CL.H.TEL,
];

const LEAVERS_SYS_COLS = [
  "Former Class", "Date Left", "Exit Reason", "Status",
  "Reinstated To", "Reinstated At",
  "Source Sheet", "Source Row", "Created At", "Updated At", "Student ID",
];

function leaversDesiredHeaders_() {
  return [...LEAVERS_CLASS_COLS, ...LEAVERS_SYS_COLS];
}

function _headerCtx_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_norm_);
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const key = _flatUpper_(headers[i]);
    if (!key) continue;
    if (map[key] == null) map[key] = i + 1;
  }
  return { headers, map, lastCol };
}

// rename legacy headers -> desired ones (case/spacing tolerant)
function _leaversHeaderAliases_() {
  const A = new Map();
  A.set("BALANCE(OUTSTANDING)", CL.H.BAL);
  A.set("BALANCEOUTSTANDING", CL.H.BAL);
  A.set("PPAID", CL.H.PAID);
  A.set("PAID", CL.H.PAID);
  A.set("LASTTERMBALANCE", CL.H.LTB);
  A.set("FULLFEE", CL.H.FEE);
  A.set("DAYBOARDING", CL.H.DB);
  A.set("NO", CL.H.NO);

  A.set("FORMERCLASS", "Former Class");
  A.set("DATELEFT", "Date Left");
  A.set("EXITREASON", "Exit Reason");
  A.set("REINSTATEDTO", "Reinstated To");
  A.set("REINSTATEDAT", "Reinstated At");
  A.set("SOURCESHEET", "Source Sheet");
  A.set("SOURCEROW", "Source Row");
  A.set("CREATEDAT", "Created At");
  A.set("UPDATEDAT", "Updated At");
  A.set("STUDENTID", "Student ID");
  return A;
}

/** Ensure Leavers sheet exists (minimal, never clears data) */
function ensureLeaversSheet_() {
  const ss = _ss_();
  const name = _resolveLeaversSheetName_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(String(CL?.LEAVERS?.SHEET || name));

  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    const head = leaversDesiredHeaders_();
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
    return sh;
  }

  const ctx = _headerCtx_(sh);
  if (!ctx.headers.some(h => _norm_(h) !== "")) {
    const head = leaversDesiredHeaders_();
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _standardizeLeaversHeaders_(sh) {
  const aliases = _leaversHeaderAliases_();
  const { headers } = _headerCtx_(sh);
  for (let c = 0; c < headers.length; c++) {
    const key = _flatUpper_(headers[c]);
    const desired = aliases.get(key);
    if (desired && _norm_(desired) !== _norm_(headers[c])) sh.getRange(1, c + 1).setValue(desired);
  }
}

function _ensureLeaversHeadersExist_(sh) {
  const desired = leaversDesiredHeaders_();
  const { map } = _headerCtx_(sh);
  const missing = [];
  for (const h of desired) if (!map[_flatUpper_(h)]) missing.push(h);
  if (missing.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
}

/** Move columns (preserves data) so headers match desired order A.. */
function _reorderLeaversColumnsToDesired_(sh) {
  _standardizeLeaversHeaders_(sh);
  _ensureLeaversHeadersExist_(sh);

  const desired = leaversDesiredHeaders_();
  for (let i = 0; i < desired.length; i++) {
    const { map } = _headerCtx_(sh);
    const currentCol = map[_flatUpper_(desired[i])];
    const targetCol = i + 1;
    if (!currentCol || currentCol === targetCol) continue;
    sh.moveColumns(sh.getRange(1, currentCol, sh.getMaxRows(), 1), targetCol);
  }
}

/** Apply the requested layout: grey header, bold, centered, borders, formats, widths */
function _formatLeaversSheet_(sh) {
  const desired = leaversDesiredHeaders_();
  const totalCols = Math.max(sh.getLastColumn(), desired.length);
  const lastRow = Math.max(1, sh.getLastRow());

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, desired.length).setValues([desired]);

  sh.getRange(1, 1, 1, totalCols)
    .setFontWeight("bold")
    .setBackground("#D9D9D9")
    .setFontColor("#000000")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  try { sh.setRowHeight(1, 28); } catch (_) {}

  sh.getRange(1, 1, lastRow, totalCols)
    .setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);

  const { map } = _headerCtx_(sh);

  [CL.H.LTB, CL.H.FEE, CL.H.PAID, CL.H.BAL].forEach(h => {
    const c = map[_flatUpper_(h)];
    if (c && lastRow > 1) sh.getRange(2, c, lastRow - 1, 1).setNumberFormat("#,##0").setHorizontalAlignment("right");
  });

  [CL.H.DATE, "Date Left"].forEach(h => {
    const c = map[_flatUpper_(h)];
    if (c && lastRow > 1) sh.getRange(2, c, lastRow - 1, 1).setNumberFormat("dd/MM/yyyy");
  });

  const W = new Map([
    [CL.H.NO, 60], [CL.H.DATE, 130], [CL.H.NAME, 240], [CL.H.SEX, 60],
    [CL.H.NAT, 130], [CL.H.DB, 120], [CL.H.LTB, 140], [CL.H.FEE, 120],
    [CL.H.PAID, 120], [CL.H.BAL, 120], [CL.H.CAUT, 110], [CL.H.TEL, 120],
    ["Former Class", 140], ["Date Left", 110], ["Exit Reason", 180], ["Status", 110],
    ["Reinstated To", 140], ["Reinstated At", 140], ["Source Sheet", 140], ["Source Row", 100],
    ["Created At", 160], ["Updated At", 160], ["Student ID", 220],
  ]);
  desired.forEach((h, i) => {
    const w = W.get(h);
    if (w) { try { sh.setColumnWidth(i + 1, w); } catch (_) {} }
  });

  try { if (!sh.getFilter()) sh.getRange(1, 1, 1, totalCols).createFilter(); } catch (_) {}
}

/** ==================== RUNNABLE BUILDERS (EDITOR) ==================== */
function prepareLeaversDebtorsSheetSafe() {
  const sh = ensureLeaversSheet_();
  _reorderLeaversColumnsToDesired_(sh);
  _formatLeaversSheet_(sh);
  _toast_(`SAFE layout applied (data kept): "${sh.getName()}"`, "Leavers – Debtors", 6);
}

function rebuildLeaversDebtorsSheetClearOnly() {
  const ss = _ss_();
  const name = _resolveLeaversSheetName_();

  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "CLEAR & REBUILD (DANGER)",
    `This will CLEAR ALL DATA in "${name}" and rebuild headers/format.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const sh = ss.getSheetByName(name) || ss.insertSheet(name);

  try { if (sh.getFilter()) sh.getFilter().remove(); } catch (_) {}
  sh.clear();

  const head = leaversDesiredHeaders_();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  _formatLeaversSheet_(sh);

  _toast_(`CLEAR rebuild complete (all data wiped): "${name}"`, "Leavers – Debtors", 6);
}

function rebuildLeaversDebtorsSheetFresh() {
  const ss = _ss_();
  const name = _resolveLeaversSheetName_();

  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "FULL REBUILD (DANGER)",
    `This will DELETE the entire sheet "${name}" and recreate it.\n\nALL DATA WILL BE LOST.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);

  const canonical = String(CL?.LEAVERS?.SHEET || name);
  const sh = ss.insertSheet(canonical);
  const head = leaversDesiredHeaders_();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  _formatLeaversSheet_(sh);

  _toast_(`FULL rebuild complete (sheet recreated): "${canonical}"`, "Leavers – Debtors", 6);
}

/** ==================== CONTEXT HELPERS ==================== */
function _getDebtorsCtx_() {
  const sh = ensureLeaversSheet_();
  const { headers, map } = _headerCtx_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const rows = (lastRow > 1) ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { sh, headers, map, rows, lastRow, lastCol };
}
function _getDebtor_(ctx, rowArr, header) {
  const col = ctx.map[_flatUpper_(header)];
  return col ? rowArr[col - 1] : "";
}

/** MOST COMMON Full Fee (MODE) in target class sheet (above TOTALS) */
function mostCommonFullFee_(sh) {
  const map = _getHeaderIndexMap_(sh);
  const cFEE = map[CL.H.FEE];
  if (!cFEE) return 0;

  const region = _getStudentRegion_(sh);
  const endRow = region.endRow;
  if (endRow < 2) return 0;

  const vals = sh.getRange(2, cFEE, endRow - 1, 1).getValues().flat();
  const freq = new Map();

  for (const v of vals) {
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0) continue;
    freq.set(n, (freq.get(n) || 0) + 1);
  }

  let best = 0, bestCount = -1;
  for (const [fee, count] of freq.entries()) {
    if (count > bestCount || (count === bestCount && fee > best)) {
      best = fee; bestCount = count;
    }
  }
  return best || 0;
}

/**
 * Find the first empty student row (by Student Name) above TOTALS.
 * This is how we “add/restore” without inserting rows.
 */
function _firstEmptyStudentRow_(sh) {
  const map = _getHeaderIndexMap_(sh);
  const cNAME = map[CL.H.NAME] || 3;

  const tRow = _totalsRowForClass_(sh);
  if (!tRow || tRow < 3) return null;

  // Ensure we physically have enough rows to reach the totals row without shifting it
  _ensureSheetHasAtLeastRows_(sh, tRow);

  const start = 2;
  const end = tRow - 1;
  const n = end - start + 1;
  if (n <= 0) return null;

  const vals = sh.getRange(start, cNAME, n, 1).getDisplayValues().flat();
  for (let i = 0; i < vals.length; i++) {
    if (!_norm_(vals[i])) return start + i;
  }
  return null; // no space
}

/**
 * Clear a student row WITHOUT deleting it (keeps totals row fixed).
 * We clear values/content across the main header area, but we keep formulas if present.
 */
function _clearStudentRowSafely_(sh, row) {
  const headerEnd = _lastHeaderCol_(sh);
  const rng = sh.getRange(row, 1, 1, headerEnd);
  const formulas = rng.getFormulas(); // 1 x N
  const values = rng.getValues();

  for (let c = 0; c < values[0].length; c++) {
    if (formulas[0][c]) continue; // keep formula cells
    values[0][c] = "";
  }
  rng.setValues(values);
}

/** ==================== ACTIONS: Record / Restore / Settle ==================== */
function recordLeaverQuick() {
  const ui = SpreadsheetApp.getUi();
  const ss = _ss_();

  try { prepareLeaversDebtorsSheetSafe(); } catch (_) {}

  const r1 = ui.prompt("Record Leaver", "Student name (exact as on class list):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const name = _norm_(r1.getResponseText());
  if (!name) return ui.alert("Name is required.");

  const r2 = ui.prompt("Record Leaver", 'Class sheet name (e.g., "Primary 3 list"):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const className = _norm_(r2.getResponseText());
  if (!className) return ui.alert("Class sheet name is required.");

  const r3 = ui.prompt("Record Leaver", "Date left (dd/mm/yyyy, blank = today):", ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  let dateLeft = _norm_(r3.getResponseText());
  if (!dateLeft) dateLeft = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const r4 = ui.prompt("Record Leaver", "Exit reason (optional):", ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK) return;
  const exitReason = _norm_(r4.getResponseText());

  const src = ss.getSheetByName(className);
  if (!src) return ui.alert(`No sheet named "${className}".`);

  _ensureStructureOneSafe_(src);

  const tRow = _totalsRowForClass_(src);
  if (!tRow) return ui.alert(`Totals row not found for "${className}". Add it to CLASS_TOTALS_ROWS or place a TOTALS label.`);

  const srcMap = _getHeaderIndexMap_(src);
  if (!srcMap[CL.H.NAME]) return ui.alert("Missing 'Student Name' header on class sheet.");

  const endRow = tRow - 1;
  if (endRow < 2) return ui.alert("No student rows found (above TOTALS).");

  const lastCol = src.getLastColumn();
  const rows = src.getRange(2, 1, endRow - 1, lastCol).getValues();

  const nameU = _flatUpper_(name);
  const idx = rows.findIndex(r => _flatUpper_(r[srcMap[CL.H.NAME] - 1]) === nameU);
  if (idx === -1) return ui.alert(`"${name}" not found in ${className} (above TOTALS).`);

  const absRow = 2 + idx;
  if (absRow >= tRow) return ui.alert("Safety stop: matched row appears at/under TOTALS. No removal performed.");

  // Prepare debtor record
  const ctx = _getDebtorsCtx_();
  const colName = ctx.map[_flatUpper_(CL.H.NAME)];
  const colFormer = ctx.map[_flatUpper_("Former Class")];
  const colStatus = ctx.map[_flatUpper_("Status")];

  if (colName && colFormer && colStatus) {
    const clash = ctx.rows.some(r =>
      _flatUpper_(r[colName - 1]) === nameU &&
      _flatUpper_(r[colFormer - 1]) === _flatUpper_(className) &&
      !["SETTLED", "REINSTATED"].includes(_flatUpper_(r[colStatus - 1]))
    );
    if (clash) return ui.alert("An unsettled debtor with the same name and class already exists.");
  }

  const out = new Array(ctx.lastCol).fill("");
  for (const h of LEAVERS_CLASS_COLS) {
    const destCol = ctx.map[_flatUpper_(h)];
    const srcCol = srcMap[h];
    if (destCol && srcCol) out[destCol - 1] = rows[idx][srcCol - 1];
  }

  const now = nowStamp_();
  const studentId = Utilities.getUuid();
  const set = (h, v) => { const c = ctx.map[_flatUpper_(h)]; if (c) out[c - 1] = v; };

  const actualClassName = src.getName();
  set("Former Class", actualClassName);
  set("Date Left", dateLeft);
  set("Exit Reason", exitReason || "");
  set("Status", "Unsettled");
  set("Reinstated To", "");
  set("Reinstated At", "");
  set("Source Sheet", actualClassName);
  set("Source Row", absRow);
  set("Created At", now);
  set("Updated At", now);
  set("Student ID", studentId);

  ctx.sh.appendRow(out);

  const go = ui.alert("Remove from class list?", `Clear "${name}" from "${className}" now?\n\n(ROW WILL NOT BE DELETED — totals row stays fixed)`, ui.ButtonSet.YES_NO);
  if (go === ui.Button.YES) {
    _clearStudentRowSafely_(src, absRow);
    // Always pass the canonical live sheet name; never a user-typed casing variant.
    _rebuildOneDefaultersSafe_(src.getName());
  }

  _toast_(`Leaver recorded: ${name}`, "Leavers – Debtors", 6);
}

function restoreLeaverToClass() {
  const ui = SpreadsheetApp.getUi();
  const ss = _ss_();

  const ctx = _getDebtorsCtx_();
  if (!ctx.rows.length) return ui.alert("No debtor records.");

  const r1 = ui.prompt("Restore Debtor", "Student name (as in Debtors):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nameInput = _norm_(r1.getResponseText());
  const nameU = _flatUpper_(nameInput);
  if (!nameU) return ui.alert("Name is required.");

  const matches = [];
  for (let i = 0; i < ctx.rows.length; i++) {
    const nm = _flatUpper_(_getDebtor_(ctx, ctx.rows[i], CL.H.NAME));
    const st = _flatUpper_(_getDebtor_(ctx, ctx.rows[i], "Status"));
    // Allow restoration whether the outstanding balance is positive OR zero.
    // Only records already restored to a class are excluded.
    if (nm === nameU && st !== "REINSTATED") matches.push(i);
  }
  if (!matches.length) return ui.alert("No restorable Leavers/Debtors record found for that name.");

  let dIdx = matches[0];
  if (matches.length > 1) {
    const lines = matches.slice(0, 15).map((ix, k) => {
      const r = ctx.rows[ix];
      const former = _norm_(_getDebtor_(ctx, r, "Former Class"));
      const left = _norm_(_getDebtor_(ctx, r, "Date Left"));
      const bal = Number(_getDebtor_(ctx, r, CL.H.BAL)) || 0;
      return `${k + 1}) ${former || "-"} | Left: ${left || "-"} | Bal: ${bal}`;
    });
    const res = ui.prompt("Restore Debtor", "Multiple matches:\n\n" + lines.join("\n") + "\n\nEnter number:", ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const n = Number(_norm_(res.getResponseText()));
    if (!Number.isFinite(n) || n < 1 || n > Math.min(15, matches.length)) return ui.alert("Invalid selection.");
    dIdx = matches[n - 1];
  }

  const d = ctx.rows[dIdx];

  const formerClass = _norm_(_getDebtor_(ctx, d, "Former Class") || _getDebtor_(ctx, d, "Source Sheet"));
  const debtorBal = Number(_getDebtor_(ctx, d, CL.H.BAL)) || 0;
  const prevFee = Number(_getDebtor_(ctx, d, CL.H.FEE)) || 0;

  const dateAdm = _getDebtor_(ctx, d, CL.H.DATE) || "";
  const sex = _getDebtor_(ctx, d, CL.H.SEX) || "";
  const nat = _getDebtor_(ctx, d, CL.H.NAT) || "";
  const db  = _getDebtor_(ctx, d, CL.H.DB) || "";
  const caut = _getDebtor_(ctx, d, CL.H.CAUT) || "";
  const tel = _getDebtor_(ctx, d, CL.H.TEL) || "";

  const r2 = ui.prompt("Restore Debtor", `Target class (default: ${formerClass || "?"})`, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const targetClass = _norm_(r2.getResponseText()) || formerClass;
  if (!targetClass) return ui.alert("Target class is required.");

  const to = ss.getSheetByName(targetClass);
  if (!to) return ui.alert(`No sheet named "${targetClass}".`);

  _ensureStructureOneSafe_(to);

  const tRow = _totalsRowForClass_(to);
  if (!tRow) return ui.alert(`Totals row not found for "${targetClass}". Add it to CLASS_TOTALS_ROWS or place a TOTALS label.`);

  // Find first available empty row above totals (NO inserts)
  const writeRow = _firstEmptyStudentRow_(to);
  if (!writeRow) {
    return ui.alert(
      `No empty student row available above TOTALS in "${targetClass}".\n\n` +
      `Because totals rows are fixed, we cannot insert rows above TOTALS.\n` +
      `Clear an empty line (remove a student) or increase capacity by moving TOTALS lower and updating CLASS_TOTALS_ROWS.`
    );
  }

  const modeFee = mostCommonFullFee_(to);
  const feeToUse = modeFee || prevFee || 0;

  // Copy formatting from previous row (visual only)
  const lastCol = to.getLastColumn();
  const writeRange = to.getRange(writeRow, 1, 1, lastCol);
  if (writeRow > 2) to.getRange(writeRow - 1, 1, 1, lastCol).copyTo(writeRange, { formatOnly: true });
  else writeRange.clearFormat();

  // Write row values
  const map = _getHeaderIndexMap_(to);
  const row = new Array(lastCol).fill("");
  const idxOf = (h) => (map[h] ? map[h] - 1 : null);

  const iNO   = idxOf(CL.H.NO);
  const iDATE = idxOf(CL.H.DATE);
  const iNAME = idxOf(CL.H.NAME);
  const iSEX  = idxOf(CL.H.SEX);
  const iNAT  = idxOf(CL.H.NAT);
  const iDB   = idxOf(CL.H.DB);
  const iLTB  = idxOf(CL.H.LTB);
  const iFEE  = idxOf(CL.H.FEE);
  const iPAID = idxOf(CL.H.PAID);
  const iCAUT = idxOf(CL.H.CAUT);
  const iTEL  = idxOf(CL.H.TEL);

  if (iNO   !== null) row[iNO] = _nextSerialNo_(to);
  if (iDATE !== null) row[iDATE] = dateAdm;
  if (iNAME !== null) row[iNAME] = nameInput;
  if (iSEX  !== null) row[iSEX] = sex;
  if (iNAT  !== null) row[iNAT] = nat;
  if (iDB   !== null) row[iDB] = db;
  if (iCAUT !== null) row[iCAUT] = caut;
  if (iTEL  !== null) row[iTEL] = tel;

  // debtor BALANCE -> class LTB
  if (iLTB !== null) row[iLTB] = debtorBal;

  // class fee = MODE (fallback prevFee)
  if (iFEE !== null) row[iFEE] = feeToUse;

  // paid resets
  if (iPAID !== null) row[iPAID] = 0;

  writeRange.setValues([row]);
SpreadsheetApp.flush();

// Verify the student was actually written into the target class before
// touching the debtor record.
const writtenName = iNAME !== null
  ? _flatUpper_(to.getRange(writeRow, iNAME + 1).getDisplayValue())
  : "";

if (writtenName !== nameU) {
  _clearStudentRowSafely_(to, writeRow);
  SpreadsheetApp.flush();
  return ui.alert(
    `Restore stopped: ${nameInput} could not be verified in ${to.getName()}.\n\n` +
    `The Leavers – Debtors record was NOT removed.`
  );
}

// Rebuild Defaulters before completing the move. If this fails, roll back
// the class write so the student remains only in Leavers – Debtors.
try {
  if (typeof rebuildOneDefaulters === "function") {
    rebuildOneDefaulters(to.getName());
  }
} catch (err) {
  _clearStudentRowSafely_(to, writeRow);
  SpreadsheetApp.flush();
  return ui.alert(
    `Restore stopped because Defaulters could not be rebuilt.\n\n` +
    `${err && err.message ? err.message : err}\n\n` +
    `The class write was rolled back and the debtor record was kept.`
  );
}

// Re-check the exact debtor record immediately before deletion. Prefer the
// immutable Student ID; fall back to Student Name for older imported rows.
const absDebtRow = 2 + dIdx;
const debtNameCol = ctx.map[_flatUpper_(CL.H.NAME)];
const debtIdCol = ctx.map[_flatUpper_("Student ID")];
const expectedStudentId = _norm_(_getDebtor_(ctx, d, "Student ID"));
const liveDebtName = debtNameCol
  ? _flatUpper_(ctx.sh.getRange(absDebtRow, debtNameCol).getDisplayValue())
  : "";
const liveStudentId = debtIdCol
  ? _norm_(ctx.sh.getRange(absDebtRow, debtIdCol).getDisplayValue())
  : "";

const sameDebtorRecord = expectedStudentId
  ? liveStudentId === expectedStudentId
  : liveDebtName === nameU;

if (!sameDebtorRecord) {
  _clearStudentRowSafely_(to, writeRow);
  try {
    if (typeof rebuildOneDefaulters === "function") rebuildOneDefaulters(to.getName());
  } catch (_) {}
  SpreadsheetApp.flush();
  return ui.alert(
    `Restore stopped: the Leavers – Debtors row changed before deletion.\n\n` +
    `The class write was rolled back so the student cannot exist in both places.`
  );
}

// COMPLETE THE MOVE: remove the source debtor row. If deletion fails for any
// reason, roll back the class write to preserve the one-place-only rule.
try {
  ctx.sh.deleteRow(absDebtRow);
  SpreadsheetApp.flush();
} catch (err) {
  _clearStudentRowSafely_(to, writeRow);
  try {
    if (typeof rebuildOneDefaulters === "function") rebuildOneDefaulters(to.getName());
  } catch (_) {}
  SpreadsheetApp.flush();
  return ui.alert(
    `Restore stopped because the debtor row could not be removed.\n\n` +
    `${err && err.message ? err.message : err}\n\n` +
    `The class write was rolled back, so there is no duplicate.`
  );
}

// Keep the class-by-class debtor report current after the source row is gone.
try {
  if (typeof rebuildLeaversByClassReport === "function") {
    rebuildLeaversByClassReport();
  }
} catch (_) {}

_toast_(
  `Restored ${nameInput} to ${to.getName()} and removed the Debtors record.`,
  "Leavers – Debtors",
  6
);
}

function markLeaverSettled() {
  const ui = SpreadsheetApp.getUi();
  const ctx = _getDebtorsCtx_();
  if (!ctx.rows.length) return ui.alert("No debtor records.");

  const r1 = ui.prompt("Settle Debtor", "Student name (as in Debtors):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nameU = _flatUpper_(_norm_(r1.getResponseText()));
  if (!nameU) return ui.alert("Name is required.");

  const matches = [];
  for (let i = 0; i < ctx.rows.length; i++) {
    const nm = _flatUpper_(_getDebtor_(ctx, ctx.rows[i], CL.H.NAME));
    const st = _flatUpper_(_getDebtor_(ctx, ctx.rows[i], "Status"));
    if (nm === nameU && st !== "SETTLED") matches.push(i);
  }
  if (!matches.length) return ui.alert("No matching unsettled record found.");

  const idx = matches[0];
  const row = 2 + idx;
  const now = nowStamp_();

  const colBal = ctx.map[_flatUpper_(CL.H.BAL)];
  const colStatus = ctx.map[_flatUpper_("Status")];
  const colUpdated = ctx.map[_flatUpper_("Updated At")];

  if (colBal) ctx.sh.getRange(row, colBal).setValue(0);
  if (colStatus) ctx.sh.getRange(row, colStatus).setValue("Settled");
  if (colUpdated) ctx.sh.getRange(row, colUpdated).setValue(now);

  _toast_("Marked as settled.", "Leavers – Debtors", 4);
}

/**
 * ONE-TIME IMPORT TOOL (unchanged behavior; safe append only)
 */
function importSheet39ToLeaversDebtors_Once() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const SRC_NAME = "Sheet39";
  const LEAVERS_NAME = _resolveLeaversSheetName_();

  const src = ss.getSheetByName(SRC_NAME);
  if (!src) {
    ui.alert(`Missing sheet: "${SRC_NAME}"`);
    return;
  }

  let leavers = ss.getSheetByName(LEAVERS_NAME);
  if (typeof ensureLeaversSheet_ === "function") {
    leavers = ensureLeaversSheet_();
  } else {
    if (!leavers) leavers = ss.insertSheet(LEAVERS_NAME);
    if (leavers.getLastRow() === 0) {
      leavers.getRange(1, 1, 1, 23).setValues([[...leaversDesiredHeaders_()]]);
      leavers.setFrozenRows(1);
    }
  }

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  const flatUpper = (v) => norm(v).replace(/\s+/g, "").toUpperCase();

  const headerMap = (sh) => {
    const lastCol = Math.max(1, sh.getLastColumn());
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(norm);
    const map = {};
    headers.forEach((h, i) => {
      const k = flatUpper(h);
      if (k && map[k] == null) map[k] = i + 1;
    });
    return { headers, map, lastCol };
  };

  const L = headerMap(leavers);
  const col = (h) => L.map[flatUpper(h)] || null;

  const cNO = col("NO.") || col("NO");
  const cDATE = col("Date of Admission");
  const cNAME = col("Student Name");
  const cSEX = col("SEX");
  const cNAT = col("NATIONALITY");
  const cDB = col("DAY/BOARDING");
  const cLTB = col("LAST TERM BALANCE");
  const cFEE = col("FULL FEE");
  const cPAID = col("P (Paid)");
  const cBAL = col("BALANCE");

  const cFormer = col("Former Class");
  const cDateLeft = col("Date Left");
  const cExit = col("Exit Reason");
  const cStatus = col("Status");
  const cSourceSheet = col("Source Sheet");
  const cSourceRow = col("Source Row");
  const cCreated = col("Created At");
  const cUpdated = col("Updated At");
  const cStudentId = col("Student ID");

  if (!cNAME || !cFormer || !cStatus) {
    ui.alert(`Leavers sheet is missing required headers (Student Name / Former Class / Status). Run prepareLeaversDebtorsSheetSafe() first.`);
    return;
  }

  const existingKeys = new Set();
  const leaversLastRow = leavers.getLastRow();
  if (leaversLastRow > 1) {
    const values = leavers.getRange(2, 1, leaversLastRow - 1, L.lastCol).getValues();
    values.forEach(r => {
      const nm = flatUpper(r[cNAME - 1]);
      const fm = flatUpper(r[cFormer - 1]);
      const st = flatUpper(r[cStatus - 1]);
      if (!nm) return;
      if (st === "SETTLED" || st === "REINSTATED") return;
      existingKeys.add(`${nm}|${fm}`);
    });
  }

  let nextNo = 1;
  if (cNO && leaversLastRow > 1) {
    const noVals = leavers.getRange(2, cNO, leaversLastRow - 1, 1).getValues().flat();
    let maxNo = 0;
    noVals.forEach(v => {
      const n = Number(v);
      if (Number.isFinite(n) && n > maxNo) maxNo = n;
    });
    nextNo = maxNo + 1;
  }

  const srcLastRow = src.getLastRow();
  const srcLastCol = Math.max(12, src.getLastColumn());
  if (srcLastRow < 1) {
    ui.alert(`"${SRC_NAME}" is empty.`);
    return;
  }

  const srcVals = src.getRange(1, 1, srcLastRow, srcLastCol).getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const nowStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  const outRows = [];
  let skippedBlank = 0;
  let skippedDup = 0;

  for (let r = 0; r < srcVals.length; r++) {
    const row = srcVals[r];

    const formerClass = norm(row[0]);     // A
    const dateAdm = row[1];               // B
    const studentName = norm(row[2]);     // C

    if (!studentName) { skippedBlank++; continue; }

    const key = `${flatUpper(studentName)}|${flatUpper(formerClass)}`;
    if (existingKeys.has(key)) { skippedDup++; continue; }

    const sex = norm(row[4]);             // E
    const nat = norm(row[5]);             // F
    const db = norm(row[6]);              // G
    const ltb = row[7];                   // H
    const fee = row[8];                   // I
    const paid = row[10];                 // K
    const bal = row[11];                  // L

    const out = new Array(L.lastCol).fill("");

    if (cNO) out[cNO - 1] = nextNo++;
    if (cDATE) out[cDATE - 1] = dateAdm;
    if (cNAME) out[cNAME - 1] = studentName;
    if (cSEX) out[cSEX - 1] = sex;
    if (cNAT) out[cNAT - 1] = nat;
    if (cDB) out[cDB - 1] = db;
    if (cLTB) out[cLTB - 1] = ltb;
    if (cFEE) out[cFEE - 1] = fee;
    if (cPAID) out[cPAID - 1] = paid;
    if (cBAL) out[cBAL - 1] = bal;

    out[cFormer - 1] = formerClass || "";
    if (cDateLeft) out[cDateLeft - 1] = today;
    if (cExit) out[cExit - 1] = "Imported from Sheet39";
    out[cStatus - 1] = "Unsettled";
    if (cSourceSheet) out[cSourceSheet - 1] = SRC_NAME;
    if (cSourceRow) out[cSourceRow - 1] = r + 1;
    if (cCreated) out[cCreated - 1] = nowStamp;
    if (cUpdated) out[cUpdated - 1] = nowStamp;
    if (cStudentId) out[cStudentId - 1] = Utilities.getUuid();

    outRows.push(out);
    existingKeys.add(key);
  }

  if (!outRows.length) {
    ui.alert(`Nothing imported.\n\nSkipped blanks: ${skippedBlank}\nSkipped duplicates: ${skippedDup}`);
    return;
  }

  const startRow = leavers.getLastRow() + 1;
  const needRows = startRow + outRows.length - 1;
  if (leavers.getMaxRows() < needRows) {
    leavers.insertRowsAfter(leavers.getMaxRows(), needRows - leavers.getMaxRows());
  }

  leavers.getRange(startRow, 1, outRows.length, L.lastCol).setValues(outRows);

  ss.toast(
    `Imported ${outRows.length} row(s) from Sheet39.\nSkipped blanks: ${skippedBlank}, duplicates: ${skippedDup}`,
    "Leavers – Debtors",
    8
  );
}

/************************************************************
 * LEAVERS / DEBTORS — FULL CLASS-BY-CLASS REPORT
 *
 * SOURCE SHEET:
 *   Leavers – Debtors
 *
 * REPORT SHEET:
 *   Leavers by Class
 *
 * CLASS ORDER:
 *   Baby class list
 *   Middle class list
 *   Top Class list
 *   Primary 1 list
 *   Primary 2 list
 *   Primary 3 list
 *   Primary 4 list
 *   Primary 5 list
 *   Primary 6 list
 *   Primary 7 list
 *   Senior 1 list
 *   Senior 2 list
 *   Senior 3 list
 *   Senior 4 list
 *   Senior 5 list
 *   Senior 6 list
 *   WAITING TO BE ORGANIZED
 *
 * INCLUDED:
 *   - BALANCE greater than zero
 *   - Status not Settled
 *   - Status not Reinstated
 *
 * RUN MANUALLY:
 *   rebuildLeaversByClassReport()
 *
 * DAILY TRIGGER:
 *   installLeaversByClassTrigger()
 *
 * HOURLY TRIGGER:
 *   installHourlyLeaversByClassTrigger()
 ************************************************************/

const LEAVERS_CLASS_REPORT_SHEET = "Leavers by Class";

const LEAVERS_ALL_CLASS_ORDER = [
  "Baby class list",
  "Middle class list",
  "Top Class list",

  "Primary 1 list",
  "Primary 2 list",
  "Primary 3 list",
  "Primary 4 list",
  "Primary 5 list",
  "Primary 6 list",
  "Primary 7 list",

  "Senior 1 list",
  "Senior 2 list",
  "Senior 3 list",
  "Senior 4 list",
  "Senior 5 list",
  "Senior 6 list",
];

/**
 * Main report builder.
 *
 * Reads Leavers – Debtors and rebuilds the complete
 * class-by-class report.
 */
function rebuildLeaversByClassReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceName =
    typeof _resolveLeaversSheetName_ === "function"
      ? _resolveLeaversSheetName_()
      : "Leavers – Debtors";

  const source = ss.getSheetByName(sourceName);

  if (!source) {
    throw new Error(
      `Source sheet "${sourceName}" was not found.`
    );
  }

  const lastRow = source.getLastRow();
  const lastCol = source.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    throw new Error(
      `Source sheet "${sourceName}" is empty.`
    );
  }

  const allValues = source
    .getRange(1, 1, lastRow, lastCol)
    .getValues();

  const headers = allValues[0].map(
    normalizeLeaversClassText_
  );

  const sourceRows = allValues.slice(1);

  /*
   * Build header map.
   */
  const headerMap = {};

  headers.forEach((header, index) => {
    const key =
      normalizeLeaversClassKey_(header);

    if (key && headerMap[key] == null) {
      headerMap[key] = index;
    }
  });

  /*
   * Required headers.
   */
  const requiredHeaders = [
    "Student Name",
    "Former Class",
    "Status",
    "BALANCE",
  ];

  requiredHeaders.forEach(header => {
    const key =
      normalizeLeaversClassKey_(header);

    if (headerMap[key] == null) {
      throw new Error(
        `Required header "${header}" is missing from "${sourceName}".`
      );
    }
  });

  const nameIndex =
    headerMap[
      normalizeLeaversClassKey_("Student Name")
    ];

  const classIndex =
    headerMap[
      normalizeLeaversClassKey_("Former Class")
    ];

  const statusIndex =
    headerMap[
      normalizeLeaversClassKey_("Status")
    ];

  const balanceIndex =
    headerMap[
      normalizeLeaversClassKey_("BALANCE")
    ];

  const dateLeftIndex =
    headerMap[
      normalizeLeaversClassKey_("Date Left")
    ];

  /*
   * Prepare every class group.
   */
  const grouped = new Map();

  LEAVERS_ALL_CLASS_ORDER.forEach(className => {
    grouped.set(className, []);
  });

  /*
   * Records with an unclear class go here.
   */
  const waitingRows = [];

  let organizedCount = 0;
  let waitingCount = 0;
  let zeroBalanceExcluded = 0;
  let completedExcluded = 0;
  let blankRowsExcluded = 0;

  /*
   * Process source records.
   */
  sourceRows.forEach(row => {
    const studentName =
      normalizeLeaversClassText_(
        row[nameIndex]
      );

    if (!studentName) {
      blankRowsExcluded++;
      return;
    }

    /*
     * Ignore Settled and Reinstated records.
     */
    const status =
      normalizeLeaversClassKey_(
        row[statusIndex]
      );

    if (
      status === "SETTLED" ||
      status === "REINSTATED"
    ) {
      completedExcluded++;
      return;
    }

    /*
     * Only include balances greater than zero.
     */
    const balance =
      parseLeaversClassAmount_(
        row[balanceIndex]
      );

    if (balance <= 0) {
      zeroBalanceExcluded++;
      return;
    }

    const originalClass =
      normalizeLeaversClassText_(
        row[classIndex]
      );

    /*
     * Standardize different class-name formats.
     */
    const canonicalClass =
      canonicalLeaversClass_(
        originalClass
      );

    if (canonicalClass) {
      const reportRow = row.slice();

      /*
       * Show the standardized class name in the report.
       * The master record is not changed.
       */
      reportRow[classIndex] =
        canonicalClass;

      grouped
        .get(canonicalClass)
        .push(reportRow);

      organizedCount++;
      return;
    }

    /*
     * Unrecognized class record.
     */
    waitingRows.push(row.slice());
    waitingCount++;
  });

  /*
   * Sort each class:
   *   1. Latest Date Left first
   *   2. Student Name alphabetically
   */
  grouped.forEach(classRows => {
    sortLeaversClassRows_(
      classRows,
      nameIndex,
      dateLeftIndex
    );
  });

  sortLeaversClassRows_(
    waitingRows,
    nameIndex,
    dateLeftIndex
  );

  /*
   * Create or retrieve the report sheet.
   */
  let report =
    ss.getSheetByName(
      LEAVERS_CLASS_REPORT_SHEET
    );

  if (!report) {
    report =
      ss.insertSheet(
        LEAVERS_CLASS_REPORT_SHEET
      );
  }

  /*
   * Remove existing filter.
   */
  try {
    if (report.getFilter()) {
      report.getFilter().remove();
    }
  } catch (_) {}

  /*
   * Break old merged cells before clearing.
   */
  try {
    report
      .getRange(
        1,
        1,
        report.getMaxRows(),
        report.getMaxColumns()
      )
      .breakApart();
  } catch (_) {}

  /*
   * Clear previous report.
   */
  report.clear();

  ensureLeaversClassReportColumns_(
    report,
    lastCol
  );

  let outputRow = 1;

  /*
   * Write all class sections in fixed order.
   */
  LEAVERS_ALL_CLASS_ORDER.forEach(
    className => {
      const classRows =
        grouped.get(className) || [];

      outputRow =
        writeLeaversClassSection_({
          sheet: report,
          outputRow: outputRow,
          title: className,
          rows: classRows,
          headers: headers,
          lastCol: lastCol,
          balanceIndex: balanceIndex,
          dateLeftIndex: dateLeftIndex,
          titleColour:
            getLeaversClassSectionColour_(
              className
            ),
          emptyMessage:
            "No debtors with a balance greater than zero.",
        });
    }
  );

  /*
   * Waiting section at the bottom.
   */
  outputRow =
    writeLeaversClassSection_({
      sheet: report,
      outputRow: outputRow,
      title: "WAITING TO BE ORGANIZED",
      rows: waitingRows,
      headers: headers,
      lastCol: lastCol,
      balanceIndex: balanceIndex,
      dateLeftIndex: dateLeftIndex,
      titleColour: "#cc0000",
      emptyMessage:
        "No records are waiting to be organized.",
    });

  /*
   * Format columns.
   */
  formatLeaversClassReportColumns_(
    report,
    headers
  );

  const reportLastRow =
    report.getLastRow();

  if (reportLastRow > 0) {
    report
      .getRange(
        1,
        1,
        reportLastRow,
        lastCol
      )
      .setVerticalAlignment("middle");
  }

  report.setFrozenRows(0);

  /*
   * Rebuild note.
   */
  report
    .getRange("A1")
    .setNote(
      `Automatically rebuilt from "${sourceName}" at ` +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      )
    );

  SpreadsheetApp.flush();

  ss.toast(
    `${organizedCount} debtor(s) organized. ` +
    `${waitingCount} waiting. ` +
    `${zeroBalanceExcluded} zero-balance record(s) excluded. ` +
    `${completedExcluded} settled/reinstated record(s) excluded.`,
    "Leavers by Class",
    10
  );
}

/**
 * Writes one class section.
 *
 * Returns the next empty row number.
 */
function writeLeaversClassSection_(options) {
  const sheet =
    options.sheet;

  let rowNumber =
    options.outputRow;

  const title =
    options.title;

  const rows =
    options.rows || [];

  const headers =
    options.headers;

  const lastCol =
    options.lastCol;

  const balanceIndex =
    options.balanceIndex;

  const dateLeftIndex =
    options.dateLeftIndex;

  const titleColour =
    options.titleColour || "#666666";

  const emptyMessage =
    options.emptyMessage ||
    "No matching records.";

  const countText =
    rows.length === 1
      ? "1 debtor"
      : `${rows.length} debtors`;

  /*
   * Section title.
   */
  sheet
    .getRange(
      rowNumber,
      1,
      1,
      lastCol
    )
    .merge()
    .setValue(
      `${title} — ${countText}`
    )
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontColor("#ffffff")
    .setBackground(titleColour)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(
    rowNumber,
    30
  );

  rowNumber++;

  /*
   * Header row.
   */
  sheet
    .getRange(
      rowNumber,
      1,
      1,
      lastCol
    )
    .setValues([headers])
    .setFontWeight("bold")
    .setFontColor("#000000")
    .setBackground("#d9e2f3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(
    rowNumber,
    28
  );

  rowNumber++;

  if (rows.length > 0) {
    const dataStartRow =
      rowNumber;

    /*
     * Write students.
     */
    sheet
      .getRange(
        dataStartRow,
        1,
        rows.length,
        lastCol
      )
      .setValues(rows)
      .setVerticalAlignment("middle");

    /*
     * Format balance.
     */
    if (balanceIndex != null) {
      sheet
        .getRange(
          dataStartRow,
          balanceIndex + 1,
          rows.length,
          1
        )
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");
    }

    /*
     * Format Date Left.
     */
    if (dateLeftIndex != null) {
      sheet
        .getRange(
          dataStartRow,
          dateLeftIndex + 1,
          rows.length,
          1
        )
        .setNumberFormat("dd/MM/yyyy");
    }

    /*
     * Add borders around headers and data.
     */
    sheet
      .getRange(
        dataStartRow - 1,
        1,
        rows.length + 1,
        lastCol
      )
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#b7b7b7",
        SpreadsheetApp.BorderStyle.SOLID
      );

    rowNumber += rows.length;

    /*
     * Section total.
     */
    if (balanceIndex != null) {
      const balanceColumn =
        columnLetterLeaversClass_(
          balanceIndex + 1
        );

      if (balanceIndex > 0) {
        sheet
          .getRange(
            rowNumber,
            1,
            1,
            balanceIndex
          )
          .merge()
          .setValue(
            `${title} Total Balance`
          )
          .setFontWeight("bold")
          .setHorizontalAlignment("right")
          .setBackground("#eeeeee");
      }

      sheet
        .getRange(
          rowNumber,
          balanceIndex + 1
        )
        .setFormula(
          `=SUM(${balanceColumn}${dataStartRow}:` +
          `${balanceColumn}${rowNumber - 1})`
        )
        .setNumberFormat("#,##0")
        .setFontWeight("bold")
        .setBackground("#eeeeee");

      sheet
        .getRange(
          rowNumber,
          1,
          1,
          lastCol
        )
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          "#b7b7b7",
          SpreadsheetApp.BorderStyle.SOLID
        );
    }

    rowNumber++;
  } else {
    /*
     * Empty class message.
     */
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastCol
      )
      .merge()
      .setValue(emptyMessage)
      .setFontStyle("italic")
      .setFontColor("#666666")
      .setBackground("#f8f8f8")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    sheet.setRowHeight(
      rowNumber,
      24
    );

    rowNumber++;
  }

  /*
   * Blank separator row.
   */
  rowNumber++;

  return rowNumber;
}

/**
 * Standardizes class names.
 *
 * Examples:
 *   Baby, Baby class, Baby class list
 *      -> Baby class list
 *
 *   P.3, P3, Primary 3, Primary 3 list
 *      -> Primary 3 list
 *
 *   S.5, S5, Senior 5, Senior 5 list
 *      -> Senior 5 list
 */
function canonicalLeaversClass_(value) {
  let text =
    normalizeLeaversClassText_(value)
      .toUpperCase();

  if (!text) {
    return "";
  }

  text = text
    .replace(/[._\-\/]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const compact =
    text.replace(/\s+/g, "");

  /*
   * Nursery classes.
   */
  const nurseryPatterns = [
    {
      className: "Baby class list",
      values: [
        "BABY",
        "BABYCLASS",
        "BABYCLASSLIST",
        "BABYLIST",
      ],
    },
    {
      className: "Middle class list",
      values: [
        "MIDDLE",
        "MIDDLECLASS",
        "MIDDLECLASSLIST",
        "MIDDLELIST",
      ],
    },
    {
      className: "Top Class list",
      values: [
        "TOP",
        "TOPCLASS",
        "TOPCLASSLIST",
        "TOPLIST",
      ],
    },
  ];

  for (const item of nurseryPatterns) {
    if (item.values.includes(compact)) {
      return item.className;
    }
  }

  /*
   * Primary classes.
   */
  const primaryPatterns = [
    {
      className: "Primary 1 list",
      values: [
        "P1",
        "P1LIST",
        "PRIMARY1",
        "PRIMARY1LIST",
        "PRIMARYONE",
        "PRIMARYONELIST",
      ],
    },
    {
      className: "Primary 2 list",
      values: [
        "P2",
        "P2LIST",
        "PRIMARY2",
        "PRIMARY2LIST",
        "PRIMARYTWO",
        "PRIMARYTWOLIST",
      ],
    },
    {
      className: "Primary 3 list",
      values: [
        "P3",
        "P3LIST",
        "PRIMARY3",
        "PRIMARY3LIST",
        "PRIMARYTHREE",
        "PRIMARYTHREELIST",
      ],
    },
    {
      className: "Primary 4 list",
      values: [
        "P4",
        "P4LIST",
        "PRIMARY4",
        "PRIMARY4LIST",
        "PRIMARYFOUR",
        "PRIMARYFOURLIST",
      ],
    },
    {
      className: "Primary 5 list",
      values: [
        "P5",
        "P5LIST",
        "PRIMARY5",
        "PRIMARY5LIST",
        "PRIMARYFIVE",
        "PRIMARYFIVELIST",
      ],
    },
    {
      className: "Primary 6 list",
      values: [
        "P6",
        "P6LIST",
        "PRIMARY6",
        "PRIMARY6LIST",
        "PRIMARYSIX",
        "PRIMARYSIXLIST",
      ],
    },
    {
      className: "Primary 7 list",
      values: [
        "P7",
        "P7LIST",
        "PRIMARY7",
        "PRIMARY7LIST",
        "PRIMARYSEVEN",
        "PRIMARYSEVENLIST",
      ],
    },
  ];

  for (const item of primaryPatterns) {
    if (item.values.includes(compact)) {
      return item.className;
    }
  }

  /*
   * Senior classes.
   */
  const seniorPatterns = [
    {
      className: "Senior 1 list",
      values: [
        "S1",
        "S1LIST",
        "SENIOR1",
        "SENIOR1LIST",
        "SENIORONE",
        "SENIORONELIST",
      ],
    },
    {
      className: "Senior 2 list",
      values: [
        "S2",
        "S2LIST",
        "SENIOR2",
        "SENIOR2LIST",
        "SENIORTWO",
        "SENIORTWOLIST",
      ],
    },
    {
      className: "Senior 3 list",
      values: [
        "S3",
        "S3LIST",
        "SENIOR3",
        "SENIOR3LIST",
        "SENIORTHREE",
        "SENIORTHREELIST",
      ],
    },
    {
      className: "Senior 4 list",
      values: [
        "S4",
        "S4LIST",
        "SENIOR4",
        "SENIOR4LIST",
        "SENIORFOUR",
        "SENIORFOURLIST",
      ],
    },
    {
      className: "Senior 5 list",
      values: [
        "S5",
        "S5LIST",
        "SENIOR5",
        "SENIOR5LIST",
        "SENIORFIVE",
        "SENIORFIVELIST",
      ],
    },
    {
      className: "Senior 6 list",
      values: [
        "S6",
        "S6LIST",
        "SENIOR6",
        "SENIOR6LIST",
        "SENIORSIX",
        "SENIORSIXLIST",
      ],
    },
  ];

  for (const item of seniorPatterns) {
    if (item.values.includes(compact)) {
      return item.className;
    }
  }

  return "";
}

/**
 * Returns the title colour for each section.
 */
function getLeaversClassSectionColour_(className) {
  const key =
    normalizeLeaversClassKey_(
      className
    );

  /*
   * Nursery.
   */
  if (
    key.includes("BABY") ||
    key.includes("MIDDLE") ||
    key.includes("TOPCLASS")
  ) {
    return "#674ea7";
  }

  /*
   * Primary.
   */
  if (key.includes("PRIMARY")) {
    return "#3d85c6";
  }

  /*
   * Senior.
   */
  if (key.includes("SENIOR")) {
    return "#38761d";
  }

  return "#666666";
}

/**
 * Converts a balance to a valid number.
 *
 * Supports:
 *   750000
 *   750,000
 *   UGX 750,000
 *   750000.00
 *   blank cells
 */
function parseLeaversClassAmount_(value) {
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? value
      : 0;
  }

  const text =
    normalizeLeaversClassText_(value);

  if (!text) {
    return 0;
  }

  const cleaned =
    text.replace(/[^\d.-]/g, "");

  if (
    !cleaned ||
    cleaned === "-" ||
    cleaned === "." ||
    cleaned === "-."
  ) {
    return 0;
  }

  const amount =
    Number(cleaned);

  return Number.isFinite(amount)
    ? amount
    : 0;
}

/**
 * Sorts rows by:
 *   1. Date Left, latest first
 *   2. Student Name alphabetically
 */
function sortLeaversClassRows_(
  rows,
  nameIndex,
  dateLeftIndex
) {
  rows.sort((a, b) => {
    const dateA =
      dateLeftIndex != null
        ? leaversClassDateValue_(
            a[dateLeftIndex]
          )
        : 0;

    const dateB =
      dateLeftIndex != null
        ? leaversClassDateValue_(
            b[dateLeftIndex]
          )
        : 0;

    if (dateA !== dateB) {
      return dateB - dateA;
    }

    const nameA =
      normalizeLeaversClassText_(
        a[nameIndex]
      );

    const nameB =
      normalizeLeaversClassText_(
        b[nameIndex]
      );

    return nameA.localeCompare(nameB);
  });
}

/**
 * Converts a date to a timestamp.
 */
function leaversClassDateValue_(value) {
  if (
    value instanceof Date &&
    !isNaN(value)
  ) {
    return value.getTime();
  }

  const text =
    normalizeLeaversClassText_(value);

  if (!text) {
    return 0;
  }

  /*
   * dd/MM/yyyy
   */
  const dayFirstMatch =
    text.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );

  if (dayFirstMatch) {
    const day =
      Number(dayFirstMatch[1]);

    const month =
      Number(dayFirstMatch[2]) - 1;

    const year =
      Number(dayFirstMatch[3]);

    const date =
      new Date(
        year,
        month,
        day
      );

    return isNaN(date)
      ? 0
      : date.getTime();
  }

  const parsed =
    new Date(text);

  return isNaN(parsed)
    ? 0
    : parsed.getTime();
}

/**
 * Normalizes text.
 */
function normalizeLeaversClassText_(value) {
  return String(
    value == null ? "" : value
  )
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Normalizes headers, statuses and keys.
 */
function normalizeLeaversClassKey_(value) {
  return normalizeLeaversClassText_(value)
    .replace(/\s+/g, "")
    .toUpperCase();
}

/**
 * Makes sure the report contains enough columns.
 */
function ensureLeaversClassReportColumns_(
  sheet,
  requiredColumns
) {
  const currentColumns =
    sheet.getMaxColumns();

  if (currentColumns < requiredColumns) {
    sheet.insertColumnsAfter(
      currentColumns,
      requiredColumns - currentColumns
    );
  }
}

/**
 * Applies report column widths.
 */
function formatLeaversClassReportColumns_(
  sheet,
  headers
) {
  const widths = {
    "NO.": 55,
    "NO": 55,
    "DATEOFADMISSION": 115,
    "STUDENTNAME": 230,
    "SEX": 55,
    "NATIONALITY": 110,
    "DAY/BOARDING": 110,
    "LASTTERMBALANCE": 125,
    "FULLFEE": 100,
    "P(PAID)": 100,
    "PPAID": 100,
    "BALANCE": 110,
    "CAUTION": 90,
    "TEL": 110,
    "FORMERCLASS": 125,
    "DATELEFT": 105,
    "EXITREASON": 170,
    "STATUS": 95,
    "REINSTATEDTO": 120,
    "REINSTATEDAT": 135,
    "SOURCESHEET": 120,
    "SOURCEROW": 85,
    "CREATEDAT": 145,
    "UPDATEDAT": 145,
    "STUDENTID": 200,
  };

  headers.forEach(
    (header, index) => {
      const key =
        normalizeLeaversClassKey_(
          header
        );

      const width =
        widths[key] || 100;

      try {
        sheet.setColumnWidth(
          index + 1,
          width
        );
      } catch (_) {}
    }
  );
}

/**
 * Converts a column number into a letter.
 *
 * Examples:
 *   1  = A
 *   10 = J
 *   27 = AA
 */
function columnLetterLeaversClass_(
  columnNumber
) {
  let number =
    Number(columnNumber);

  let letters = "";

  while (number > 0) {
    const remainder =
      (number - 1) % 26;

    letters =
      String.fromCharCode(
        65 + remainder
      ) + letters;

    number =
      Math.floor(
        (number - 1) / 26
      );
  }

  return letters;
}

/**
 * Manual refresh function.
 */
function refreshLeaversByClassNow() {
  rebuildLeaversByClassReport();
}

/**
 * Installs a daily trigger at approximately 1:00 AM.
 *
 * Run once manually.
 */
function installLeaversByClassTrigger() {
  const handler =
    "rebuildLeaversByClassReport";

  /*
   * Remove existing triggers for this function.
   */
  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (
        trigger.getHandlerFunction() ===
        handler
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  ScriptApp
    .newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      "Daily Leavers-by-Class trigger installed for approximately 1:00 AM.",
      "Trigger Installed",
      7
    );
}

/**
 * Installs an hourly trigger.
 *
 * Run this instead of the daily trigger when required.
 */
function installHourlyLeaversByClassTrigger() {
  const handler =
    "rebuildLeaversByClassReport";

  /*
   * Remove existing triggers for this function.
   */
  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (
        trigger.getHandlerFunction() ===
        handler
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  ScriptApp
    .newTrigger(handler)
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      "Hourly Leavers-by-Class trigger installed.",
      "Trigger Installed",
      6
    );
}

/**
 * Removes all triggers connected to this report.
 */
function removeLeaversByClassTriggers() {
  const handler =
    "rebuildLeaversByClassReport";

  let removed = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (
        trigger.getHandlerFunction() ===
        handler
      ) {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      `${removed} Leavers-by-Class trigger(s) removed.`,
      "Triggers Removed",
      6
    );
}