/***** ===== CLASS LISTS — TERM AUTOMATION (UTILS) =====
 * Hardened Utils.gs (safe, deterministic, no “random sheet” accidents).
 *
 * KEY FIXES:
 * ✅ getClassSheets_() is now DETERMINISTIC:
 *    - Uses CL.CLASS_ORDER ONLY (no dangerous fallback that includes other sheets).
 *    - Optional strict fallback: only sheets ending with CL.CLASS_SUFFIX (still excludes system/defaulters).
 * ✅ Totals detection is safer:
 *    - Prefers CL.FIXED_TOTALS[sheetName].row when available
 *    - Otherwise scans for "TOTALS" (fast col A, then wide scan up to 12 cols)
 * ✅ Header mapping helpers added:
 *    - getHeaderIndexMap_() keeps exact keys (backwards compatible)
 *    - getHeaderIndexMapUpper_() provides case-insensitive matching when needed
 * ✅ Student region is always rows 2..(totalsRow-1) (never includes totals row)
 * ✅ Small guards to reduce timeouts: scan limits + minimal ranges
 *
 * Provides:
 *  - ss_(), toast_()
 *  - _normText_(), _flatUpper_()
 *  - colLetter_()
 *  - getHeaderIndexMap_(), getHeaderIndexMapUpper_()
 *  - findTotalsRow_(), findTotalsRowWide_(), findTotalsRowSmart_()
 *  - getStudentRegion_()
 *  - nextSerialNo_(), nextRowAfterLastNumberedA_()
 *  - withDocLock_()
 *  - normalizeTargetClassName_()
 *  - isDefaultersSheetName_(), isSystemSheetName_()
 *  - getClassSheets_()  (safe)
 ************************************************************/

/** Active spreadsheet wrapper */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Toast wrapper (never throws) */
function toast_(msg, title, seconds) {
  try {
    ss_().toast(String(msg || ""), String(title || "Info"), Number(seconds || 5));
  } catch (_) {}
}

/** Normalize text: trim + collapse whitespace */
function _normText_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().replace(/\s+/g, " ");
}

/** Upper + remove spaces (good for matching headers/names) */
function _flatUpper_(v) {
  return _normText_(v).toUpperCase().replace(/\s+/g, "");
}

/** Column index (1-based) to letter (A, B, ..., AA, AB...) */
function colLetter_(colIndex) {
  let n = Number(colIndex) || 0;
  if (n <= 0) throw new Error("colLetter_: invalid colIndex " + colIndex);

  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* =========================
 * Sheet-name helpers (safe)
 * ========================= */

function isDefaultersSheetName_(name) {
  const n = String(name || "").trim().toLowerCase();
  const suf = (typeof CL !== "undefined" && CL && CL.DEFS_SUFFIX) ? String(CL.DEFS_SUFFIX).trim().toLowerCase() : " defaulters";
  if (suf && n.endsWith(suf)) return true;
  return /defaulters$/i.test(String(name || ""));
}

function _systemNamesSet_() {
  const s = new Set();

  // Canonical
  try { if (CL?.TERM?.LOG_SHEET) s.add(String(CL.TERM.LOG_SHEET)); } catch (_) {}
  try { if (CL?.LEAVERS?.SHEET) s.add(String(CL.LEAVERS.SHEET)); } catch (_) {}
  try { if (CL?.SHEETNAMES?.SUMMARIES) s.add(String(CL.SHEETNAMES.SUMMARIES)); } catch (_) {}
  try { if (CL?.SHEETNAMES?.GRADUATES) s.add(String(CL.SHEETNAMES.GRADUATES)); } catch (_) {}

  // Optional centralized list
  try {
    if (Array.isArray(CL?.SYSTEM_SHEETS)) CL.SYSTEM_SHEETS.forEach(n => s.add(String(n)));
  } catch (_) {}

  // Aliases (if provided)
  try {
    const a = CL?.SHEET_ALIASES || {};
    Object.keys(a).forEach(k => {
      const arr = a[k];
      if (Array.isArray(arr)) arr.forEach(n => s.add(String(n)));
    });
  } catch (_) {}

  // Common fallbacks
  ["SUMMARIES", "Summaries", "CONFIG", "Config", "Graduates", "_TermLog_CLASSLISTS"].forEach(n => s.add(n));

  return s;
}

function isSystemSheetName_(name) {
  const nm = String(name || "");
  const sys = _systemNamesSet_();
  return sys.has(nm);
}

/** Basic sheet name normalizer (safe default) */
function normalizeTargetClassName_(name) {
  return _normText_(name);
}

/* =========================
 * Headers
 * ========================= */

/**
 * Map header text (row 1) -> column index (1-based).
 * Keys are EXACT header text (trimmed). Backwards compatible.
 */
function getHeaderIndexMap_(sh) {
  if (!sh) throw new Error("getHeaderIndexMap_: sheet required");

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const h = _normText_(headers[c]);
    if (!h) continue;
    if (map[h] == null) map[h] = c + 1;
  }
  return map;
}

/**
 * Header map using normalized uppercase keys (case-insensitive lookups).
 * Example: map[_flatUpper_("Full Fee")] -> column index
 */
function getHeaderIndexMapUpper_(sh) {
  if (!sh) throw new Error("getHeaderIndexMapUpper_: sheet required");

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const key = _flatUpper_(headers[c]);
    if (!key) continue;
    if (map[key] == null) map[key] = c + 1;
  }
  return map;
}

/* =========================
 * Totals row detection
 * ========================= */

/** Find TOTALS row by scanning column A (fast). Returns row number or null. */
function findTotalsRow_(sh) {
  if (!sh) return null;

  // Prefer fixed totals row if configured
  try {
    const cfg = CL?.FIXED_TOTALS?.[sh.getName()];
    const r = Number(cfg?.row);
    if (Number.isFinite(r) && r >= 2) return r;
  } catch (_) {}

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const maxScan = Math.min(lastRow, 4000);
  const colA = sh.getRange(1, 1, maxScan, 1).getValues();

  for (let r = 0; r < colA.length; r++) {
    if (_flatUpper_(colA[r][0]) === "TOTALS") return r + 1;
  }
  return null;
}

/** Find TOTALS row by scanning a wider area (robust). Returns row number or null. */
function findTotalsRowWide_(sh) {
  if (!sh) return null;

  // Prefer fixed totals row if configured
  try {
    const cfg = CL?.FIXED_TOTALS?.[sh.getName()];
    const r = Number(cfg?.row);
    if (Number.isFinite(r) && r >= 2) return r;
  } catch (_) {}

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const scanRows = Math.min(lastRow, 4000);
  const scanCols = Math.min(12, Math.max(1, sh.getLastColumn()));
  const vals = sh.getRange(1, 1, scanRows, scanCols).getValues();

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      if (_flatUpper_(vals[r][c]) === "TOTALS") return r + 1;
    }
  }
  return null;
}

/** Smart totals row: fixed (if valid) -> col A scan -> wide scan */
function findTotalsRowSmart_(sh) {
  return findTotalsRow_(sh) || findTotalsRowWide_(sh) || null;
}

/**
 * Student region = rows 2..(totalsRow-1) if totals exists, else 2..lastRow
 * Returns: { startRow, endRow, totalsRow }
 *
 * NOTE:
 * - endRow is the LAST STUDENT ROW (never the totals row).
 */
function getStudentRegion_(sh) {
  if (!sh) throw new Error("getStudentRegion_: sheet required");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { startRow: 2, endRow: 1, totalsRow: null };

  const totalsRow = findTotalsRowSmart_(sh);
  const endRow = totalsRow ? (totalsRow - 1) : lastRow;

  return {
    startRow: 2,
    endRow: Math.max(1, endRow),
    totalsRow: totalsRow || null,
  };
}

/* =========================
 * Serial / append helpers
 * ========================= */

/** Next serial number based on max numeric in column A above TOTALS */
function nextSerialNo_(sh) {
  const region = getStudentRegion_(sh);
  const endRow = region.endRow;
  if (!endRow || endRow < 2) return 1;

  const vals = sh.getRange(2, 1, endRow - 1, 1).getValues().flat();
  let max = 0;
  for (const v of vals) {
    const n = Number(v);
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max + 1;
}

/**
 * Next row after last numbered value in column A (above TOTALS).
 * If none, returns 2.
 */
function nextRowAfterLastNumberedA_(sh) {
  const region = getStudentRegion_(sh);
  const endRow = region.endRow;
  if (!endRow || endRow < 2) return 2;

  const vals = sh.getRange(2, 1, endRow - 1, 1).getValues().flat();
  let lastIdx = -1;
  for (let i = 0; i < vals.length; i++) {
    const n = Number(vals[i]);
    if (Number.isFinite(n) && n > 0) lastIdx = i;
  }
  return 2 + (lastIdx + 1);
}

/* =========================
 * Locking
 * ========================= */

/** Document lock wrapper (prevents simultaneous runs) */
function withDocLock_(fn) {
  const lock = LockService.getDocumentLock();
  const start = Date.now();
  const maxMs = 30 * 1000;

  while (true) {
    try {
      if (lock.tryLock(3000)) break;
    } catch (_) {}

    if (Date.now() - start > maxMs) {
      throw new Error(
        "Lock timeout: another script execution is holding the document lock. " +
        "End other runs (Apps Script → Executions) and retry."
      );
    }
    Utilities.sleep(600);
  }

  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* =========================
 * Class sheet selection (CRITICAL)
 * ========================= */

/**
 * Get class list sheets SAFELY.
 *
 * Behavior:
 *  1) Primary: CL.CLASS_ORDER (deterministic; safest)
 *  2) Strict fallback (only if CL.CLASS_ORDER missing/empty):
 *     - include ONLY sheets that end with CL.CLASS_SUFFIX
 *     - exclude defaulters + system sheets
 *
 * NO “include anything non-system” fallback — that caused many of your problems.
 */
function getClassSheets_() {
  const ss = ss_();

  // 1) Deterministic list (recommended)
  try {
    if (typeof CL !== "undefined" && CL && Array.isArray(CL.CLASS_ORDER) && CL.CLASS_ORDER.length) {
      const out = [];
      CL.CLASS_ORDER.forEach(n => {
        const sh = ss.getSheetByName(String(n));
        if (sh) out.push(sh);
      });
      return out; // return even if empty; empty = configuration/name mismatch to be fixed
    }
  } catch (_) {}

  // 2) Strict fallback: only “... list”
  const suf = (typeof CL !== "undefined" && CL && CL.CLASS_SUFFIX) ? String(CL.CLASS_SUFFIX) : " list";
  const sufLower = String(suf).toLowerCase();

  return ss.getSheets().filter(sh => {
    const name = sh.getName();
    if (isSystemSheetName_(name)) return false;
    if (isDefaultersSheetName_(name)) return false;
    return String(name).toLowerCase().endsWith(sufLower);
  });
}
