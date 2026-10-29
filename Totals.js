/***** ===== CLASS LISTS — TERM AUTOMATION (TOTALS) =====
 * v1.7-totals-manual (MANUAL totals row + NO auto writes)
 *
 * POLICY (matches Config v1.6 + Utils v1.7):
 * ✅ TOTALS row is MANUAL: we locate it by the label text (CL.TOTALS.LABEL_TEXT or "TOTALS")
 * ✅ No fixed totals anchors unless CL.TOTALS.USE_FIXED_ANCHORS === true
 * ✅ No automatic totals formula writing unless CL.TOTALS.AUTO_WRITE_FORMULAS === true OR force=true passed
 *
 * WHY THIS FIXES YOUR ISSUE:
 * - Previously, fixed totals anchors + row shifts caused totals formulas to land inside student rows.
 * - Now totals formulas cannot be written automatically, and totals row detection is label-based + bottom-up.
 ************************************************************/

/** Local spreadsheet getter */
function TOTS_ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* =========================
   Safe local helpers
========================= */

function TOTS_norm_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().replace(/\s+/g, " ");
}
function TOTS_flatUpper_(v) {
  return TOTS_norm_(v).toUpperCase().replace(/\s+/g, "");
}
function TOTS_toast_(msg, title, seconds) {
  try { if (typeof toast_ === "function") return toast_(msg, title, seconds); } catch (_) {}
  try { TOTS_ss_().toast(String(msg || ""), String(title || "Info"), Number(seconds || 5)); } catch (_) {}
}

/** Read totals policy from CL (safe defaults) */
function TOTS_policy_() {
  const label = String(CL?.TOTALS?.LABEL_TEXT || "TOTALS");
  const useFixed = CL?.TOTALS?.USE_FIXED_ANCHORS === true;
  const autoWrite = CL?.TOTALS?.AUTO_WRITE_FORMULAS === true;
  return {
    labelKey: TOTS_flatUpper_(label),
    useFixedAnchors: useFixed,
    autoWriteFormulas: autoWrite,
  };
}

/** Convert 1-based column index to letter (A, B, ... AA...) */
function TOTS_colLetter_(colIndex) {
  // Prefer project helper if it exists AND returns a valid A1 letter
  try {
    if (typeof colLetter_ === "function") {
      const v = colLetter_(colIndex);
      if (typeof v === "string" && /^[A-Z]+$/.test(v)) return v;
    }
  } catch (_) {}

  const n = Number(colIndex);
  if (!Number.isFinite(n) || n < 1) throw new Error(`Invalid column index: ${colIndex}`);
  let x = Math.floor(n);
  let s = "";
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/** Header index map (exact header text -> 1-based col) */
function TOTS_getHeaderIndexMap_(sh) {
  try {
    if (typeof getHeaderIndexMap_ === "function") {
      const m = getHeaderIndexMap_(sh);
      if (m && typeof m === "object") return m;
    }
  } catch (_) {}

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(TOTS_norm_);
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (!h) continue;
    if (map[h] == null) map[h] = i + 1;
  }
  return map;
}

/* =========================
   Totals row detection (MANUAL + bottom-up)
========================= */

function TOTS_isTotalsLabel_(v) {
  const p = TOTS_policy_();
  const key = TOTS_flatUpper_(v);

  if (key === p.labelKey) return true;

  // If label is TOTALS, allow things like "TOTALS:" / "TOTALS -" etc
  if (p.labelKey === "TOTALS" && key.startsWith("TOTALS")) return true;

  return false;
}

/** Prefer CL.FIXED_TOTALS[name].row only if policy allows */
function TOTS_fixedTotalsRowIfAllowed_(sh) {
  const p = TOTS_policy_();
  if (!p.useFixedAnchors) return null;

  try {
    const name = sh.getName();
    const r = Number(CL?.FIXED_TOTALS?.[name]?.row);
    if (Number.isFinite(r) && r >= 2) return r;
  } catch (_) {}
  return null;
}

/** Narrow find TOTALS row (col A only) — bottom-up */
function TOTS_findTotalsRow_(sh) {
  // Prefer external helper if it exists (updated Utils respects manual policy)
  try {
    if (typeof findTotalsRow_ === "function") {
      const r = findTotalsRow_(sh);
      if (r) return r;
    }
  } catch (_) {}

  const fixed = TOTS_fixedTotalsRowIfAllowed_(sh);
  if (fixed) return fixed;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const maxScan = Math.min(lastRow, 4000);
  const colA = sh.getRange(1, 1, maxScan, 1).getValues();

  for (let i = colA.length - 1; i >= 0; i--) {
    if (TOTS_isTotalsLabel_(colA[i][0])) return i + 1;
  }
  return null;
}

/** Wide find TOTALS row (robust) — bottom-up */
function TOTS_findTotalsRowWide_(sh) {
  // Prefer external helper if it exists (updated Utils respects manual policy)
  try {
    if (typeof findTotalsRowWide_ === "function") {
      const r = findTotalsRowWide_(sh, 12);
      if (r) return r;
    }
  } catch (_) {}

  const fixed = TOTS_fixedTotalsRowIfAllowed_(sh);
  if (fixed) return fixed;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const scanRows = Math.min(lastRow, 4000);
  const scanCols = Math.min(12, Math.max(1, sh.getLastColumn()));
  const vals = sh.getRange(1, 1, scanRows, scanCols).getValues();

  for (let r = vals.length - 1; r >= 0; r--) {
    for (let c = 0; c < vals[r].length; c++) {
      if (TOTS_isTotalsLabel_(vals[r][c])) return r + 1;
    }
  }
  return null;
}

/** Smart totals row: narrow -> wide */
function TOTS_findTotalsRowSmart_(sh) {
  return TOTS_findTotalsRow_(sh) || TOTS_findTotalsRowWide_(sh) || null;
}

/* =========================
   Class-sheet selection (deterministic)
========================= */

function TOTS_isDefaultersName_(name) {
  const n = String(name || "").trim().toLowerCase();
  const suf = String(CL?.DEFS_SUFFIX || " Defaulters").trim().toLowerCase();
  if (suf && n.endsWith(suf)) return true;
  return /defaulters$/i.test(name || "");
}

/** Strict class sheets list: getClassSheets_() OR CL.CLASS_ORDER OR suffix-only fallback */
function TOTS_getClassSheets_() {
  const ss = TOTS_ss_();

  try {
    if (typeof getClassSheets_ === "function") {
      const res = getClassSheets_();
      if (Array.isArray(res)) return res;
    }
  } catch (_) {}

  try {
    if (CL && Array.isArray(CL.CLASS_ORDER) && CL.CLASS_ORDER.length) {
      const out = [];
      CL.CLASS_ORDER.forEach(n => {
        const sh = ss.getSheetByName(String(n));
        if (sh) out.push(sh);
      });
      return out;
    }
  } catch (_) {}

  const suf = String(CL?.CLASS_SUFFIX || " list").toLowerCase();
  return ss.getSheets().filter(sh => {
    const nm = sh.getName();
    if (TOTS_isDefaultersName_(nm)) return false;
    return String(nm).toLowerCase().endsWith(suf);
  });
}

/* =========================
   Optional NO-GREEN enforcement
========================= */

function TOTS_removeConditionalFormatting_(sh) {
  try {
    if (!sh) return;
    if (TOTS_isDefaultersName_(sh.getName())) return;
    const rules = sh.getConditionalFormatRules();
    if (rules && rules.length) sh.setConditionalFormatRules([]);
  } catch (_) {}
}

/* =========================
   Safety cleanup (misplaced totals formulas)
========================= */

/**
 * Detects the totals-type formula we write (signature-based).
 * We only clear formulas that match this signature in student rows.
 */
function TOTS_isTotalsFormulaSignature_(f) {
  const s = String(f || "").replace(/\s+/g, "").toUpperCase();
  if (!s) return false;

  // Our totals formula pattern has SUM(...) and ROW()-1 logic (or similar).
  const hasSum = s.includes("SUM(");
  const hasRow = s.includes("ROW()");
  const hasRowMinus1 = s.includes("ROW()-1") || s.includes("ROW()-1)");
  const hasIndexOrIndirect = s.includes("INDEX(") || s.includes("INDIRECT(");

  // Also common prefix used in older version:
  const hasIfRow3 = s.includes("IF(ROW()<3");

  return hasSum && (hasRowMinus1 || (hasRow && hasIndexOrIndirect) || hasIfRow3);
}

/**
 * Clear misplaced totals formulas in student rows for totals columns.
 * Only clears cells whose formula matches totals signature.
 */
function TOTS_cleanupMisplacedTotalsFormulas_(sh, totalsRow, cols) {
  if (!sh || !totalsRow || totalsRow < 3) return;
  if (!cols || !cols.length) return;

  const startRow = 2;
  const endRow = totalsRow - 1;
  if (endRow < startRow) return;

  cols.forEach(colIndex => {
    try {
      const rng = sh.getRange(startRow, colIndex, endRow - startRow + 1, 1);
      const formulas = rng.getFormulas(); // 2D
      let any = false;

      for (let i = 0; i < formulas.length; i++) {
        const f = formulas[i][0];
        if (TOTS_isTotalsFormulaSignature_(f)) {
          formulas[i][0] = ""; // clear this formula
          any = true;
        }
      }

      if (any) rng.setFormulas(formulas);
    } catch (_) {}
  });
}

/* =========================
   Public API (MANUAL by default)
========================= */

/**
 * Fix totals on all class sheets.
 * IMPORTANT: NO-OP unless policy allows auto write OR force=true passed.
 *
 * @param {{force?:boolean, enforceNoGreen?:boolean, cleanupMisplaced?:boolean}=} opt
 */
function fixTotalsAllSheets(opt) {
  const p = TOTS_policy_();
  const force = opt?.force === true;

  if (!p.autoWriteFormulas && !force) {
    TOTS_toast_("Totals auto-write is OFF (manual mode). No changes made.", "Totals", 6);
    return;
  }

  const sheets = TOTS_getClassSheets_();
  const errors = [];

  sheets.forEach(sh => {
    try {
      fixTotalsOne_(sh, opt);
    } catch (e) {
      errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  SpreadsheetApp.flush();

  if (errors.length) {
    TOTS_toast_(`Totals finished with ${errors.length} issue(s). Check logs.`, "Totals", 8);
    Logger.log("fixTotalsAllSheets() issues:\n" + errors.join("\n"));
    return;
  }

  TOTS_toast_("Totals applied to all class sheets.", "Totals", 6);
}

/**
 * Fix totals on a single class sheet.
 * IMPORTANT: NO-OP unless policy allows auto write OR force=true passed.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {{force?:boolean, enforceNoGreen?:boolean, cleanupMisplaced?:boolean}=} opt
 */
function fixTotalsOne_(sh, opt) {
  if (!sh) throw new Error("No sheet provided.");
  if (TOTS_isDefaultersName_(sh.getName())) return;

  const p = TOTS_policy_();
  const force = opt?.force === true;

  if (!p.autoWriteFormulas && !force) {
    // Manual mode: do nothing when called indirectly (prevents the random damage).
    return;
  }

  if (opt?.enforceNoGreen === true) TOTS_removeConditionalFormatting_(sh);

  // Ensure structure exists (safe; does not write totals)
  try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}

  const map = TOTS_getHeaderIndexMap_(sh);

  const cFEE  = map[CL.H.FEE];
  const cPAID = map[CL.H.PAID];
  const cBAL  = map[CL.H.BAL];
  const cLTB  = map[CL.H.LTB]; // optional totals

  if (!cFEE || !cPAID || !cBAL) {
    throw new Error(`Missing expected headers: "${CL.H.FEE}", "${CL.H.PAID}", "${CL.H.BAL}"`);
  }

  const totalsRow = TOTS_findTotalsRowSmart_(sh);
  if (!totalsRow) throw new Error(`No "${CL?.TOTALS?.LABEL_TEXT || "TOTALS"}" row found.`);

  // Optional cleanup: remove misplaced totals formulas above totals row (safe signature-based)
  if (opt?.cleanupMisplaced === true) {
    const colsToCheck = [cLTB, cFEE, cPAID, cBAL].filter(Boolean);
    TOTS_cleanupMisplacedTotalsFormulas_(sh, totalsRow, colsToCheck);
  }

  // Formula placed ONLY in totalsRow; sums rows 2..(totalsRow-1)
  // Uses INDEX instead of INDIRECT (less volatile)
  function setSumAtTotalsRow_(colIndex) {
    const L = TOTS_colLetter_(colIndex);
    const f = `=IF(ROW()<3,0,SUM(${L}2:INDEX(${L}:${L},ROW()-1)))`;
    sh.getRange(totalsRow, colIndex).setFormula(f);
  }

  setSumAtTotalsRow_(cFEE);
  setSumAtTotalsRow_(cPAID);
  setSumAtTotalsRow_(cBAL);
  if (cLTB) setSumAtTotalsRow_(cLTB);
}

/**
 * LEGACY: Fixed-position totals writer.
 * Disabled unless CL.TOTALS.USE_FIXED_ANCHORS === true AND force=true is passed.
 *
 * @param {{force?:boolean, enforceNoGreen?:boolean}=} opt
 */
function fixTotals_UseFixedPositions(opt) {
  const p = TOTS_policy_();
  const force = opt?.force === true;

  if (!p.useFixedAnchors || !force) {
    TOTS_toast_("Fixed-position totals is DISABLED (manual totals mode). No changes made.", "Totals", 7);
    return;
  }

  const ss = TOTS_ss_();
  const fixed = (CL && CL.FIXED_TOTALS) ? CL.FIXED_TOTALS : {};
  const names = Object.keys(fixed || {});
  const errors = [];

  names.forEach(name => {
    try {
      const cfg = fixed[name];
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      if (TOTS_isDefaultersName_(sh.getName())) return;

      if (opt?.enforceNoGreen === true) TOTS_removeConditionalFormatting_(sh);
      try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}

      const r = Number(cfg.row);
      if (!Number.isFinite(r) || r < 2) throw new Error(`Invalid fixed totals row for "${name}"`);

      const labelCell = sh.getRange(r, 1);
      if (!TOTS_isTotalsLabel_(labelCell.getValue())) labelCell.setValue(CL?.TOTALS?.LABEL_TEXT || "TOTALS");

      const feeF  = `=SUM(${cfg.fee}2:${cfg.fee}${r - 1})`;
      const paidF = `=SUM(${cfg.paid}2:${cfg.paid}${r - 1})`;
      const balF  = `=SUM(${cfg.bal}2:${cfg.bal}${r - 1})`;
      sh.getRange(`${cfg.fee}${r}`).setFormula(feeF);
      sh.getRange(`${cfg.paid}${r}`).setFormula(paidF);
      sh.getRange(`${cfg.bal}${r}`).setFormula(balF);

      if (cfg.ltb) {
        const ltbF = `=SUM(${cfg.ltb}2:${cfg.ltb}${r - 1})`;
        sh.getRange(`${cfg.ltb}${r}`).setFormula(ltbF);
      }
    } catch (e) {
      errors.push(`${name}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    TOTS_toast_(`Fixed-position totals finished with ${errors.length} issue(s). Check logs.`, "Totals", 8);
    Logger.log("fixTotals_UseFixedPositions() issues:\n" + errors.join("\n"));
    return;
  }

  TOTS_toast_("Fixed-position totals written.", "Totals", 6);
}

/**
 * Formatting helper (manual by default):
 * - Optional NO-GREEN enforcement
 * - Optional trim below TOTALS (safe)
 *
 * NO-OP unless policy allows auto write OR force=true passed.
 *
 * @param {{force?:boolean, enforceNoGreen?:boolean, trimBelowTotals?:boolean}=} opt
 */
function fixTotalsFormattingAll_(opt) {
  const p = TOTS_policy_();
  const force = opt?.force === true;

  if (!p.autoWriteFormulas && !force) {
    TOTS_toast_("Totals/format auto-run is OFF (manual mode). No changes made.", "Format Fix", 6);
    return;
  }

  const sheets = TOTS_getClassSheets_();
  const errors = [];

  sheets.forEach(sh => {
    try {
      if (TOTS_isDefaultersName_(sh.getName())) return;

      if (opt?.enforceNoGreen === true) TOTS_removeConditionalFormatting_(sh);

      if (opt?.trimBelowTotals === true) {
        if (typeof trimBelowTotals_One === "function") trimBelowTotals_One(sh);
        else TOTS_trimBelowTotalsFallback_(sh);
      }
    } catch (e) {
      errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    TOTS_toast_(`Format fix ran with ${errors.length} issue(s). Check logs.`, "Format Fix", 8);
    Logger.log("fixTotalsFormattingAll_() issues:\n" + errors.join("\n"));
    return;
  }

  TOTS_toast_("Format fix complete.", "Format Fix", 6);
}

/** Fallback: clear contents below TOTALS without deleting physical rows. */
function TOTS_trimBelowTotalsFallback_(sh) {
  const tRow = TOTS_findTotalsRowSmart_(sh);
  if (!tRow) return;
  const extra = sh.getMaxRows() - tRow;
  if (extra > 0) {
    const lastCol = Math.max(1, sh.getLastColumn());
    sh.getRange(tRow + 1, 1, extra, lastCol).clearContent();
  }
}

/* =========================
   MANUAL convenience entry points (recommended)
========================= */

/**
 * MANUAL: Apply totals formulas to ALL class sheets.
 * This is the function you should run manually when you WANT totals formulas.
 */
function MANUAL_WriteTotals_AllClassSheets() {
  fixTotalsAllSheets({ force: true, enforceNoGreen: false, cleanupMisplaced: true });
  TOTS_toast_("Manual totals applied (forced).", "Totals", 6);
}

/**
 * MANUAL: Apply totals formulas to ONE sheet by name (safe).
 */
function MANUAL_WriteTotals_OneSheetByName(sheetName) {
  const sh = TOTS_ss_().getSheetByName(String(sheetName || ""));
  if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);
  fixTotalsOne_(sh, { force: true, enforceNoGreen: false, cleanupMisplaced: true });
  TOTS_toast_(`Manual totals applied: ${sh.getName()}`, "Totals", 6);
}