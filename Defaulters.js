/***** ===== CLASS LISTS — TERM AUTOMATION (DEFAULTERS) =====
 * Purpose:
 * - Build/refresh each "<Class> Defaulters" sheet using ONE formula.
 * - Defaulters = student rows where BALANCE > 0 (STRICTLY from rows ABOVE TOTALS).
 * - Appends ONE computed TOTALS row (numbers only in LTB/FEE/PAID/BAL columns).
 *
 * HARD FIXES INCLUDED IN THIS UPDATE:
 * ✅ Prevent double TOTALS even if totals-row detection fails:
 *    - Adds an explicit "NOT TOTALS row" filter based on the DATE/first column text.
 * ✅ Keep DATE formatting:
 *    - Uses clearContents() (not clear()) so formats (including date formats) are preserved.
 *    - Re-applies a date format to column A after formula set (safe even if already correct).
 *
 * Depends on:
 * - Config.gs (CL)
 * - Utils.gs (ss_, getClassSheets_, getHeaderIndexMap_, colLetter_, findTotalsRowWide_, toast_)
 ************************************************************/

/** Build the formula for ONE defaulters sheet from a given class list sheet */
function makeDefaultersFormula_(sheetName) {
  if (!sheetName) throw new Error("sheetName is required");

  const ss = ss_();
  const src = ss.getSheetByName(sheetName);
  if (!src) throw new Error(`No sheet named "${sheetName}"`);

  const map = getHeaderIndexMap_(src);

  // Output range: DATE..REMARKS (keeps your layout)
  const cS = map[CL.H.DATE];
  const cE = map[CL.H.REM];

  // Money columns used for filtering & totals
  const cLTB = map[CL.H.LTB];
  const cFEE = map[CL.H.FEE];
  const cPAI = map[CL.H.PAID];
  const cBAL = map[CL.H.BAL];

  if (!cS || !cE || !cLTB || !cFEE || !cPAI || !cBAL) {
    throw new Error(
      `"${sheetName}": missing headers (DATE/REMARKS/LTB/FEE/PAID/BALANCE). Check row 1 exactly.`
    );
  }

  // STRICT limit: rows above TOTALS (wide detection)
  const tRow = findTotalsRowWide_(src);
  const lastRow = src.getLastRow();
  const endRow = tRow ? (tRow - 1) : lastRow; // last student row (best effort)
  const dataEndRow = Math.max(2, endRow);     // keep ranges valid (at least row 2)

  const qSheet = `'${String(sheetName).replace(/'/g, "''")}'`;

  const Ls = colLetter_(cS);
  const Le = colLetter_(cE);

  const headerRange = `${qSheet}!${Ls}1:${Le}1`;
  const dataRange   = `${qSheet}!${Ls}2:${Le}${dataEndRow}`;

  const nCols = cE - cS + 1;

  // Column letters for numeric conversion
  const L_LTB  = colLetter_(cLTB);
  const L_FEE  = colLetter_(cFEE);
  const L_PAID = colLetter_(cPAI);
  const L_BAL  = colLetter_(cBAL);

  // Numeric BAL range (works for numbers OR "100,000" text)
  const balRange = `${qSheet}!${L_BAL}2:${L_BAL}${dataEndRow}`;
  const balNum   = `IFERROR(VALUE(REGEXREPLACE(TO_TEXT(${balRange}),",","")),0)`;

  // EXTRA SAFETY: explicitly exclude any row where the first output column contains TOTALS
  // (Prevents double totals if totals row was included in the source range.)
  const firstColRange = `${qSheet}!${Ls}2:${Ls}${dataEndRow}`;
  const isTotalsRow   = `REGEXMATCH(TO_TEXT(${firstColRange}),"(?i)^\\s*TOTALS\\b")`;
  const notTotalsRow  = `NOT(${isTotalsRow})`;

  // If FILTER finds no matches, show this placeholder row instead of an error
  const noneCells = Array.from({ length: nCols }, (_, i) => (i === 0 ? `"NONE"` : `""`));
  const noneRow = `{${noneCells.join(",")}}`;

  // Defaulters rows (ALWAYS returns something: either rows, or ONE "NONE" row)
  const defRows = `IFERROR(FILTER(${dataRange}, (${balNum} > 0) * ${notTotalsRow}), ${noneRow})`;

  // Totals computed using the SAME numeric conversion + same BAL>0 condition + same NOT TOTALS row
  function colNum_(L) {
    const r = `${qSheet}!${L}2:${L}${dataEndRow}`;
    return `IFERROR(VALUE(REGEXREPLACE(TO_TEXT(${r}),",","")),0)`;
  }

  const ltbNum  = colNum_(L_LTB);
  const feeNum  = colNum_(L_FEE);
  const paidNum = colNum_(L_PAID);

  const tLTB  = `IFERROR(SUM(FILTER(${ltbNum},  (${balNum} > 0) * ${notTotalsRow})),0)`;
  const tFEE  = `IFERROR(SUM(FILTER(${feeNum},  (${balNum} > 0) * ${notTotalsRow})),0)`;
  const tPAID = `IFERROR(SUM(FILTER(${paidNum}, (${balNum} > 0) * ${notTotalsRow})),0)`;
  const tBAL  = `IFERROR(SUM(FILTER(${balNum},  (${balNum} > 0) * ${notTotalsRow})),0)`;

  // Positions of money columns inside the output (DATE..REMARKS), 1-based
  const rLTB = cLTB - cS + 1;
  const rFEE = cFEE - cS + 1;
  const rPAI = cPAI - cS + 1;
  const rBAL = cBAL - cS + 1;

  // Build one totals row aligned to the output columns
  const totalsCells = [];
  for (let i = 1; i <= nCols; i++) {
    if (i === 1) totalsCells.push(`"TOTALS"`);
    else if (i === rLTB) totalsCells.push(tLTB);
    else if (i === rFEE) totalsCells.push(tFEE);
    else if (i === rPAI) totalsCells.push(tPAID);
    else if (i === rBAL) totalsCells.push(tBAL);
    else totalsCells.push(`""`);
  }
  const totalsRow = `{${totalsCells.join(",")}}`;

  // Final output: header + rows (or NONE) + totals
  return `={${headerRange}; ${defRows}; ${totalsRow}}`;
}

/** Rebuild defaulters sheet for one class list */
function rebuildOneDefaulters(sheetName) {
  const ss = ss_();

  // Allow manual run with no parameter: uses active sheet if it's a class list.
  if (!sheetName) {
    const active = ss.getActiveSheet().getName();
    if (String(active).endsWith(CL.CLASS_SUFFIX)) sheetName = active;
    else throw new Error(`sheetName is required (or run while inside a "...${CL.CLASS_SUFFIX}" sheet).`);
  }

  const src = ss.getSheetByName(sheetName);
  if (!src) throw new Error(`No sheet named "${sheetName}"`);

  const defName = sheetName.replace(new RegExp(String(CL.CLASS_SUFFIX).replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "$"), CL.DEFS_SUFFIX);
  const def = ss.getSheetByName(defName) || ss.insertSheet(defName);

  // IMPORTANT: keep formats (especially date formats) -> do NOT use clear()
  def.clearContents();

  def.getRange(1, 1).setFormula(makeDefaultersFormula_(sheetName));
  def.setFrozenRows(1);

  // Re-apply a safe date format to col A (DATE column in your layout)
  // (Does not harm if already correct; helps when formats were lost historically.)
  try {
    const maxRows = Math.max(2, def.getMaxRows());
    def.getRange(2, 1, maxRows - 1, 1).setNumberFormat("dd/mm/yyyy");
  } catch (_) {}

  return defName;
}

/** Rebuild all defaulters sheets for all class lists */
function rebuildAllDefaulters_QUERY() {
  const ss = ss_();
  const sheets = getClassSheets_();

  const errors = [];
  sheets.forEach(src => {
    try {
      rebuildOneDefaulters(src.getName());
    } catch (e) {
      errors.push(`${src.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    toast_(`Defaulters rebuilt with ${errors.length} issue(s). Check logs.`, "Defaulters", 8);
    Logger.log("rebuildAllDefaulters_QUERY() issues:\n" + errors.join("\n"));
    return;
  }

  ss.toast('Defaulters rebuilt (shows "NONE" when empty; one TOTALS row only).', "Done", 5);
}