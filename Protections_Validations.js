/***** ===== CLASS LISTS — DROPDOWN VALIDATIONS =====
 * Adds dropdowns for:
 * SEX
 * NATIONALITY
 * DAY/BOARDING
 *
 * RULE:
 * Dropdowns are applied ONLY when Column A of that row is filled.
 * If Column A is blank, dropdowns are removed from that row.
 ************************************************************/

const CLASS_DROPDOWNS = {
  SEX: ["M", "F"],

  DAY_BOARDING: ["DAY", "BOARDING"],

  NATIONALITIES: [
    "AFGHAN",
    "ALBANIAN",
    "ALGERIAN",
    "AMERICAN",
    "ANGOLAN",
    "ARGENTINIAN",
    "AUSTRALIAN",
    "AUSTRIAN",
    "BANGLADESHI",
    "BELGIAN",
    "BRAZILIAN",
    "BRITISH",
    "BURUNDIAN",
    "CAMEROONIAN",
    "CANADIAN",
    "CHINESE",
    "CONGOLESE",
    "DUTCH",
    "EGYPTIAN",
    "ERITREAN",
    "ETHIOPIAN",
    "FRENCH",
    "GERMAN",
    "GHANAIAN",
    "INDIAN",
    "IRANIAN",
    "IRAQI",
    "IRISH",
    "ITALIAN",
    "JAPANESE",
    "KENYAN",
    "KOREAN",
    "LIBERIAN",
    "MEXICAN",
    "NIGERIAN",
    "PAKISTANI",
    "PHILIPPINE",
    "RWANDAN",
    "RWANDESE",
    "SOMALI",
    "SOMALIAN",
    "SOUTH AFRICAN",
    "SOUTH SUDANESE",
    "SUDANESE",
    "SYRIAN",
    "TANZANIAN",
    "TURKISH",
    "UGANDAN",
    "YEMENI",
    "ZAMBIAN",
    "ZIMBABWEAN"
  ],
};

/**
 * Run this manually from Apps Script.
 */
function applyClassDropdownsAll() {
  const sheets = getClassSheets_();
  const errors = [];

  sheets.forEach(sh => {
    try {
      applyClassDropdownsOne_(sh);
    } catch (e) {
      errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  if (errors.length) {
    Logger.log("applyClassDropdownsAll() issues:\n" + errors.join("\n"));
    toast_(`Dropdowns applied with ${errors.length} issue(s). Check logs.`, "Dropdowns", 8);
    return;
  }

  toast_("Dropdowns applied to all class sheets.", "Dropdowns", 6);
}

/**
 * Apply dropdowns to one class sheet.
 * Dropdowns only exist where Column A is filled.
 */
function applyClassDropdownsOne_(sh) {
  if (!sh) return;

  const sheetName = sh.getName();
  if (CLASS_DROPDOWNS_isDefaultersSheet_(sheetName)) return;

  const maxRows = sh.getMaxRows();
  const maxCols = sh.getMaxColumns();

  if (maxRows < 2 || maxCols < 1) return;

  const cSex = CLASS_DROPDOWNS_findHeaderCol_(sh, [
    "SEX",
    CL && CL.H && CL.H.SEX
  ]);

  const cNat = CLASS_DROPDOWNS_findHeaderCol_(sh, [
    "NATIONALITY",
    "NATIONAL",
    "NATION",
    CL && CL.H && CL.H.NAT
  ]);

  const cDB = CLASS_DROPDOWNS_findDayBoardingCol_(sh);

  const targetCols = [cSex, cNat, cDB]
    .filter(c => c && c >= 1 && c <= maxCols);

  if (!targetCols.length) return;

  let totalsRow = null;

  try {
    totalsRow =
      typeof findTotalsRowWide_ === "function"
        ? findTotalsRowWide_(sh)
        : null;
  } catch (_) {
    totalsRow = null;
  }

  let endRow = totalsRow ? totalsRow - 1 : maxRows;

  // IMPORTANT FIX:
  // Prevents "coordinates outside dimensions" on Senior 2, Senior 6, etc.
  endRow = Math.min(endRow, maxRows);

  if (endRow < 2) return;

  const numRows = endRow - 1;

  const colAValues = sh.getRange(2, 1, numRows, 1).getDisplayValues();

  const sexRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.SEX);
  const natRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.NATIONALITIES);
  const dbRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.DAY_BOARDING);

  for (let i = 0; i < colAValues.length; i++) {
    const row = i + 2;
    const colA = String(colAValues[i][0] || "").trim();

    const rowIsActive = !!colA;

    if (rowIsActive) {
      if (cSex && cSex <= maxCols) sh.getRange(row, cSex).setDataValidation(sexRule);
      if (cNat && cNat <= maxCols) sh.getRange(row, cNat).setDataValidation(natRule);
      if (cDB && cDB <= maxCols) sh.getRange(row, cDB).setDataValidation(dbRule);
    } else {
      if (cSex && cSex <= maxCols) sh.getRange(row, cSex).clearDataValidations();
      if (cNat && cNat <= maxCols) sh.getRange(row, cNat).clearDataValidations();
      if (cDB && cDB <= maxCols) sh.getRange(row, cDB).clearDataValidations();
    }
  }

  // Extra safety: clear validations below the real processing area,
  // only if those rows exist.
  const clearStart = endRow + 1;
  if (clearStart <= maxRows) {
    const clearRows = maxRows - clearStart + 1;

    targetCols.forEach(col => {
      try {
        sh.getRange(clearStart, col, clearRows, 1).clearDataValidations();
      } catch (_) {}
    });
  }
}

/**
 * Applies or removes dropdowns for one row.
 * Used by onEdit.
 */
function applyClassDropdownsForRow_(sh, row) {
  if (!sh) return;

  const maxRows = sh.getMaxRows();
  const maxCols = sh.getMaxColumns();

  if (row < 2 || row > maxRows) return;

  const sheetName = sh.getName();
  if (CLASS_DROPDOWNS_isDefaultersSheet_(sheetName)) return;

  let totalsRow = null;

  try {
    totalsRow =
      typeof findTotalsRowWide_ === "function"
        ? findTotalsRowWide_(sh)
        : null;
  } catch (_) {
    totalsRow = null;
  }

  if (totalsRow && row >= totalsRow) return;

  const cSex = CLASS_DROPDOWNS_findHeaderCol_(sh, [
    "SEX",
    CL && CL.H && CL.H.SEX
  ]);

  const cNat = CLASS_DROPDOWNS_findHeaderCol_(sh, [
    "NATIONALITY",
    "NATIONAL",
    "NATION",
    CL && CL.H && CL.H.NAT
  ]);

  const cDB = CLASS_DROPDOWNS_findDayBoardingCol_(sh);

  const colA = String(sh.getRange(row, 1).getDisplayValue() || "").trim();
  const rowIsActive = !!colA;

  const sexRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.SEX);
  const natRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.NATIONALITIES);
  const dbRule = CLASS_DROPDOWNS_buildRule_(CLASS_DROPDOWNS.DAY_BOARDING);

  if (rowIsActive) {
    if (cSex && cSex <= maxCols) sh.getRange(row, cSex).setDataValidation(sexRule);
    if (cNat && cNat <= maxCols) sh.getRange(row, cNat).setDataValidation(natRule);
    if (cDB && cDB <= maxCols) sh.getRange(row, cDB).setDataValidation(dbRule);
  } else {
    if (cSex && cSex <= maxCols) sh.getRange(row, cSex).clearDataValidations();
    if (cNat && cNat <= maxCols) sh.getRange(row, cNat).clearDataValidations();
    if (cDB && cDB <= maxCols) sh.getRange(row, cDB).clearDataValidations();
  }
}

/**
 * Add this inside your existing onEdit(e), if you already have one:
 *
 * handleClassDropdownsOnEdit_(e);
 */
function handleClassDropdownsOnEdit_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh) return;

  const sheetName = sh.getName();
  if (CLASS_DROPDOWNS_isDefaultersSheet_(sheetName)) return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  if (editedRow < 2) return;

  // Only Column A controls whether dropdowns exist.
  if (editedCol !== 1) return;

  const rows = e.range.getNumRows();

  for (let r = editedRow; r < editedRow + rows; r++) {
    applyClassDropdownsForRow_(sh, r);
  }
}

/**
 * Keep this only if you do NOT already have another onEdit(e).
 * If another file already has onEdit(e), delete this one and call:
 *
 * handleClassDropdownsOnEdit_(e);
 *
 * inside the existing onEdit(e).
 */
function onEdit(e) {
  handleClassDropdownsOnEdit_(e);
}

/**
 * Build dropdown validation rule.
 */
function CLASS_DROPDOWNS_buildRule_(values) {
  const cleanedValues = values
    .map(v => String(v || "").trim())
    .filter(Boolean);

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(cleanedValues, true)
    .setAllowInvalid(true)
    .build();
}

/**
 * Optional helper kept for compatibility.
 */
function setDropdownValidation_(range, values) {
  range.setDataValidation(CLASS_DROPDOWNS_buildRule_(values));
}

/**
 * Detects defaulters sheets safely.
 */
function CLASS_DROPDOWNS_isDefaultersSheet_(name) {
  try {
    if (typeof isDefaultersSheetName_ === "function") {
      return isDefaultersSheetName_(name);
    }
  } catch (_) {}

  const n = String(name || "").toLowerCase();
  return n.includes("defaulter");
}

/**
 * Finds normal header columns, even when spacing/case is different.
 */
function CLASS_DROPDOWNS_findHeaderCol_(sh, possibleHeaders) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const maxCols = sh.getMaxColumns();

  if (lastCol < 1 || maxCols < 1) return null;

  const safeLastCol = Math.min(lastCol, maxCols);
  const headers = sh.getRange(1, 1, 1, safeLastCol).getDisplayValues()[0];

  const wanted = possibleHeaders
    .filter(Boolean)
    .map(h => CLASS_DROPDOWNS_cleanHeader_(h));

  for (let i = 0; i < headers.length; i++) {
    const current = CLASS_DROPDOWNS_cleanHeader_(headers[i]);
    if (wanted.includes(current)) return i + 1;
  }

  return null;
}

/**
 * Strong finder for DAY/BOARDING.
 */
function CLASS_DROPDOWNS_findDayBoardingCol_(sh) {
  const maxCols = sh.getMaxColumns();
  const lastCol = Math.min(Math.max(1, sh.getLastColumn()), maxCols);

  if (maxCols < 1) return null;

  try {
    if (CL && CL.H && CL.H.DB) {
      const c = CLASS_DROPDOWNS_findHeaderCol_(sh, [CL.H.DB]);
      if (c && c <= maxCols) return c;
    }
  } catch (_) {}

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  for (let i = 0; i < headers.length; i++) {
    const raw = String(headers[i] || "");
    const cleaned = CLASS_DROPDOWNS_cleanHeader_(raw);

    if (
      cleaned === "DAYBOARDING" ||
      cleaned.includes("DAYBOARDING") ||
      (cleaned.includes("DAY") && cleaned.includes("BOARD"))
    ) {
      return i + 1;
    }
  }

  // Layout fallback: F is DAY/BOARDING in your sheets.
  if (maxCols >= 6) {
    const fHeader = String(sh.getRange(1, 6).getDisplayValue() || "");
    const fClean = CLASS_DROPDOWNS_cleanHeader_(fHeader);

    if (fClean.includes("DAY") || fClean.includes("BOARD")) {
      return 6;
    }
  }

  return null;
}

/**
 * Normalizes headers.
 */
function CLASS_DROPDOWNS_cleanHeader_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/\n/g, " ")
    .replace(/\r/g, " ")
    .replace(/\//g, " ")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, "")
    .trim();
}