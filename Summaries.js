/***** ===== CLASS LISTS — TERM AUTOMATION (SUMMARIES) =====
 * SAFE SUMMARIES BUILDER
 * Correct totals cells + enrolment + grand totals
 ************************************************************/

/** Where the summaries output will be written */
const SUMMARY_LAYOUT = {
  MONEY_START_ROW: 1,
  MONEY_START_COL: 1,

  DEBTORS_ROW: 10,
  HEADCOUNT_START_ROW: 13,

  CLEAR_WIDTH: 6,   // A..F
  CLEAR_ROWS: 140,
};

const HEADCOUNT_GROUP_STYLE = {
  bg: "#E6E6E6",
  fontColor: "#000000",
  bold: true,
};

const YELLOW_TOTAL_STYLE = {
  bg: "#FFF2CC",
  fontColor: "#000000",
  bold: true,
};

/**
 * CONFIRMED TOTALS ROWS FROM YOUR SCREENSHOTS
 *
 * G = LAST TERM BALANCE / LTB
 * H = TERM BALANCE
 * I = PAID
 * J = BALANCE
 */
const TOTALS_CELLS = {
  // Nursery
  "Baby class list":   { row: 32 },
  "Middle class list": { row: 38 },
  "Top Class list":    { row: 34 },

  // Primary
  "Primary 1 list": { row: 37 },
  "Primary 2 list": { row: 44 },
  "Primary 3 list": { row: 48 },
  "Primary 4 list": { row: 43 },
  "Primary 5 list": { row: 44 },
  "Primary 6 list": { row: 42 },
  "Primary 7 list": { row: 59 },

  // Secondary
  "Senior 1 list": { row: 65 },
  "Senior 2 list": { row: 65 },
  "Senior 3 list": { row: 81 },
  "Senior 4 list": { row: 64 },
  "Senior 5 list": { row: 51 },
  "Senior 6 list": { row: 67 },
};

const SUMMARY_SECTIONS = [
  {
    label: "TOTAL NURSERY",
    sheets: ["Baby class list", "Middle class list", "Top Class list"],
  },
  {
    label: "TOTAL PRIMARY",
    sheets: [
      "Primary 1 list",
      "Primary 2 list",
      "Primary 3 list",
      "Primary 4 list",
      "Primary 5 list",
      "Primary 6 list",
      "Primary 7 list",
    ],
  },
  {
    label: "TOTAL SECONDARY",
    sheets: [
      "Senior 1 list",
      "Senior 2 list",
      "Senior 3 list",
      "Senior 4 list",
      "Senior 5 list",
      "Senior 6 list",
    ],
  },
];

function refreshSummaries() {
  const ss = ss_();
  const S =
    ss.getSheetByName(CL.SHEETNAMES.SUMMARIES) ||
    ss.insertSheet(CL.SHEETNAMES.SUMMARIES);

  clearSummariesBlock_(S);

  writeSummariesMoneyTotals_(S);
  writeDebtorsOutstanding_(S);
  writeSummariesHeadcounts_(S);

  toast_("SUMMARIES refreshed.", "SUMMARIES", 6);
}

function refreshSummariesTotalsLTB() {
  refreshSummaries();
}

function clearSummariesBlock_(S) {
  S.getRange(
    SUMMARY_LAYOUT.MONEY_START_ROW,
    SUMMARY_LAYOUT.MONEY_START_COL,
    SUMMARY_LAYOUT.CLEAR_ROWS,
    SUMMARY_LAYOUT.CLEAR_WIDTH
  )
    .clearContent()
    .clearFormat();
}

function _sumTotalsFormulaStrict_(sheetNames, colLetter) {
  const refs = sheetNames
    .map(name => {
      const info = TOTALS_CELLS[name];
      if (!info || !info.row) return null;

      const safeName = String(name).replace(/'/g, "''");
      return `'${safeName}'!${colLetter}${info.row}`;
    })
    .filter(Boolean);

  return refs.length ? `=SUM(${refs.join(",")})` : "=0";
}

function _sumAllTotalsFormulaStrict_(colLetter) {
  const refs = Object.keys(TOTALS_CELLS).map(name => {
    const info = TOTALS_CELLS[name];
    const safeName = String(name).replace(/'/g, "''");
    return `'${safeName}'!${colLetter}${info.row}`;
  });

  return refs.length ? `=SUM(${refs.join(",")})` : "=0";
}

/** TOP MONEY TABLE */
function writeSummariesMoneyTotals_(S) {
  const r0 = SUMMARY_LAYOUT.MONEY_START_ROW;
  const c0 = SUMMARY_LAYOUT.MONEY_START_COL;

  S.getRange(r0, c0, 1, 5)
    .setValues([["", "TERM BALANCE", "PAID", "BALANCE", "LTB"]])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  S.getRange(r0, c0 + 1).setBackground("#00FF00");
  S.getRange(r0, c0 + 2).setBackground("#00B0F0");
  S.getRange(r0, c0 + 3).setBackground("#FF0000").setFontColor("#FFFFFF");
  S.getRange(r0, c0 + 4).setBackground("#FFD966");

  try {
    S.setRowHeight(r0, 28);
  } catch (_) {}

  SUMMARY_SECTIONS.forEach((sec, i) => {
    const row = r0 + 1 + i;

    S.getRange(row, c0)
      .setValue(sec.label)
      .setFontWeight("bold");

    S.getRange(row, c0 + 1).setFormula(_sumTotalsFormulaStrict_(sec.sheets, "H"));
    S.getRange(row, c0 + 2).setFormula(_sumTotalsFormulaStrict_(sec.sheets, "I"));
    S.getRange(row, c0 + 3).setFormula(_sumTotalsFormulaStrict_(sec.sheets, "J"));
    S.getRange(row, c0 + 4).setFormula(_sumTotalsFormulaStrict_(sec.sheets, "G"));
  });

  const grandRow = r0 + 1 + SUMMARY_SECTIONS.length;

  S.getRange(grandRow, c0)
    .setValue("GRAND TOTAL")
    .setFontWeight("bold");

  S.getRange(grandRow, c0 + 1).setFormula(_sumAllTotalsFormulaStrict_("H"));
  S.getRange(grandRow, c0 + 2).setFormula(_sumAllTotalsFormulaStrict_("I"));
  S.getRange(grandRow, c0 + 3).setFormula(_sumAllTotalsFormulaStrict_("J"));
  S.getRange(grandRow, c0 + 4).setFormula(_sumAllTotalsFormulaStrict_("G"));

  S.getRange(grandRow, c0, 1, 5)
    .setBackground(YELLOW_TOTAL_STYLE.bg)
    .setFontColor(YELLOW_TOTAL_STYLE.fontColor)
    .setFontWeight("bold");

  const enrolRow = grandRow + 1;

  S.getRange(enrolRow, c0)
    .setValue("TOTAL ENROLMENT")
    .setFontWeight("bold");

  S.getRange(enrolRow, c0 + 1).clearContent();
  S.getRange(enrolRow, c0 + 2).clearContent();
  S.getRange(enrolRow, c0 + 3).setFormula("=C35");
  S.getRange(enrolRow, c0 + 4).clearContent();

  S.getRange(enrolRow, c0, 1, 5)
    .setBackground(YELLOW_TOTAL_STYLE.bg)
    .setFontWeight("bold");

  S.getRange(r0 + 1, c0 + 1, SUMMARY_SECTIONS.length + 2, 4)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setHorizontalAlignment("right");

  S.getRange(r0, c0, enrolRow - r0 + 1, 5)
    .setBorder(true, true, true, true, true, true, "#555555", SpreadsheetApp.BorderStyle.SOLID);

  try {
    S.setColumnWidth(c0, 180);
    S.setColumnWidth(c0 + 1, 140);
    S.setColumnWidth(c0 + 2, 120);
    S.setColumnWidth(c0 + 3, 120);
    S.setColumnWidth(c0 + 4, 120);
  } catch (_) {}
}

function writeDebtorsOutstanding_(S) {
  try {
    ensureLeaversSheet_();
  } catch (_) {}

  const ss = ss_();
  const leaversName = CL.LEAVERS.SHEET || "Leavers – Debtors";
  const L = ss.getSheetByName(leaversName);
  const r = SUMMARY_LAYOUT.DEBTORS_ROW;

  S.getRange(r, 1)
    .setValue("LEAVERS – DEBTORS OUTSTANDING (UNSETTLED)")
    .setFontWeight("bold");

  if (!L) {
    S.getRange(r, 4).setFormula("=0").setFontWeight("bold");
    return;
  }

  const headers = L.getRange(1, 1, 1, L.getLastColumn())
    .getValues()[0]
    .map(h => String(h || "").trim());

  const iStatus = headers.findIndex(h => h.toLowerCase() === "status");
  const iBal = headers.findIndex(
    h => h.toLowerCase() === String(CL.H.BAL).toLowerCase()
  );

  if (iStatus === -1 || iBal === -1) {
    S.getRange(r, 4).setFormula("=0").setFontWeight("bold");
    return;
  }

  const colStatus = colLetter_(iStatus + 1);
  const colBal = colLetter_(iBal + 1);
  const safeName = leaversName.replace(/'/g, "''");

  const formula =
    `=IFERROR(SUM(FILTER('${safeName}'!${colBal}2:${colBal},` +
    `UPPER('${safeName}'!${colStatus}2:${colStatus})<>"SETTLED",` +
    `UPPER('${safeName}'!${colStatus}2:${colStatus})<>"REINSTATED")),0)`;

  S.getRange(r, 4)
    .setFormula(formula)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setHorizontalAlignment("right");

  S.getRange(r, 1, 1, 5)
    .setBackground("#FFF2CC")
    .setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);
}

function _getTotalsRowForHeadcount_(sh, sheetName) {
  const fixed = TOTALS_CELLS[sheetName] && TOTALS_CELLS[sheetName].row;
  if (fixed) return fixed;

  try {
    return findTotalsRowWide_(sh) || null;
  } catch (_) {
    return null;
  }
}

function _countStudentsFormulaForSheet_(sheetName, sh) {
  if (!sh) return "=0";

  const tRow = _getTotalsRowForHeadcount_(sh, sheetName);
  if (!tRow) return "=0";

  const endRow = tRow - 1;
  if (endRow < 2) return "=0";

  let cName = 3;

  try {
    const map = getHeaderIndexMap_(sh);
    cName = map[CL.H.NAME] || 3;
  } catch (_) {}

  const L = colLetter_(cName);
  const safeName = String(sheetName).replace(/'/g, "''");

  return `=COUNTIF('${safeName}'!${L}2:${L}${endRow},"<>")`;
}

function writeSummariesHeadcounts_(S) {
  const ss = ss_();

  const COUNTS = [
    {
      group: "TOTAL NURSERY",
      items: [
        ["BABY CLASS", "Baby class list"],
        ["MIDDLE CLASS", "Middle class list"],
        ["TOP CLASS", "Top Class list"],
      ],
    },
    {
      group: "TOTAL PRIMARY",
      items: [
        ["PRIMARY 1", "Primary 1 list"],
        ["PRIMARY 2", "Primary 2 list"],
        ["PRIMARY 3", "Primary 3 list"],
        ["PRIMARY 4", "Primary 4 list"],
        ["PRIMARY 5", "Primary 5 list"],
        ["PRIMARY 6", "Primary 6 list"],
        ["PRIMARY 7", "Primary 7 list"],
      ],
    },
    {
      group: "TOTAL SECONDARY",
      items: [
        ["SENIOR 1", "Senior 1 list"],
        ["SENIOR 2", "Senior 2 list"],
        ["SENIOR 3", "Senior 3 list"],
        ["SENIOR 4", "Senior 4 list"],
        ["SENIOR 5", "Senior 5 list"],
        ["SENIOR 6", "Senior 6 list"],
      ],
    },
  ];

  let row = SUMMARY_LAYOUT.HEADCOUNT_START_ROW;
  const groupTotalRows = [];

  COUNTS.forEach(g => {
    const groupRow = row;
    groupTotalRows.push(groupRow);

    S.getRange(groupRow, 1)
      .setValue(g.group)
      .setFontWeight("bold");

    S.getRange(groupRow, 1, 1, 3)
      .setBackground(HEADCOUNT_GROUP_STYLE.bg)
      .setFontColor(HEADCOUNT_GROUP_STYLE.fontColor)
      .setFontWeight(HEADCOUNT_GROUP_STYLE.bold ? "bold" : "normal");

    row++;

    const firstItemRow = row;

    g.items.forEach(([label, tab]) => {
      S.getRange(row, 2).setValue(label);

      const sh = ss.getSheetByName(tab);
      S.getRange(row, 3).setFormula(_countStudentsFormulaForSheet_(tab, sh));

      row++;
    });

    const lastItemRow = row - 1;

    S.getRange(groupRow, 3)
      .setFormula(
        lastItemRow >= firstItemRow
          ? `=SUM(C${firstItemRow}:C${lastItemRow})`
          : "=0"
      )
      .setFontWeight("bold")
      .setHorizontalAlignment("right");

    row++;
  });

  const overallRow = row;
  const refs = groupTotalRows.map(r => `C${r}`).join(",");

  S.getRange(overallRow, 1)
    .setValue("OVERALL TOTAL ENROLMENT")
    .setFontWeight("bold");

  S.getRange(overallRow, 3)
    .setFormula(`=SUM(${refs})`)
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  S.getRange(overallRow, 1, 1, 3)
    .setBackground(YELLOW_TOTAL_STYLE.bg)
    .setFontWeight("bold");

  const start = SUMMARY_LAYOUT.HEADCOUNT_START_ROW;
  const end = overallRow;

  S.getRange(start, 1, end - start + 1, 3)
    .setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);

  S.getRange(start, 3, end - start + 1, 1)
    .setHorizontalAlignment("right")
    .setNumberFormat("#,##0");
}