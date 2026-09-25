/***** ===== CLASS LISTS — TERM AUTOMATION (CONFIG) =====
 * Configuration/constants only (no logic).
 *
 * GOALS (this config supports them):
 * ✅ One source of truth for sheet names, headers, class order, fees, totals layout.
 * ✅ Prevent “silent fallbacks” by making names explicit + providing allowed aliases.
 * ✅ Handle the “Leavers – Debtors” dash issue by allowing safe name variants.
 * ✅ Keep totals columns standardized (G:J) while still relying on headers in scripts.
 ********************************************************/

const CL = {
  VERSION: "v1.5-config-hardened-names-aliases",

  /* =========================
   * Sheet naming rules
   * ========================= */
  CLASS_SUFFIX: " list",
  DEFS_SUFFIX: " Defaulters",

  /**
   * Some sheet names may exist with small variations (dash types, spacing, case).
   * Put the canonical name first; scripts should prefer canonical, then try aliases.
   * (No logic here—just the list of acceptable names.)
   */
  SHEET_ALIASES: {
    // IMPORTANT: “–” (en dash) vs “-” (hyphen)
    LEAVERS: ["Leavers – Debtors", "Leavers - Debtors", "Leavers–Debtors", "Leavers—Debtors"],
    GRADUATES: ["Graduates", "GRADUATES"],
    SUMMARIES: ["SUMMARIES", "Summaries", "SUMMARY", "Summary"],
    TERM_LOG: ["_TermLog_CLASSLISTS", "_TermLog", "Term Log", "TERM LOG"],
  },

  /* =========================
   * Term / Archive settings
   * ========================= */
  TERM: {
    LOG_SHEET: "_TermLog_CLASSLISTS",
    ARCHIVE_FOLDER: "Class Lists Archives",
  },

  ARCHIVE: {
    PROTECT_SHEETS: true,
    EDITORS_POLICY: "FILE_EDITORS", // anyone with edit access to the file
  },

  /* =========================
   * Canonical headers
   * MUST match sheet headers exactly (row 1)
   * ========================= */
  H: {
    NO: "NO.",
    DATE: "Date of Admission",
    NAME: "Student Name",
    SEX: "SEX",
    NAT: "NATIONALITY",
    DB: "DAY/BOARDING",

    LTB: "LAST TERM BALANCE",
    FEE: "FULL FEE",
    PAID: "P (Paid)",
    BAL: "BALANCE",

    CAUT: "CAUTION",
    TEL: "TEL",
    REM: "REMARKS",
  },

  /**
   * Optional accepted header variants (for safety).
   * Scripts MAY use these to find columns if some sheets have slightly different text.
   * (Still config-only.)
   */
  H_ALIASES: {
    // Keep canonical first where possible
    PAID: ["P (Paid)", "P(Paid)", "PAID", "Paid", "P (PAID)"],
    LTB: ["LAST TERM BALANCE", "LTB", "Last Term Balance"],
    FEE: ["FULL FEE", "FEE", "Full Fee"],
    BAL: ["BALANCE", "Balance"],
    DATE: ["Date of Admission", "Admission Date", "Date"],
    NAME: ["Student Name", "Name"],
    DB: ["DAY/BOARDING", "DAY / BOARDING", "Boarding", "Day/Boarding"],
    NAT: ["NATIONALITY", "Nationality"],
    SEX: ["SEX", "Sex", "Gender"],
    NO: ["NO.", "NO", "No."],
    CAUT: ["CAUTION", "Caution"],
    TEL: ["TEL", "Telephone", "Phone"],
    REM: ["REMARKS", "Remarks", "Comment", "Comments"],
  },

  /* =========================
   * Standard FULL FEE defaults
   * Used when FULL FEE is blank/0 for student rows.
   * ========================= */
  FEES: {
    "Baby class list": 500000,
    "Middle class list": 500000,
    "Top Class list": 500000,

    "Primary 1 list": 650000,
    "Primary 2 list": 650000,
    "Primary 3 list": 650000,
    "Primary 4 list": 650000,
    "Primary 5 list": 650000,
    "Primary 6 list": 650000,
    "Primary 7 list": 650000,

    "Senior 1 list": 750000,
    "Senior 2 list": 750000,
    "Senior 3 list": 750000,
    "Senior 4 list": 750000,
    "Senior 5 list": 750000,
    "Senior 6 list": 750000,
  },

  /* =========================
   * Totals columns standardization
   * These are your intended letters on class lists:
   *   LTB=G, FEE=H, PAID=I, BAL=J
   * (Scripts should still prefer header lookup; these are for layout documentation
   *  and fixed totals positioning where needed.)
   * ========================= */
  TOTALS_COLS: {
    LTB: "G",
    FEE: "H",
    PAID: "I",
    BAL: "J",
  },

  /**
   * Optional: numeric versions (useful if any module wants indices).
   * G=7, H=8, I=9, J=10
   */
  TOTALS_COLS_INDEX: {
    LTB: 7,
    FEE: 8,
    PAID: 9,
    BAL: 10,
  },

  /* =========================
   * General limits
   * ========================= */
  ROWS: 800,

  /* =========================
   * Admin + protections config
   * ========================= */
  ADMINS: ["vishvamdatta@gmail.com", "ddaannson@gmail.com"],

  // Sheets to lock/hide for non-admins (protection module will use this)
  SHEETS_LOCKED: ["SUMMARIES"],

  /* =========================
   * Fixed totals positions (optional layout anchor)
   * If you stop deleting rows during promotion, these remain stable.
   * ========================= */
  FIXED_TOTALS: {
    // Nursery
    "Baby class list":   { row: 22, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Middle class list": { row: 28, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Top Class list":    { row: 31, ltb: "G", fee: "H", paid: "I", bal: "J" },

    // Primary
    "Primary 1 list": { row: 31, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 2 list": { row: 31, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 3 list": { row: 25, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 4 list": { row: 31, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 5 list": { row: 20, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 6 list": { row: 33, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Primary 7 list": { row: 34, ltb: "G", fee: "H", paid: "I", bal: "J" },

    // Secondary
    "Senior 1 list": { row: 74, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Senior 2 list": { row: 59, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Senior 3 list": { row: 78, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Senior 4 list": { row: 64, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Senior 5 list": { row: 60, ltb: "G", fee: "H", paid: "I", bal: "J" },
    "Senior 6 list": { row: 54, ltb: "G", fee: "H", paid: "I", bal: "J" },
  },

  /* =========================
   * Promotion order (Start New Year)
   * ========================= */
  CLASS_ORDER: [
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
  ],

  /* =========================
   * Leavers / Debtors system
   * ========================= */
  LEAVERS: {
    // Canonical name
    SHEET: "Leavers – Debtors",

    // Required A–L match class list columns, then system fields
    HEAD: [
      "NO.",
      "Date of Admission",
      "Student Name",
      "SEX",
      "NATIONALITY",
      "DAY/BOARDING",
      "LAST TERM BALANCE",
      "FULL FEE",
      "P (Paid)",
      "BALANCE",
      "CAUTION",
      "TEL",

      "Former Class",
      "Date Left",
      "Exit Reason",
      "Status",          // Unsettled / Settled
      "Reinstated To",
      "Reinstated At",
      "Source Sheet",
      "Source Row",
      "Created At",
      "Updated At",
      "Student ID",
    ],
  },

  /* =========================
   * Other system sheet names
   * ========================= */
  SHEETNAMES: {
    SUMMARIES: "SUMMARIES",
    GRADUATES: "Graduates",
  },

  /**
   * Centralized system sheet list (helps other modules exclude them deterministically).
   * (Still config-only.)
   */
  SYSTEM_SHEETS: [
    "_TermLog_CLASSLISTS",
    "SUMMARIES",
    "Graduates",
    "Leavers – Debtors",
  ],
};
