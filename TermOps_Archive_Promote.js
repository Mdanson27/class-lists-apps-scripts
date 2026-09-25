/***** ===== CLASS LISTS — TERM OPS (5 STEPS) =====
 * v2.0-termops-steps-runnable
 *
 * MAIN UPDATE (YOUR REQUEST):
 * ✅ You can run STEP 1, 2, 3, 4, 5 individually from the Apps Script editor.
 *
 * WORKFLOW:
 * 1) Run ONE init function:
 *    - termOps_InitNewTerm()   OR
 *    - termOps_InitNewYear()
 *
 * 2) Then run steps individually:
 *    - termOps_Step1()
 *    - termOps_Step2()
 *    - termOps_Step3()
 *    - termOps_Step4()
 *    - termOps_Step5()
 *
 * Optional:
 * - termOps_RunAllRemaining()  (runs from current step → end)
 * - termOps_ResetProgress()
 *
 * IMPORTANT NO-GREEN RULE:
 * ✅ Class sheets MUST NOT have conditional formatting.
 *   We enforce this by calling fixTotalsFormattingAll_() in Step 5
 *   AND ensureStructureOne_ already clears conditional formatting on class sheets (in your updated Structure file).
 *
 * DEPENDS ON:
 * - Config.gs (CL)
 * - Utils.gs: ss_(), toast_(), withDocLock_(), getClassSheets_(), getStudentRegion_(), getHeaderIndexMap_(), _flatUpper_()
 * - ClassLists_Structure.gs: ensureStructureAll(), ensureStructureOne_()
 * - Totals.gs: fixTotalsAllSheets(), fixTotalsFormattingAll_()
 * - Defaulters.gs: rebuildAllDefaulters_QUERY()
 ************************************************************/

/* =========================
   Local wrappers (no collisions)
========================= */

function TOPS_ss_() {
  if (typeof ss_ === "function") return ss_();
  return SpreadsheetApp.getActiveSpreadsheet();
}

function TOPS_normText_(v) {
  if (typeof _normText_ === "function") return _normText_(v);
  if (v === null || v === undefined) return "";
  return String(v).trim().replace(/\s+/g, " ");
}

function TOPS_flatUpper_(v) {
  if (typeof _flatUpper_ === "function") return _flatUpper_(v);
  return TOPS_normText_(v).toUpperCase().replace(/\s+/g, "");
}

function TOPS_toast_(msg, title, seconds) {
  if (typeof toast_ === "function") return toast_(msg, title, seconds);
  try { TOPS_ss_().toast(String(msg || ""), String(title || "Info"), Number(seconds || 5)); } catch (_) {}
}

/** Lock helper (uses your withDocLock_ if available) */
function TOPS_withDocLock_(fn) {
  if (typeof withDocLock_ === "function") return withDocLock_(fn);

  const lock = LockService.getDocumentLock();
  const start = Date.now();
  const maxMs = 30 * 1000;
  const stepMs = 600;

  while (true) {
    try { if (lock.tryLock(3000)) break; } catch (_) {}
    if (Date.now() - start > maxMs) throw new Error("Lock timeout. End other runs and retry.");
    Utilities.sleep(stepMs);
  }

  try { return fn(); }
  finally { try { lock.releaseLock(); } catch (_) {} }
}

function TOPS_fail_(ui, context, err) {
  const msg = (err && err.message) ? err.message : String(err);
  ui.alert(`${context} FAILED.\n\n${msg}\n\nApps Script → Executions shows exact line.`);
  throw err;
}

function TOPS_assertConfig_() {
  if (!CL) throw new Error("Missing CL (Config.gs).");
  if (!CL.H) throw new Error("Missing CL.H (Config.gs).");
  if (!CL.TERM || !CL.TERM.LOG_SHEET || !CL.TERM.ARCHIVE_FOLDER) {
    throw new Error("Missing CL.TERM.LOG_SHEET or CL.TERM.ARCHIVE_FOLDER (Config.gs).");
  }
  if (!Array.isArray(CL.CLASS_ORDER) || !CL.CLASS_ORDER.length) {
    throw new Error("Missing/empty CL.CLASS_ORDER (Config.gs).");
  }
  if (!CL.FEES) TOPS_toast_("Warning: CL.FEES not found. Fee fallback will be 0.", "Config", 8);
}

/* =========================
   Progress state (resumable)
========================= */

const TOPS_STATE_KEY_ = "TERMOPS_STATE_5STEP";

function TOPS_stateGet_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(TOPS_STATE_KEY_);
    return raw ? JSON.parse(raw) : null;
  } catch (_) { return null; }
}
function TOPS_stateSet_(obj) {
  PropertiesService.getScriptProperties().setProperty(TOPS_STATE_KEY_, JSON.stringify(obj || {}));
}
function TOPS_stateClear_() {
  PropertiesService.getScriptProperties().deleteProperty(TOPS_STATE_KEY_);
}

function termOps_ResetProgress() {
  TOPS_stateClear_();
  TOPS_toast_("Cleared saved progress.", "Reset", 6);
}

/* =========================
   Sheet resolution helpers
========================= */

function TOPS_getSheetByAliases_(canonicalName, aliasKey) {
  const ss = TOPS_ss_();
  if (canonicalName) {
    const sh = ss.getSheetByName(String(canonicalName));
    if (sh) return sh;
  }
  try {
    const aliases = CL?.SHEET_ALIASES?.[aliasKey];
    if (Array.isArray(aliases)) {
      for (const n of aliases) {
        const sh = ss.getSheetByName(String(n));
        if (sh) return sh;
      }
    }
  } catch (_) {}
  return null;
}

function TOPS_isClassSheet_(sh) {
  if (!sh) return false;
  const suf = String(CL?.CLASS_SUFFIX || " list").toLowerCase();
  return String(sh.getName()).toLowerCase().endsWith(suf);
}

/** Always prefer CL.CLASS_ORDER (deterministic; your updated Utils does this too) */
function TOPS_getClassSheets_() {
  const ss = TOPS_ss_();
  const out = [];
  (CL.CLASS_ORDER || []).forEach(name => {
    const sh = ss.getSheetByName(String(name));
    if (sh) out.push(sh);
  });
  return out;
}

/* =========================
   Header/totals/region fallbacks
========================= */

function TOPS_getHeaderIndexMap_(sh) {
  if (typeof getHeaderIndexMap_ === "function") return getHeaderIndexMap_(sh);
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(TOPS_normText_);
  const map = {};
  for (let i = 0; i < headers.length; i++) if (headers[i] && map[headers[i]] == null) map[headers[i]] = i + 1;
  return map;
}

function TOPS_headerMapUpper_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(TOPS_normText_);
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const k = TOPS_flatUpper_(headers[i]);
    if (k && map[k] == null) map[k] = i + 1;
  }
  return map;
}

function TOPS_getStudentRegion_(sh) {
  if (typeof getStudentRegion_ === "function") return getStudentRegion_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { startRow: 2, endRow: 1, totalsRow: null };
  const totalsRow = (typeof findTotalsRowWide_ === "function") ? findTotalsRowWide_(sh) : null;
  const endRow = totalsRow ? (totalsRow - 1) : lastRow;
  return { startRow: 2, endRow, totalsRow };
}

function TOPS_nextRowAfterLastNumberedA_(sh) {
  if (typeof nextRowAfterLastNumberedA_ === "function") return nextRowAfterLastNumberedA_(sh);

  const region = TOPS_getStudentRegion_(sh);
  const endRow = region.endRow;
  if (!endRow || endRow < 2) return 2;

  const colA = sh.getRange(2, 1, endRow - 1, 1).getValues().flat();
  let lastIdx = -1;
  for (let i = 0; i < colA.length; i++) {
    const n = Number(colA[i]);
    if (Number.isFinite(n) && n > 0) lastIdx = i;
  }
  return Math.max(2, 2 + (lastIdx + 1));
}

function TOPS_ensureRows_(sh, rowNum) {
  if (typeof ensureRows_ === "function") return ensureRows_(sh, rowNum);
  const need = rowNum - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
}

/* =========================
   FULL FEE rules (MODE + fallback)
========================= */

function TOPS_feeMode_(sh) {
  const exact = TOPS_getHeaderIndexMap_(sh);
  const upper = TOPS_headerMapUpper_(sh);
  const cFEE = exact[CL.H.FEE] || upper[TOPS_flatUpper_(CL.H.FEE)];
  if (!cFEE) return 0;

  const region = TOPS_getStudentRegion_(sh);
  if (!region || region.endRow < 2) return 0;

  const vals = sh.getRange(2, cFEE, region.endRow - 1, 1).getValues().flat();
  const freq = new Map();

  for (const v of vals) {
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0) continue;
    freq.set(n, (freq.get(n) || 0) + 1);
  }

  let best = 0, bestCount = -1;
  for (const [fee, count] of freq.entries()) {
    if (count > bestCount || (count === bestCount && fee > best)) { best = fee; bestCount = count; }
  }
  return best || 0;
}

function TOPS_feeFallback_(sheetName) {
  const n = String(sheetName || "");
  const f = (CL.FEES && CL.FEES[n] != null) ? Number(CL.FEES[n]) : 0;
  return Number.isFinite(f) ? f : 0;
}

function TOPS_getSheetFee_(sh) {
  const mode = TOPS_feeMode_(sh);
  if (mode > 0) return mode;
  return TOPS_feeFallback_(sh.getName());
}

/** STEP 2 helper: Set FULL FEE only where blank/0 (students only) */
function TOPS_applyFeeDefaultsOne_(sh) {
  const exact = TOPS_getHeaderIndexMap_(sh);
  const upper = TOPS_headerMapUpper_(sh);

  const cNAME = exact[CL.H.NAME] || upper[TOPS_flatUpper_(CL.H.NAME)];
  const cFEE  = exact[CL.H.FEE]  || upper[TOPS_flatUpper_(CL.H.FEE)];
  if (!cNAME || !cFEE) return 0;

  const region = TOPS_getStudentRegion_(sh);
  if (!region || region.endRow < 2) return 0;

  const fee = TOPS_getSheetFee_(sh);
  if (!fee) return 0;

  const nRows = region.endRow - 1;

  const names = sh.getRange(2, cNAME, nRows, 1).getValues();
  const fees  = sh.getRange(2, cFEE,  nRows, 1).getValues();

  let changed = 0;
  const out = fees.map((row, i) => {
    const nm = TOPS_flatUpper_(names[i][0]);
    if (!nm || nm === "TOTALS") return [row[0]];

    const v = Number(row[0]) || 0;
    if (v > 0) return [row[0]];
    changed++;
    return [fee];
  });

  if (changed) sh.getRange(2, cFEE, nRows, 1).setValues(out);
  return changed;
}

/* =========================
   APPEND ABOVE TOTALS (VALUES ONLY, NO FORMAT COPY)
   (format copying can re-introduce green; we avoid it)
========================= */

function TOPS_appendRowsAboveTotals_(to, rows) {
  if (!rows || !rows.length) return;

  const lastCol = Math.max(1, to.getLastColumn());
  const norm = rows.map(r => {
    const row = Array.isArray(r) ? r.slice(0, lastCol) : [];
    while (row.length < lastCol) row.push("");
    return row;
  });

  const region = TOPS_getStudentRegion_(to);
  const tRow = region && region.totalsRow ? region.totalsRow : null;

  const writeBlock = (startRow, block) => {
    TOPS_ensureRows_(to, startRow + block.length - 1);
    const rng = to.getRange(startRow, 1, block.length, lastCol);
    rng.setValues(block);

    // keep inserted/filled rows neutral (no green)
    try { rng.clearFormat(); } catch (_) {}
  };

  if (!tRow) {
    const start = Math.max(2, to.getLastRow() + 1);
    writeBlock(start, norm);
    return;
  }

  const startCandidate = TOPS_nextRowAfterLastNumberedA_(to);
  const freeSlots = Math.max(0, tRow - startCandidate);

  if (freeSlots >= norm.length) {
    writeBlock(startCandidate, norm);
    return;
  }

  if (freeSlots > 0) {
    writeBlock(startCandidate, norm.slice(0, freeSlots));
  }

  const remaining = norm.length - freeSlots;
  if (remaining > 0) {
    to.insertRowsBefore(tRow, remaining);
    writeBlock(tRow, norm.slice(freeSlots));
  }
}

/* =========================
   MAP + PROMOTION helpers
========================= */

function TOPS_mapRowsByHeaders_(fromSheet, toSheet, fromRows) {
  const fromMap = TOPS_headerMapUpper_(fromSheet);
  const toMap = TOPS_headerMapUpper_(toSheet);
  const toLastCol = Math.max(1, toSheet.getLastColumn());

  const keys = Object.keys(toMap);
  return (fromRows || []).map(r => {
    const out = new Array(toLastCol).fill("");
    keys.forEach(k => {
      const cTo = toMap[k];
      const cFrom = fromMap[k];
      if (cTo && cFrom) out[cTo - 1] = r[cFrom - 1];
    });
    return out;
  });
}

function TOPS_forceDestFee_(toSheet, mappedRows) {
  if (!mappedRows || !mappedRows.length) return mappedRows;

  const fee = TOPS_getSheetFee_(toSheet);
  if (!fee) return mappedRows;

  const exact = TOPS_getHeaderIndexMap_(toSheet);
  const upper = TOPS_headerMapUpper_(toSheet);
  const cFEE = exact[CL.H.FEE] || upper[TOPS_flatUpper_(CL.H.FEE)];
  if (!cFEE) return mappedRows;

  mappedRows.forEach(r => { if (Array.isArray(r) && r.length >= cFEE) r[cFEE - 1] = fee; });
  return mappedRows;
}

/* =========================
   LEAVERS + GRADUATES (ensure)
========================= */

function TOPS_ensureLeaversSheet_() {
  if (!CL?.LEAVERS?.SHEET || !Array.isArray(CL.LEAVERS.HEAD) || !CL.LEAVERS.HEAD.length) {
    throw new Error("Leavers failed: CL.LEAVERS.SHEET/HEAD missing.");
  }
  const ss = TOPS_ss_();

  // resolve via aliases if needed
  let sh = TOPS_getSheetByAliases_(CL.LEAVERS.SHEET, "LEAVERS");
  if (!sh) sh = ss.insertSheet(CL.LEAVERS.SHEET);

  const desired = CL.LEAVERS.HEAD;
  if (sh.getMaxColumns() < desired.length) sh.insertColumnsAfter(sh.getMaxColumns(), desired.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, desired.length).setValues([desired]);
  sh.setFrozenRows(1);
  return sh;
}

function TOPS_graduatesHeaders_() {
  return [
    CL.H.NO, CL.H.DATE, CL.H.NAME, CL.H.SEX, CL.H.NAT, CL.H.DB,
    CL.H.LTB, CL.H.FEE, CL.H.PAID, CL.H.BAL, CL.H.CAUT, CL.H.TEL, CL.H.REM,
    "Graduation Year",
  ];
}

function TOPS_ensureGraduatesSheet_() {
  const ss = TOPS_ss_();
  const canonical = (CL.SHEETNAMES && CL.SHEETNAMES.GRADUATES) ? CL.SHEETNAMES.GRADUATES : "Graduates";

  let sh = TOPS_getSheetByAliases_(canonical, "GRADUATES");
  if (!sh) sh = ss.insertSheet(canonical);

  const desired = TOPS_graduatesHeaders_();
  if (sh.getMaxColumns() < desired.length) sh.insertColumnsAfter(sh.getMaxColumns(), desired.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, desired.length).setValues([desired]);
  sh.setFrozenRows(1);
  return sh;
}

/* =========================
   ARCHIVE + TERM LOG
========================= */

function TOPS_getCurrentFileEditorEmails_() {
  const ss = TOPS_ss_();
  const file = DriveApp.getFileById(ss.getId());
  const editors = file.getEditors() || [];
  return editors.map(u => (u.getEmail ? u.getEmail() : "")).filter(e => e);
}

function TOPS_protectArchiveSheet_(sh, allowedEditorEmails) {
  const p = sh.protect();
  p.setDescription("Archived (protected)");
  p.setWarningOnly(false);
  try {
    p.setDomainEditorsAllowed(false);
    p.getEditors().forEach(ed => p.removeEditor(ed));
    if (allowedEditorEmails && allowedEditorEmails.length) p.addEditors(allowedEditorEmails);
  } catch (_) {}
}

function TOPS_archiveWorkbook_(label) {
  const ss = TOPS_ss_();
  const file = DriveApp.getFileById(ss.getId());

  const it = DriveApp.getFoldersByName(CL.TERM.ARCHIVE_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CL.TERM.ARCHIVE_FOLDER);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH.mm");
  const newName = `${ss.getName()} — Archive — ${label} — ${stamp}`;

  const copyFile = file.makeCopy(newName, folder);
  const copySS = SpreadsheetApp.openById(copyFile.getId());

  let allowedEditors = [];
  const policy = (CL.ARCHIVE && CL.ARCHIVE.EDITORS_POLICY) ? CL.ARCHIVE.EDITORS_POLICY : "OWNER_ONLY";
  if (policy === "FILE_EDITORS") allowedEditors = TOPS_getCurrentFileEditorEmails_();

  copySS.getSheets().forEach(sh => {
    const r = sh.getDataRange();
    r.setValues(r.getValues()); // values-only
    if (CL.ARCHIVE && CL.ARCHIVE.PROTECT_SHEETS) TOPS_protectArchiveSheet_(sh, allowedEditors);
  });

  return { id: copyFile.getId(), url: copyFile.getUrl() };
}

function TOPS_protectTermLog_(sh) {
  const policy = (CL.ARCHIVE && CL.ARCHIVE.EDITORS_POLICY) ? CL.ARCHIVE.EDITORS_POLICY : "OWNER_ONLY";
  const allowed = (policy === "FILE_EDITORS") ? TOPS_getCurrentFileEditorEmails_() : [];

  const existing = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
  const p = existing.length ? existing[0] : sh.protect();

  p.setDescription("Term Log (protected)");
  p.setWarningOnly(false);

  try {
    p.setDomainEditorsAllowed(false);
    p.getEditors().forEach(ed => p.removeEditor(ed));
    if (allowed.length) p.addEditors(allowed);
  } catch (_) {}
}

function TOPS_writeTermLogRow_(label, archiveUrl) {
  const ss = TOPS_ss_();
  let log = TOPS_getSheetByAliases_(CL.TERM.LOG_SHEET, "TERM_LOG");
  if (!log) log = ss.insertSheet(CL.TERM.LOG_SHEET);

  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 3).setValues([["TERM", "SHEET LINK", "CREATED AT"]]);
    log.setFrozenRows(1);
  }

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  log.appendRow([label, archiveUrl, stamp]);
  TOPS_protectTermLog_(log);
}

/* ============================================================
   ✅ PUBLIC: INIT FUNCTIONS (YOU RUN THESE FIRST)
============================================================ */

/** Initialize a NEW TERM run (does not auto-run steps). */
function termOps_InitNewTerm() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    try {
      TOPS_assertConfig_();

      const p = ui.prompt("Init: Start New Term", "Enter term label (e.g., 2025 TERM 2):", ui.ButtonSet.OK_CANCEL);
      if (p.getSelectedButton() !== ui.Button.OK) return;

      const termLabel = TOPS_normText_(p.getResponseText()).toUpperCase();
      if (!termLabel) return ui.alert("Term label required.");

      const classNames = TOPS_getClassSheets_().map(s => s.getName());

      TOPS_stateSet_({
        mode: "TERM",
        label: termLabel,
        step: 1,
        cursor: 0,
        classSheets: classNames,
        stats: { feeDefaults: 0, rolloverStudents: 0 }
      });

      ui.alert("Initialized ✅\nNow run: termOps_Step1()");
    } catch (err) {
      TOPS_fail_(ui, "Init New Term", err);
    }
  });
}

/** Initialize a NEW YEAR run (does not auto-run steps). */
function termOps_InitNewYear() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    try {
      TOPS_assertConfig_();

      const y = ui.prompt("Init: Start New Year", "Enter academic year (e.g., 2026):", ui.ButtonSet.OK_CANCEL);
      if (y.getSelectedButton() !== ui.Button.OK) return;

      const year = TOPS_normText_(y.getResponseText()).toUpperCase();
      if (!year) return ui.alert("Year required.");

      const gy = ui.prompt("Graduation Year", `Default: ${year}`, ui.ButtonSet.OK_CANCEL);
      if (gy.getSelectedButton() !== ui.Button.OK) return;

      const gradYear = TOPS_normText_(gy.getResponseText()).toUpperCase() || year;
      const termLabel = `${year} TERM 1`;

      const classNames = TOPS_getClassSheets_().map(s => s.getName());

      TOPS_stateSet_({
        mode: "YEAR",
        year,
        gradYear,
        label: termLabel,
        step: 1,
        promoIndex: (CL.CLASS_ORDER.length - 1), // downward
        cursor: 0,
        classSheets: classNames,
        stats: { feeDefaults: 0, promoted: 0, s6Debtors: 0, graduates: 0, rolloverStudents: 0 }
      });

      ui.alert("Initialized ✅\nNow run: termOps_Step1()");
    } catch (err) {
      TOPS_fail_(ui, "Init New Year", err);
    }
  });
}

/* ============================================================
   ✅ PUBLIC: RUN STEPS INDIVIDUALLY (EDITOR-RUNNABLE)
============================================================ */

function termOps_Step1() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run termOps_InitNewTerm() or termOps_InitNewYear() first.");
    try {
      TOPS_step1_Archive_(st);
      st.step = 2;
      TOPS_stateSet_(st);
      ui.alert("Step 1 done ✅\nNow run: termOps_Step2()");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, "Step 1", err);
    }
  });
}

function termOps_Step2() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run init first.");
    try {
      TOPS_step2_StructureFeesFormulas_(st);
      st.step = 3;
      TOPS_stateSet_(st);
      ui.alert("Step 2 done ✅\nNow run: termOps_Step3()");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, "Step 2", err);
    }
  });
}

function termOps_Step3() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run init first.");
    try {
      if (st.mode === "YEAR") {
        TOPS_step3_Promote_(st);
      } else {
        // TERM mode: promotion is skipped
        TOPS_toast_("TERM mode: Step 3 (Promotion) skipped.", "Step 3", 6);
      }
      st.step = 4;
      TOPS_stateSet_(st);
      ui.alert("Step 3 done ✅\nNow run: termOps_Step4()");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, "Step 3", err);
    }
  });
}

function termOps_Step4() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run init first.");
    try {
      TOPS_step4_Rollover_(st);
      st.step = 5;
      TOPS_stateSet_(st);
      ui.alert("Step 4 done ✅\nNow run: termOps_Step5()");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, "Step 4", err);
    }
  });
}

function termOps_Step5() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run init first.");
    try {
      TOPS_step5_HeavyFixes_(st);
      TOPS_stateClear_();
      ui.alert("Step 5 done ✅\nAll steps complete.");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, "Step 5", err);
    }
  });
}

/** Optional: run from current saved step to the end (like Resume). */
function termOps_RunAllRemaining() {
  const ui = SpreadsheetApp.getUi();
  return TOPS_withDocLock_(() => {
    const st = TOPS_stateGet_();
    if (!st) return ui.alert("No progress found. Run init first.");

    try {
      if (st.step <= 1) { TOPS_step1_Archive_(st); st.step = 2; TOPS_stateSet_(st); }
      if (st.step <= 2) { TOPS_step2_StructureFeesFormulas_(st); st.step = 3; TOPS_stateSet_(st); }
      if (st.step <= 3) {
        if (st.mode === "YEAR") TOPS_step3_Promote_(st);
        st.step = 4; TOPS_stateSet_(st);
      }
      if (st.step <= 4) { TOPS_step4_Rollover_(st); st.step = 5; TOPS_stateSet_(st); }
      if (st.step <= 5) { TOPS_step5_HeavyFixes_(st); TOPS_stateClear_(); }

      ui.alert("All remaining steps complete ✅");
    } catch (err) {
      TOPS_stateSet_(st);
      TOPS_fail_(ui, `RunAllRemaining (Step ${st.step})`, err);
    }
  });
}

/* ============================================================
   ✅ STEP IMPLEMENTATIONS (NO PROMPTS)
============================================================ */

function TOPS_step1_Archive_(st) {
  // TERM: archive once
  // YEAR: archive “Before Promotion”
  const label = (st.mode === "YEAR") ? `${st.year} NEW YEAR (Before Promotion)` : st.label;

  const snap = TOPS_archiveWorkbook_(label);
  TOPS_writeTermLogRow_(label, snap.url);

  st.snap1Url = snap.url;
  TOPS_toast_("Step 1 complete: Archived + logged.", "Step 1", 8);
}

function TOPS_step2_StructureFeesFormulas_(st) {
  // Ensure structure (LTB + BAL formulas) everywhere
  if (typeof ensureStructureAll === "function") ensureStructureAll();
  else TOPS_getClassSheets_().forEach(sh => { try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {} });

  // Apply fee defaults (blank/0 only)
  const ss = TOPS_ss_();
  const names = st.classSheets || [];
  let changedTotal = 0;

  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (!sh || !TOPS_isClassSheet_(sh)) continue;

    const changed = TOPS_applyFeeDefaultsOne_(sh);
    changedTotal += (changed || 0);

    // Ensure BAL formulas (safe)
    try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}
  }

  st.stats.feeDefaults = (st.stats.feeDefaults || 0) + changedTotal;
  TOPS_toast_(`Step 2 complete: Structure + fee defaults (${changedTotal}).`, "Step 2", 10);
}

function TOPS_step3_Promote_(st) {
  const ss = TOPS_ss_();
  const classOrder = CL.CLASS_ORDER;
  const gradSheet = TOPS_ensureGraduatesSheet_();
  const leavers = TOPS_ensureLeaversSheet_();

  // Work downward: Senior6 -> ... -> Baby
  let promotedTotal = 0;

  for (let i = classOrder.length - 1; i >= 0; i--) {
    const fromName = String(classOrder[i]);
    const from = ss.getSheetByName(fromName);
    if (!from) continue;

    try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(from); } catch (_) {}

    const region = TOPS_getStudentRegion_(from);
    const endRow = region.endRow;
    if (!endRow || endRow < 2) continue;

    const num = endRow - 1;
    const data = from.getRange(2, 1, num, from.getLastColumn()).getValues();

    // Senior 6: split debtors/graduates
    if (i === classOrder.length - 1) {
      const exact = TOPS_getHeaderIndexMap_(from);
      const upper = TOPS_headerMapUpper_(from);
      const cBAL = exact[CL.H.BAL] || upper[TOPS_flatUpper_(CL.H.BAL)];
      if (!cBAL) throw new Error(`"${fromName}" missing "${CL.H.BAL}" header.`);

      const debtors = [];
      const cleared = [];

      data.forEach((r, idx) => {
        const bal = Number(r[cBAL - 1]) || 0;
        const absRow = 2 + idx;
        if (bal > 0) debtors.push({ row: r, absRow });
        else cleared.push(r);
      });

      // Debtors -> Leavers
      if (debtors.length) {
        const leaversMap = TOPS_headerMapUpper_(leavers);
        const fromMap = TOPS_headerMapUpper_(from);

        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        const dateLeft = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

        const outRows = debtors.map(item => {
          const srcRow = item.row;
          const absRow = item.absRow;

          const out = new Array(Math.max(1, leavers.getLastColumn())).fill("");

          Object.keys(leaversMap).forEach(k => {
            const cTo = leaversMap[k];
            const cFrom = fromMap[k];
            if (cTo && cFrom) out[cTo - 1] = srcRow[cFrom - 1];
          });

          const set = (h, v) => {
            const c = leaversMap[TOPS_flatUpper_(h)];
            if (c) out[c - 1] = v;
          };

          set("Student ID", Utilities.getUuid());
          set("Former Class", fromName);
          set("Date Left", dateLeft);
          set("Exit Reason", "Graduation pending (Outstanding balance)");
          set("Status", "Unsettled");
          set("Source Sheet", fromName);
          set("Source Row", absRow);
          set("Created At", now);
          set("Updated At", now);

          return out;
        });

        TOPS_appendRowsAboveTotals_(leavers, outRows);
        st.stats.s6Debtors = (st.stats.s6Debtors || 0) + debtors.length;
      }

      // Cleared -> Graduates
      if (cleared.length) {
        const desired = TOPS_graduatesHeaders_();
        const gLast = gradSheet.getLastColumn();
        const fromMapExact = TOPS_getHeaderIndexMap_(from);
        const gUpper = TOPS_headerMapUpper_(gradSheet);
        const cGY = gUpper[TOPS_flatUpper_("Graduation Year")] || null;

        const out = cleared.map(srcRow => {
          const row = new Array(gLast).fill("");
          desired.forEach((h, idxH) => {
            if (TOPS_flatUpper_(h) === TOPS_flatUpper_("Graduation Year")) return;
            const cFrom = fromMapExact[h];
            if (cFrom) row[idxH] = srcRow[cFrom - 1];
          });
          if (cGY) row[cGY - 1] = st.gradYear;
          return row;
        });

        const start = gradSheet.getLastRow() + 1;
        TOPS_ensureRows_(gradSheet, start + out.length - 1);
        const rng = gradSheet.getRange(start, 1, out.length, out[0].length);
        rng.setValues(out);
        try { rng.clearFormat(); } catch (_) {}

        st.stats.graduates = (st.stats.graduates || 0) + cleared.length;
      }

      // Clear Senior 6 students (keep totals row)
      from.deleteRows(2, num);

      promotedTotal += data.length;
      st.stats.promoted = (st.stats.promoted || 0) + data.length;
      continue;
    }

    // Normal promotion: from class i -> class i+1
    const toName = String(classOrder[i + 1]);
    const to = ss.getSheetByName(toName);
    if (!to) continue;

    try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(to); } catch (_) {}

    let mapped = TOPS_mapRowsByHeaders_(from, to, data);
    mapped = TOPS_forceDestFee_(to, mapped);

    TOPS_appendRowsAboveTotals_(to, mapped);

    // Clear source students (keep totals row)
    from.deleteRows(2, num);

    promotedTotal += data.length;
    st.stats.promoted = (st.stats.promoted || 0) + data.length;
  }

  TOPS_toast_(`Step 3 complete: Promoted ${promotedTotal}`, "Step 3", 10);

  // Archive after promotion (Before Rollover)
  const snap2 = TOPS_archiveWorkbook_(st.label);
  TOPS_writeTermLogRow_(st.label, snap2.url);
  st.snap2Url = snap2.url;
}

function TOPS_step4_Rollover_(st) {
  const ss = TOPS_ss_();
  const names = st.classSheets || [];
  let touchedTotal = 0;

  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (!sh || !TOPS_isClassSheet_(sh)) continue;

    try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}

    const exact = TOPS_getHeaderIndexMap_(sh);
    const upper = TOPS_headerMapUpper_(sh);

    const cNAME = exact[CL.H.NAME] || upper[TOPS_flatUpper_(CL.H.NAME)];
    const cLTB  = exact[CL.H.LTB]  || upper[TOPS_flatUpper_(CL.H.LTB)];
    const cPAID = exact[CL.H.PAID] || upper[TOPS_flatUpper_(CL.H.PAID)];
    const cBAL  = exact[CL.H.BAL]  || upper[TOPS_flatUpper_(CL.H.BAL)];
    if (!cNAME || !cLTB || !cPAID || !cBAL) continue;

    const region = TOPS_getStudentRegion_(sh);
    if (!region || region.endRow < 2) continue;

    const nRows = region.endRow - 1;

    const namesCol = sh.getRange(2, cNAME, nRows, 1).getValues();
    const balVal   = sh.getRange(2, cBAL,  nRows, 1).getValues();
    const ltbOld   = sh.getRange(2, cLTB,  nRows, 1).getValues();
    const paidOld  = sh.getRange(2, cPAID, nRows, 1).getValues();

    const isStudent = (v) => {
      const s = TOPS_flatUpper_(v);
      return s && s !== "TOTALS";
    };

    const newLTB = ltbOld.map((v, idx) => {
      if (isStudent(namesCol[idx][0])) { touchedTotal++; return [Number(balVal[idx][0]) || 0]; }
      return [v[0]];
    });

    const newPaid = paidOld.map((v, idx) => [isStudent(namesCol[idx][0]) ? "" : v[0]]);

    sh.getRange(2, cLTB,  nRows, 1).setValues(newLTB);
    sh.getRange(2, cPAID, nRows, 1).setValues(newPaid);

    // Re-apply BAL formulas safely (student region only)
    try { if (typeof ensureStructureOne_ === "function") ensureStructureOne_(sh); } catch (_) {}
  }

  st.stats.rolloverStudents = (st.stats.rolloverStudents || 0) + touchedTotal;
  TOPS_toast_(`Step 4 complete: Rollover students ${touchedTotal}`, "Step 4", 10);
}

function TOPS_step5_HeavyFixes_(st) {
  // Totals
  if (typeof fixTotalsAllSheets === "function") fixTotalsAllSheets();

  // NO-GREEN enforcement + trim below totals (your updated Totals does remove rules)
  if (typeof fixTotalsFormattingAll_ === "function") fixTotalsFormattingAll_();

  // Defaulters rebuild
  if (typeof rebuildAllDefaulters_QUERY === "function") rebuildAllDefaulters_QUERY();

  // Optional: apply workbook layout formatting (safe) if you want it here
  // if (typeof formatWorkbookLayout === "function") formatWorkbookLayout();

  TOPS_toast_("Step 5 complete: Totals + No-green + Defaulters.", "Step 5", 10);
}

/* =========================
   Health check (optional)
========================= */

function termOps_HealthCheck() {
  const ui = SpreadsheetApp.getUi();
  try {
    TOPS_assertConfig_();
    const ss = TOPS_ss_();
    const missing = [];
    (CL.CLASS_ORDER || []).forEach(n => { if (!ss.getSheetByName(String(n))) missing.push(String(n)); });
    if (missing.length) ui.alert("Missing sheets in CL.CLASS_ORDER:\n\n" + missing.join("\n"));
    else ui.alert("HealthCheck ✅ Config + CL.CLASS_ORDER sheets look OK.");
  } catch (e) {
    TOPS_fail_(ui, "HealthCheck", e);
  }
}
