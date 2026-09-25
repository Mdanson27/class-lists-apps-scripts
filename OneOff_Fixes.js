  /**
   * ===== ONE-OFF + MAINTENANCE FIXES =====
   * v3.0-absolute-safe-nightly
   *
   * GUARANTEES:
   * ✅ Will NEVER sort a sheet unless TOTALS is EXACTLY at the fixed row in FIXED_TOTALS_ROWS.
   * ✅ Will NEVER include TOTALS row in sorting.
   * ✅ Will NEVER sort blank/green zones (sorts only up to last real student row).
   * ✅ Will NEVER insert/delete rows above TOTALS.
   *
   * What you run:
   * 1) VALIDATE_FixedTotalsMap_AllClassSheets()
   * 2) Fix any sheets it reports (move TOTALS back to the fixed row; no students under totals)
   * 3) MANUAL_SortAndRenumber_OneSheet("Primary 5 list") to test
   * 4) NIGHTLY_SortAndRenumber_AllClassSheets()
   * 5) INSTALL_NightlySortTrigger_AtHour(2) once (optional)
   */

  /* =========================================================
  * 0) FIXED TOTALS ROW MAP (AUTHORITATIVE)
  * ========================================================= */
  const FIXED_TOTALS_ROWS = {
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

  /* =========================================================
  * 1) SAFE HELPERS
  * ========================================================= */

  function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

  function _toast(msg, title = "Info", seconds = 6) {
    try { _ss().toast(String(msg || ""), String(title || "Info"), Number(seconds || 6)); } catch (_) {}
  }

  function _norm(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\u00A0/g, " ").trim().replace(/\s+/g, " ");
  }
  function _flatUpper(v) {
    return _norm(v).toUpperCase().replace(/\s+/g, "");
  }

  function _withDocLock(fn) {
    const lock = LockService.getDocumentLock();
    const start = Date.now();
    const maxMs = 30 * 1000;

    while (true) {
      try { if (lock.tryLock(3000)) break; } catch (_) {}
      if (Date.now() - start > maxMs) throw new Error("Lock timeout: another execution is running.");
      Utilities.sleep(600);
    }

    try { return fn(); }
    finally { try { lock.releaseLock(); } catch (_) {} }
  }

  function _getHeaderIndexMap(sh) {
    try { if (typeof getHeaderIndexMap_ === "function") return getHeaderIndexMap_(sh); } catch (_) {}
    const lastCol = Math.max(1, sh.getLastColumn());
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_norm);
    const map = {};
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (!h) continue;
      if (map[h] == null) map[h] = i + 1;
    }
    return map;
  }

  function _lastHeaderCol(sh) {
    const lastCol = Math.max(1, sh.getLastColumn());
    const row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    let end = 1;
    for (let c = 0; c < row1.length; c++) {
      if (_norm(row1[c])) end = c + 1;
    }
    return end;
  }

  function _isTotalsLabel(v) {
    const t = _flatUpper(v);
    return t === "TOTALS" || t.startsWith("TOTALS");
  }

  function _findActualTotalsRowInColA_(sh) {
    const last = Math.min(sh.getLastRow(), 4000);
    if (last < 2) return null;

    const colA = sh.getRange(1, 1, last, 1).getDisplayValues().flat();
    for (let i = colA.length - 1; i >= 0; i--) {
      if (_isTotalsLabel(colA[i])) return i + 1;
    }
    return null;
  }

  /** Deterministic class sheets: CL.CLASS_ORDER only (strict fallback suffix) */
  function _getClassSheets_() {
    const ss = _ss();

    try {
      if (typeof CL !== "undefined" && CL && Array.isArray(CL.CLASS_ORDER) && CL.CLASS_ORDER.length) {
        const out = [];
        CL.CLASS_ORDER.forEach(n => {
          const sh = ss.getSheetByName(String(n));
          if (sh) out.push(sh);
        });
        return out;
      }
    } catch (_) {}

    const suf = String(CL?.CLASS_SUFFIX || " list").toLowerCase();
    return ss.getSheets().filter(sh => String(sh.getName()).toLowerCase().endsWith(suf));
  }

  /**
   * STRICT safety validation for one sheet.
   * Returns {ok:boolean, reason:string, expected:number, actual:number|null}
   */
  function _validateSheetForNightlySort_(sh) {
    const name = sh.getName();
    const expected = Number(FIXED_TOTALS_ROWS[name]);

    if (!Number.isFinite(expected) || expected < 3) {
      return { ok: false, reason: "Sheet not in FIXED_TOTALS_ROWS map.", expected: expected, actual: null };
    }

    // Must physically have that row
    if (sh.getMaxRows() < expected) {
      return {
        ok: false,
        reason: `Sheet has only ${sh.getMaxRows()} rows but expected TOTALS row is ${expected}. (Do NOT auto-insert; fix sheet manually.)`,
        expected,
        actual: null
      };
    }

    // Totals label must be exactly at fixed row, column A
    const a = sh.getRange(expected, 1).getDisplayValue();
    if (!_isTotalsLabel(a)) {
      const actual = _findActualTotalsRowInColA_(sh);
      return {
        ok: false,
        reason: `TOTALS label is NOT at fixed row ${expected}. Found actual TOTALS row: ${actual || "none"}.`,
        expected,
        actual: actual || null
      };
    }

    // Student Name column
    const map = _getHeaderIndexMap(sh);
    const cName = map[CL.H.NAME] || 3;

    // Ensure no students below totals (we only scan a safe window to detect obvious breaks)
    const scanBelow = Math.min(sh.getLastRow(), expected + 40);
    if (scanBelow > expected) {
      const belowVals = sh.getRange(expected + 1, cName, scanBelow - expected, 1).getDisplayValues().flat();
      const hasStudentBelow = belowVals.some(v => _norm(v));
      if (hasStudentBelow) {
        return {
          ok: false,
          reason: `There are student names BELOW the fixed TOTALS row ${expected}. (TOTALS has moved or data is under it.)`,
          expected,
          actual: expected
        };
      }
    }

    return { ok: true, reason: "OK", expected, actual: expected };
  }

  /** Find last real student row (by Student Name) above TOTALS */
  function _findLastStudentRow_(sh, totalsRow, cName) {
    const start = 2;
    const end = Math.max(start, totalsRow - 1);
    if (end < start) return 1;

    const n = end - start + 1;
    const vals = sh.getRange(start, cName, n, 1).getDisplayValues().flat();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (_norm(vals[i])) return start + i;
    }
    return 1;
  }

  /* =========================================================
  * 2) NIGHTLY SORT + RENUMBER (ABSOLUTE SAFE)
  * ========================================================= */

  function NIGHTLY_SortAndRenumber_AllClassSheets() {
    return _withDocLock(() => {
      const sheets = _getClassSheets_();
      const errors = [];
      const skipped = [];

      sheets.forEach(sh => {
        try {
          const v = _validateSheetForNightlySort_(sh);
          if (!v.ok) {
            skipped.push(`${sh.getName()}: ${v.reason}`);
            return;
          }
          _sortAndRenumberOne_Strict_(sh, v.expected);
        } catch (e) {
          errors.push(`${sh.getName()}: ${e && e.message ? e.message : e}`);
        }
      });

      if (skipped.length) {
        Logger.log("NIGHTLY skipped sheets:\n" + skipped.join("\n"));
      }
      if (errors.length) {
        Logger.log("NIGHTLY errors:\n" + errors.join("\n"));
      }

      if (errors.length) {
        _toast(`Nightly sort finished with ${errors.length} ERROR(s). Check Logs.`, "Nightly Sort", 9);
        return;
      }

      if (skipped.length) {
        _toast(`Nightly sort completed. Skipped ${skipped.length} sheet(s) due to safety checks.`, "Nightly Sort", 9);
        return;
      }

      _toast("Nightly sort + NO. renumber completed safely for all class sheets.", "Nightly Sort", 7);
    });
  }

  function MANUAL_SortAndRenumber_OneSheet(sheetName) {
    return _withDocLock(() => {
      const sh = _ss().getSheetByName(String(sheetName || ""));
      if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

      const v = _validateSheetForNightlySort_(sh);
      if (!v.ok) throw new Error(`SAFETY STOP — ${v.reason}`);

      _sortAndRenumberOne_Strict_(sh, v.expected);
      _toast(`Sorted + renumbered safely: ${sh.getName()}`, "Sort", 6);
    });
  }

  /**
   * Strict sorter:
   * - Sorts ONLY rows 2..lastStudentRow (never to totalsRow-1)
   * - Renumbers NO. only in that used zone
   * - Clears NO. below used zone (above totals)
   */
  function _sortAndRenumberOne_Strict_(sh, totalsRow) {
    const map = _getHeaderIndexMap(sh);
    const cName = map[CL.H.NAME] || 3;

    const lastStudentRow = _findLastStudentRow_(sh, totalsRow, cName);
    if (lastStudentRow < 2) {
      // No students — clear NO. area above totals (optional)
      const blanks = totalsRow - 2;
      if (blanks > 0) sh.getRange(2, 1, blanks, 1).clearContent();
      return;
    }

    const tableCols = _lastHeaderCol(sh);

    // Sort only the used block (prevents blank/green zones being dragged upward)
    const rng = sh.getRange(2, 1, lastStudentRow - 1, tableCols);
    rng.sort([{ column: cName, ascending: true }]);

    // Renumber NO. for used block based on non-empty Student Name
    const nameVals = sh.getRange(2, cName, lastStudentRow - 1, 1).getDisplayValues().flat();
    const noOut = [];
    let n = 0;
    for (let i = 0; i < nameVals.length; i++) {
      const hasName = !!_norm(nameVals[i]);
      if (hasName) n++;
      noOut.push([hasName ? n : ""]);
    }
    sh.getRange(2, 1, noOut.length, 1).setValues(noOut);

    // Clear NO. for remaining rows above totals (so numbering doesn’t show in blank area)
    const remaining = (totalsRow - 1) - (lastStudentRow);
    if (remaining > 0) {
      sh.getRange(lastStudentRow + 1, 1, remaining, 1).clearContent();
    }

    // Final safety: totals label must still be present at fixed row
    const a = sh.getRange(totalsRow, 1).getDisplayValue();
    if (!_isTotalsLabel(a)) {
      throw new Error("SAFETY STOP — totals label moved during operation (unexpected). No further actions.");
    }
  }

  /* =========================================================
  * 3) VALIDATOR YOU RUN FIRST
  * ========================================================= */

  /**
   * Run this FIRST.
   * It will report which sheets are SAFE and which are NOT SAFE for nightly sort.
   */
  function VALIDATE_FixedTotalsMap_AllClassSheets() {
    const sheets = _getClassSheets_();
    const bad = [];
    const ok = [];

    sheets.forEach(sh => {
      const v = _validateSheetForNightlySort_(sh);
      if (v.ok) ok.push(`${sh.getName()}: OK (TOTALS at row ${v.expected})`);
      else bad.push(`${sh.getName()}: ${v.reason}`);
    });

    Logger.log("VALIDATE OK:\n" + ok.join("\n"));
    if (bad.length) Logger.log("VALIDATE PROBLEMS:\n" + bad.join("\n"));

    if (bad.length) {
      _toast(`Validation found ${bad.length} sheet(s) NOT safe. Open Apps Script Logs to see details.`, "Validate", 10);
    } else {
      _toast("All class sheets are SAFE for nightly sort.", "Validate", 7);
    }
  }

  /* =========================================================
  * 4) TRIGGER INSTALLER
  * ========================================================= */

  function INSTALL_NightlySortTrigger_AtHour(hour0to23) {
    const hour = Number(hour0to23);
    if (!Number.isFinite(hour) || hour < 0 || hour > 23) throw new Error("Hour must be 0..23");

    const fnName = "NIGHTLY_SortAndRenumber_AllClassSheets";
    const triggers = ScriptApp.getProjectTriggers();

    triggers.forEach(t => {
      if (t.getHandlerFunction && t.getHandlerFunction() === fnName) {
        try { ScriptApp.deleteTrigger(t); } catch (_) {}
      }
    });

    ScriptApp.newTrigger(fnName)
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .create();

    _toast(`Nightly sort trigger installed (daily at ${hour}:00).`, "Trigger", 7);
  }

  /* =========================================================
  * 5) KEEP: Senior 6 one-off (unchanged from your last working version)
  * ========================================================= */

  function FIX_SENIOR6_restoreBalanceFormulas_keepTotalsOnly() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shName = "Senior 6 list";
    const sh = ss.getSheetByName(shName);
    if (!sh) throw new Error(`Sheet not found: "${shName}"`);

    const COL_LTB = 7;   // G
    const COL_FEE = 8;   // H
    const COL_PAID = 9;  // I
    const COL_BAL = 10;  // J (BALANCE)

    const totalsRow = Number(FIXED_TOTALS_ROWS[shName]);
    if (!totalsRow || totalsRow < 3) throw new Error(`Invalid fixed TOTALS row for "${shName}".`);
    if (!_isTotalsLabel(sh.getRange(totalsRow, 1).getDisplayValue())) {
      throw new Error(`TOTALS label is not at fixed row ${totalsRow} on "${shName}". Fix it first.`);
    }

    const studentStart = 2;
    const studentEnd = totalsRow - 1;
    const numRows = studentEnd - studentStart + 1;
    if (numRows <= 0) throw new Error(`No student rows found above TOTALS in "${shName}".`);

    const lastCol = Math.max(1, sh.getLastColumn());

    let balTemplateR1C1 = null;

    // 1) From Senior 6 itself
    {
      const formulas = sh.getRange(studentStart, COL_BAL, numRows, 1).getFormulasR1C1();
      for (let i = 0; i < formulas.length; i++) {
        const f = formulas[i][0];
        if (f && String(f).trim()) { balTemplateR1C1 = f; break; }
      }
    }

    // 2) From another class list
    if (!balTemplateR1C1) {
      const suf = String(CL?.CLASS_SUFFIX || " list").toLowerCase();
      for (const other of ss.getSheets()) {
        const nm = other.getName();
        if (nm === shName) continue;
        if (!String(nm).toLowerCase().endsWith(suf)) continue;

        const otherTotals = Number(FIXED_TOTALS_ROWS[nm]) || null;
        if (!otherTotals || otherTotals < 3) continue;

        const otherEnd = otherTotals - 1;
        const otherNumRows = otherEnd - 1;
        if (otherNumRows <= 0) continue;

        const fR1C1 = other.getRange(2, COL_BAL, otherNumRows, 1).getFormulasR1C1();
        for (let i = 0; i < fR1C1.length; i++) {
          const f = fR1C1[i][0];
          if (f && String(f).trim()) { balTemplateR1C1 = f; break; }
        }
        if (balTemplateR1C1) break;
      }
    }

    if (!balTemplateR1C1) {
      throw new Error(
        'Could not find a BALANCE formula template.\n' +
        'Ensure at least ONE other "... list" sheet has correct Balance formulas in column J.'
      );
    }

    // Convert student area to values (clears bad formulas), then restore balance formulas
    const studentRange = sh.getRange(studentStart, 1, numRows, lastCol);
    studentRange.setValues(studentRange.getValues());
    sh.getRange(studentStart, COL_BAL, numRows, 1).setFormulaR1C1(balTemplateR1C1);

    // Fix totals formulas
    sh.getRange(totalsRow, 1).setValue("TOTALS");
    sh.getRange(totalsRow, COL_LTB, 1, 4).clearContent();
    sh.getRange(totalsRow, COL_LTB).setFormula(`=SUM(G${studentStart}:G${studentEnd})`);
    sh.getRange(totalsRow, COL_FEE).setFormula(`=SUM(H${studentStart}:H${studentEnd})`);
    sh.getRange(totalsRow, COL_PAID).setFormula(`=SUM(I${studentStart}:I${studentEnd})`);
    sh.getRange(totalsRow, COL_BAL).setFormula(`=SUM(J${studentStart}:J${studentEnd})`);

    ss.toast(`Senior 6 fixed safely.`, "Senior 6 Fix", 7);
  }