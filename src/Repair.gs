/**
 * ==========================================
 * REPAIR, MIGRATION & FORMULA INTEGRITY
 * ==========================================
 * Non-destructive operations safe to run against a live sheet at any time.
 * Nothing in this file calls sheet.clear().
 */

const INTEGRITY_PROP_KEY = "wealthscript.formulaBaseline";

/* ------------------------------------------------------------------ *
 * FORMULA INTEGRITY
 * ------------------------------------------------------------------ */

/**
 * Pure helper: compares current formula counts against a stored baseline.
 * A DROP is the signal — growth is fine (the user added accounts).
 *
 * @param {Object<string,number>} current - label -> formula count
 * @param {Object<string,number>} baseline - label -> formula count
 * @returns {{ok: boolean, regressions: Array<{label: string, was: number, now: number}>}}
 */
function _diffFormulaCounts(current, baseline) {
  const regressions = [];
  Object.keys(baseline || {}).forEach(label => {
    const was = Number(baseline[label]) || 0;
    const now = Number(current && current[label]) || 0;
    if (now < was) regressions.push({ label: label, was: was, now: now });
  });
  return { ok: regressions.length === 0, regressions: regressions };
}

/**
 * Counts non-empty formulas in each managed range.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Object<string,number>}
 */
function _countManagedFormulas(ss) {
  const counts = {};
  _managedRanges().forEach(spec => {
    const sheet = ss.getSheetByName(spec.sheet);
    if (!sheet) { counts[spec.label] = 0; return; }
    let n = 0;
    sheet.getRange(spec.a1).getFormulas().forEach(row => {
      row.forEach(f => { if (f) n++; });
    });
    counts[spec.label] = n;
  });
  return counts;
}

/**
 * Pure helper: classifies Brokerage rows against the Holdings tab.
 *
 * A missing formula is only damage when there is something to link TO. Tracking
 * a brokerage as a single manually-typed number is a legitimate pattern, and
 * injecting a SUMPRODUCT there would evaluate to 0 and silently erase the
 * balance. So the rule is: broken only if holdings exist under that exact name.
 *
 * A hardcoded ZERO is never intentional tracking. It is the fingerprint of a
 * SUMPRODUCT that was already returning 0 (usually an account-name mismatch)
 * at the moment something froze the column into literals. Those rows are
 * reported as SUSPECT: repair cannot fix them, because there is nothing to link
 * to until a human reconciles the names.
 *
 * @param {Array<Array<*>>} ledgerRows - A..J values from LEDGER_FIRST_ROW
 * @param {Object<string,number>} holdingAccounts - account name -> holdings row count
 * @param {Array<string>} formulaCol - column E formulas, parallel to ledgerRows
 * @param {Array<string>} [linkedClasses] - asset classes sourced from Holdings
 * @returns {{broken: Array, suspect: Array, unlinked: Array, orphaned: Array}}
 */
function _classifyBrokerageRows(ledgerRows, holdingAccounts, formulaCol, linkedClasses) {
  const classes = linkedClasses || DASHBOARD_CONFIG.holdingsLinkedClasses || ["Brokerage"];
  const broken = [], suspect = [], unlinked = [];
  const seen = {};

  for (let i = 0; i < ledgerRows.length; i++) {
    const assetClass = String(ledgerRows[i][1]).trim();
    if (classes.indexOf(assetClass) === -1) continue;

    const account = String(ledgerRows[i][0] || "").trim();
    const status = String(ledgerRows[i][9] || "").trim();
    const value = ledgerRows[i][4];
    const row = LEDGER_FIRST_ROW + i;
    seen[account] = true;

    if (formulaCol[i]) continue;                       // linked — nothing to do

    const base = { row: row, account: account, status: status, assetClass: assetClass, value: value };

    if (holdingAccounts[account]) {
      broken.push(Object.assign({ holdings: holdingAccounts[account] }, base));
    } else if (value === "" || value === null || Number(value) === 0) {
      suspect.push(base);
    } else {
      unlinked.push(base);
    }
  }

  // Holdings whose account name matches no Brokerage row: real money that the
  // dashboard is not counting anywhere.
  const orphaned = Object.keys(holdingAccounts)
    .filter(n => !seen[n])
    .map(n => ({ account: n, rows: holdingAccounts[n] }));

  return { broken: broken, suspect: suspect, unlinked: unlinked, orphaned: orphaned };
}

/**
 * Reads the sheet and classifies every Brokerage row.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {{broken: Array, unlinked: Array, orphaned: Array}}
 */
function _auditBrokerageLinks(ss) {
  const ledger = ss.getSheetByName("Dashboard & Ledger");
  const holdings = ss.getSheetByName("Brokerage Holdings");
  if (!ledger) return { broken: [], suspect: [], unlinked: [], orphaned: [] };

  const holdingAccounts = {};
  if (holdings) {
    holdings.getRange(HOLDINGS_FIRST_ROW, 1, HOLDINGS_NUM_ROWS, 1).getValues()
      .forEach(r => {
        const n = String(r[0] || "").trim();
        if (n) holdingAccounts[n] = (holdingAccounts[n] || 0) + 1;
      });
  }

  const rows = ledger.getRange(LEDGER_FIRST_ROW, 1, LEDGER_NUM_ROWS, 10).getValues();
  const formulaCol = ledger.getRange(LEDGER_FIRST_ROW, 5, LEDGER_NUM_ROWS, 1)
    .getFormulas().map(r => r[0]);

  return _classifyBrokerageRows(rows, holdingAccounts, formulaCol);
}

/**
 * Full health check. Called on demand from the menu and defensively by
 * captureSnapshot() before it writes a row.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @returns {{healthy: boolean, regressions: Array, brokenLinks: Array, counts: Object, hadBaseline: boolean}}
 */
function auditFormulaHealth(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  const counts = _countManagedFormulas(ss);
  const links = _auditBrokerageLinks(ss);

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(INTEGRITY_PROP_KEY);
  let baseline = null;
  try { baseline = raw ? JSON.parse(raw) : null; } catch (e) { baseline = null; }

  const diff = _diffFormulaCounts(counts, baseline);

  return {
    // Only genuine damage blocks a snapshot. Unlinked and orphaned accounts are
    // reported for review but may well be intentional.
    healthy: diff.ok && links.broken.length === 0 && links.suspect.length === 0,
    regressions: diff.regressions,
    brokenLinks: links.broken,
    suspect: links.suspect,
    unlinked: links.unlinked,
    orphaned: links.orphaned,
    counts: counts,
    hadBaseline: !!baseline
  };
}

/** Records the current formula counts as the accepted baseline. */
function _saveFormulaBaseline(ss) {
  PropertiesService.getDocumentProperties()
    .setProperty(INTEGRITY_PROP_KEY, JSON.stringify(_countManagedFormulas(ss)));
}

/** Pure helper: renders an audit result as a human-readable report. */
function _formatHealthReport(audit) {
  const notices = [];

  (audit.unlinked || []).forEach(u => notices.push(
    `  • Row ${u.row} — ${u.account}${u.status ? ` (${u.status})` : ""}: manual value, no holdings tracked. Fine if intentional.`));
  (audit.orphaned || []).forEach(o => notices.push(
    `  • "${o.account}" has ${o.rows} holdings row(s) but NO Brokerage row on the ledger — this money is not counted in your net worth.`));

  const lines = [];

  if (audit.healthy) {
    lines.push(audit.hadBaseline
      ? "✅ All managed formulas are intact."
      : "✅ No damage detected. Baseline recorded — future checks compare against it.");
  } else {
    lines.push("⚠️ Formula damage detected.", "");
    if (audit.regressions.length) {
      lines.push("Formula count dropped in:");
      audit.regressions.forEach(r => lines.push(`  • ${r.label} — was ${r.was}, now ${r.now}`));
      lines.push("");
    }
    if (audit.brokenLinks.length) {
      lines.push("Rows with holdings but no link (Repair Formulas fixes these):");
      audit.brokenLinks.forEach(b => lines.push(
        `  • Row ${b.row} — ${b.account} [${b.assetClass}] (${b.holdings} holdings row(s) waiting)`));
      lines.push("");
    }
    if ((audit.suspect || []).length) {
      lines.push("Rows frozen at 0 with no matching holdings — NEEDS YOUR ATTENTION:");
      audit.suspect.forEach(x => lines.push(
        `  • Row ${x.row} — ${x.account} [${x.assetClass}]. Counting as $0 in your net worth.`));
      lines.push("    Repair cannot fix these: the account name likely doesn't match the");
      lines.push("    Holdings tab. Reconcile the names, then run Repair Formulas.");
      lines.push("");
    }
    if (audit.brokenLinks.length) lines.push("Run WealthScript > 🛠 Repair Formulas to restore the links above.");
  }

  if (notices.length) {
    lines.push("", "ℹ️ For review (not errors, nothing will be changed):", ...notices);
  }

  return lines.join("\n");
}

/** Menu action: report formula health and record a baseline if none exists. */
function checkFormulaHealth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const audit = auditFormulaHealth(ss);
  if (!audit.hadBaseline) _saveFormulaBaseline(ss);
  SpreadsheetApp.getUi().alert(_formatHealthReport(audit));
}

/* ------------------------------------------------------------------ *
 * REPAIR
 * ------------------------------------------------------------------ */

/**
 * Re-injects every managed formula into a LIVE sheet without clearing anything.
 *
 * Safety rule: a cell holding a hardcoded literal in a manual-entry range is
 * never overwritten. That protects deliberately pinned values — option contract
 * prices on the Holdings tab, manually typed balances on the ledger — which a
 * naive fill-down would destroy.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{restored: number, preserved: number, details: Array<string>}}
 */
function repairFormulas(ss_inject, silent = false) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName("Dashboard & Ledger");
  const holdings = ss.getSheetByName("Brokerage Holdings");

  const details = [];
  let restored = 0, preserved = 0;

  if (!ledger || !holdings) {
    return { restored: 0, preserved: 0, details: ["Required tabs not found — run First Time Setup."] };
  }

  // --- 1. Fixed KPI cells. These are always formulas; a literal here is damage.
  const fixed = _ledgerFixedFormulas();
  Object.keys(fixed).forEach(a1 => {
    const cell = ledger.getRange(a1);
    if (cell.getFormula() !== fixed[a1]) {
      cell.setFormula(fixed[a1]);
      restored++;
      details.push(`Ledger ${a1} restored`);
    }
  });

  // --- 2. Ledger per-row columns F, H, I. Always formulas.
  for (let r = LEDGER_FIRST_ROW; r <= LEDGER_LAST_ROW; r++) {
    const f = _ledgerRowFormulas(r);
    ["F", "H", "I"].forEach(col => {
      const cell = ledger.getRange(`${col}${r}`);
      if (cell.getFormula() !== f[col]) { cell.setFormula(f[col]); restored++; }
    });
  }

  // --- 3. Ledger column E, Brokerage rows.
  //        CRITICAL: only link rows whose account name actually appears on the
  //        Holdings tab. Injecting a SUMPRODUCT that matches nothing evaluates
  //        to 0 and would silently erase a manually maintained balance.
  const links = _auditBrokerageLinks(ss);
  links.broken.forEach(b => {
    const want = _buildBrokerageFormula(b.row);
    const cell = ledger.getRange(b.row, 5);
    if (cell.getFormula() === want) return;
    if (cell.getFormula() === "" && cell.getValue() !== "") {
      details.push(`Row ${b.row} (${b.account}): replaced a frozen literal with the live Holdings link`);
    }
    cell.setFormula(want);
    restored++;
  });
  links.unlinked.forEach(u => {
    preserved++;
    details.push(`Row ${u.row} (${u.account}): kept manual value ${u.value} — no holdings under this name`);
  });
  links.suspect.forEach(x => {
    details.push(`⚠️ Row ${x.row} (${x.account}) frozen at 0 with no matching holdings — counting as $0. Fix the account name, then re-run Repair.`);
  });
  links.orphaned.forEach(o => {
    details.push(`⚠️ "${o.account}" has ${o.rows} holdings row(s) but no ledger row — not counted in net worth`);
  });

  // --- 4. Holdings E (price) and F (total). E is skipped where the user has
  //        pinned a literal — options and unquotable tickers have no live price.
  for (let r = HOLDINGS_FIRST_ROW; r <= HOLDINGS_LAST_ROW; r++) {
    const f = _holdingsRowFormulas(r);

    const priceCell = holdings.getRange(r, 5);
    const hasLiteralPrice = priceCell.getFormula() === "" && priceCell.getValue() !== "";
    if (hasLiteralPrice) {
      preserved++;
      details.push(`Holdings E${r}: kept pinned price ${priceCell.getValue()}`);
    } else if (priceCell.getFormula() !== f.E) {
      priceCell.setFormula(f.E);
      restored++;
    }

    const totalCell = holdings.getRange(r, 6);
    if (totalCell.getFormula() !== f.F) { totalCell.setFormula(f.F); restored++; }
  }

  _saveFormulaBaseline(ss);

  if (!silent) {
    const msg = [`✅ Repair complete.`, ``,
                 `Formulas restored: ${restored}`,
                 `Pinned values preserved: ${preserved}`, ``]
      .concat(details.length ? ["Notes:"].concat(details.slice(0, 25).map(d => "  • " + d)) : [])
      .join("\n");
    SpreadsheetApp.getUi().alert(msg);
  }

  return { restored: restored, preserved: preserved, details: details };
}

/* ------------------------------------------------------------------ *
 * MIGRATION
 * ------------------------------------------------------------------ */

/**
 * Brings a pre-existing sheet up to the current layout. Idempotent — running
 * it twice is a no-op.
 *
 * Migration 1: insert Settings!B22 (FIRE Target), shifting currencies to B24/B25.
 * Migration 2: create hidden FX helper cells M:O on the Dashboard.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{applied: Array<string>, skipped: Array<string>}}
 */
function migrateSheetLayout(ss_inject, silent = false) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName("Settings & Config");
  const ledger = ss.getSheetByName("Dashboard & Ledger");

  const applied = [], skipped = [];

  if (!cfg || !ledger) {
    return { applied: [], skipped: ["Required tabs not found — run First Time Setup."] };
  }

  // --- Migration 1: FIRE target row ---
  if (String(cfg.getRange("A22").getValue()).indexOf("FIRE Target") === 0) {
    skipped.push("Settings!B22 FIRE target already present");
  } else {
    cfg.insertRowBefore(22);
    cfg.getRange("A22:B22")
      .setValues([["FIRE Target Net Worth (USD)", (DASHBOARD_CONFIG.fireTargetUSD || 3000000)]])
      .setBackground(THEME.kpiCardBg).setVerticalAlignment("middle");
    cfg.getRange("B22").setNumberFormat("$#,##0")
      .setNote("The FIRE Progress card and the Time-to-FIRE chart read this cell live. Change it any time — no rebuild needed.");
    applied.push("Inserted Settings!B22 (FIRE Target); currencies now B24/B25, ZPID map now A29:B45");
  }

  // --- Migration 2: hidden FX helper cells ---
  if (ledger.getRange("M2").getFormula() && ledger.getRange("O3").getFormula()) {
    skipped.push("Hidden helper cells M:O already present");
  } else {
    ledger.getRange("M1").setValue("⚙ INTERNAL — do not edit or delete")
      .setFontColor(THEME.mutedText).setFontSize(8);
    ledger.getRange("M2:M3").setNumberFormat("0.000000");
    ledger.getRange("N2:O3").setNumberFormat("#,##0.00");
    applied.push("Created hidden FX helper cells M1:O3");
  }

  // Formulas + column hiding are idempotent, so always reassert them.
  const fixed = _ledgerFixedFormulas();
  Object.keys(fixed).forEach(a1 => ledger.getRange(a1).setFormula(fixed[a1]));
  ["E2", "E3", "I2", "I3"].forEach(a1 => ledger.getRange(a1).setNumberFormat("@"));
  ledger.hideColumns(13, 3);   // M, N, O
  applied.push("Reasserted KPI card formulas and text formatting");

  _saveFormulaBaseline(ss);

  if (!silent) {
    const msg = ["🧭 Migration complete.", ""]
      .concat(applied.length ? ["Applied:"].concat(applied.map(a => "  • " + a)) : [])
      .concat(skipped.length ? ["", "Already up to date:"].concat(skipped.map(s => "  • " + s)) : [])
      .join("\n");
    SpreadsheetApp.getUi().alert(msg);
  }

  return { applied: applied, skipped: skipped };
}

/* ------------------------------------------------------------------ *
 * DESTRUCTIVE OPERATION GUARDS
 * ------------------------------------------------------------------ */

/**
 * Pure helper: does this ledger hold real user data?
 * Sample rows shipped by First Time Setup don't count as real data.
 *
 * @param {Array<Array<*>>} ledgerRows - values from A7 onward
 * @returns {boolean}
 */
function _hasUserLedgerData(ledgerRows) {
  const sampleNames = DEFAULT_PORTFOLIO_DATA.map(r => String(r[0]));
  for (let i = 0; i < (ledgerRows || []).length; i++) {
    const name = String(ledgerRows[i][0] || "").trim();
    if (!name) continue;
    if (sampleNames.indexOf(name) === -1) return true;
  }
  return false;
}

/**
 * Requires the user to type a confirmation phrase before a destructive rebuild.
 * @param {string} phrase
 * @param {string} warning
 * @returns {boolean} whether the user confirmed
 */
function _confirmDestructive(phrase, warning) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "⚠️ Destructive operation",
    `${warning}\n\nThis CANNOT be undone from the menu. Recover via File > Version history if needed.\n\nType ${phrase} to proceed:`,
    ui.ButtonSet.OK_CANCEL
  );
  return res.getSelectedButton() === ui.Button.OK &&
         String(res.getResponseText()).trim().toUpperCase() === phrase;
}

/** Menu action: full rebuild, behind a typed confirmation. */
function rebuildEverythingDestructive() {
  if (!_confirmDestructive("ERASE",
      "This ERASES every tab and replaces your ledger with the sample portfolio.")) {
    SpreadsheetApp.getUi().alert("Cancelled — nothing was changed.");
    return;
  }
  runFirstTimeSetup(null, false, true);
}

/** Menu action: rebuild only the Cash Flow tab, behind a typed confirmation. */
function rebuildCashFlowDestructive() {
  if (!_confirmDestructive("ERASE",
      "This ERASES your entire expense ledger on the Cash Flow & Burn tab.")) {
    SpreadsheetApp.getUi().alert("Cancelled — nothing was changed.");
    return;
  }
  buildCashFlowTab();
}
