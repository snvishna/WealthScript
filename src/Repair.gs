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
 * Structural check: every row whose Asset Class is "Brokerage" MUST have a
 * formula in its Current Value cell. Needs no baseline — the rule is absolute.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Array<{row: number, account: string}>} broken rows
 */
function _auditBrokerageLinks(ss) {
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) return [];

  const values = sheet.getRange(LEDGER_FIRST_ROW, 1, LEDGER_NUM_ROWS, 2).getValues();
  const formulas = sheet.getRange(LEDGER_FIRST_ROW, 5, LEDGER_NUM_ROWS, 1).getFormulas();

  const broken = [];
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1]).trim() !== "Brokerage") continue;
    if (!formulas[i][0]) {
      broken.push({ row: LEDGER_FIRST_ROW + i, account: String(values[i][0]) });
    }
  }
  return broken;
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
  const brokenLinks = _auditBrokerageLinks(ss);

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(INTEGRITY_PROP_KEY);
  let baseline = null;
  try { baseline = raw ? JSON.parse(raw) : null; } catch (e) { baseline = null; }

  const diff = _diffFormulaCounts(counts, baseline);

  return {
    healthy: diff.ok && brokenLinks.length === 0,
    regressions: diff.regressions,
    brokenLinks: brokenLinks,
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
  if (audit.healthy) {
    return audit.hadBaseline
      ? "✅ All managed formulas are intact."
      : "✅ No damage detected. Baseline recorded — future checks compare against it.";
  }

  const lines = ["⚠️ Formula damage detected.", ""];
  if (audit.regressions.length) {
    lines.push("Formula count dropped in:");
    audit.regressions.forEach(r => lines.push(`  • ${r.label} — was ${r.was}, now ${r.now}`));
    lines.push("");
  }
  if (audit.brokenLinks.length) {
    lines.push("Brokerage rows no longer linked to the Holdings tab:");
    audit.brokenLinks.forEach(b => lines.push(`  • Row ${b.row} — ${b.account}`));
    lines.push("");
  }
  lines.push("Run WealthScript > 🛠 Repair Formulas to restore them.");
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

  // --- 3. Ledger column E: Brokerage rows only. Manual rows are left alone.
  const meta = ledger.getRange(LEDGER_FIRST_ROW, 1, LEDGER_NUM_ROWS, 2).getValues();
  for (let i = 0; i < meta.length; i++) {
    const r = LEDGER_FIRST_ROW + i;
    if (String(meta[i][1]).trim() !== "Brokerage") continue;
    const want = _buildBrokerageFormula(r);
    const cell = ledger.getRange(r, 5);
    if (cell.getFormula() === want) continue;
    if (cell.getFormula() === "" && cell.getValue() !== "") {
      details.push(`Row ${r} (${meta[i][0]}): replaced a frozen literal with the live Holdings link`);
    }
    cell.setFormula(want);
    restored++;
  }

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
