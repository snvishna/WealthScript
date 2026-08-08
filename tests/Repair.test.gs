/**
 * Tests for the non-destructive repair, migration and integrity layer.
 * All pure — no Sheets calls.
 */

function test_diffFormulaCounts() {
  const base = { "Exchange Rate column": 70, "KPI totals (USD)": 3, "Holdings Total Value column": 99 };

  const clean = _diffFormulaCounts({ "Exchange Rate column": 70, "KPI totals (USD)": 3, "Holdings Total Value column": 99 }, base);
  Assert.isTrue(clean.ok, 'diffCounts: identical counts report healthy');

  const grown = _diffFormulaCounts({ "Exchange Rate column": 90, "KPI totals (USD)": 3, "Holdings Total Value column": 99 }, base);
  Assert.isTrue(grown.ok, 'diffCounts: growth is not a regression (user added accounts)');

  const damaged = _diffFormulaCounts({ "Exchange Rate column": 0, "KPI totals (USD)": 3, "Holdings Total Value column": 99 }, base);
  Assert.isTrue(!damaged.ok, 'diffCounts: a drop is flagged');
  Assert.equal(damaged.regressions.length, 1, 'diffCounts: exactly one range regressed');
  Assert.equal(damaged.regressions[0].label, "Exchange Rate column", 'diffCounts: names the damaged range');
  Assert.equal(damaged.regressions[0].was, 70, 'diffCounts: reports the previous count');
  Assert.equal(damaged.regressions[0].now, 0, 'diffCounts: reports the current count');

  const noBaseline = _diffFormulaCounts({ "Exchange Rate column": 70 }, null);
  Assert.isTrue(noBaseline.ok, 'diffCounts: absent baseline is healthy, not a false alarm');
}

function test_hasUserLedgerData() {
  const sampleOnly = DEFAULT_PORTFOLIO_DATA.slice(0, 5).map(r => [r[0]]);
  Assert.isTrue(!_hasUserLedgerData(sampleOnly), 'userData: shipped sample rows are not user data');

  Assert.isTrue(!_hasUserLedgerData([[""], [""], [null]]), 'userData: blank ledger is not user data');

  const withReal = sampleOnly.concat([["IBKR"]]);
  Assert.isTrue(_hasUserLedgerData(withReal), 'userData: a real account is detected');

  Assert.isTrue(!_hasUserLedgerData([]), 'userData: empty array is safe');
}

function test_ledgerFixedFormulas() {
  const f = _ledgerFixedFormulas();

  Assert.isTrue(f.I4.indexOf("'Settings & Config'!$B$22") > -1, 'fixed: FIRE progress reads the settings cell, not a hardcoded target');
  Assert.isTrue(f.H4.indexOf("'Settings & Config'!$B$22") > -1, 'fixed: FIRE label reads the settings cell');
  Assert.isTrue(f.I4.indexOf("3000000") === -1, 'fixed: no hardcoded 3M remains');

  Assert.equal(f.N2, '=B2*$M$2', 'fixed: card 2 numeric helper');
  Assert.equal(f.O2, '=B2*$M$3', 'fixed: card 3 numeric helper');
  Assert.isTrue(f.E2.indexOf('$N$2') > -1, 'fixed: card 2 display reads its helper cell');
  Assert.isTrue(f.I2.indexOf('$O$2') > -1, 'fixed: card 3 display reads its helper cell');

  Assert.isTrue(f.B4.indexOf('"Cash"') > -1 && f.B4.indexOf('"Receivable"') > -1, 'fixed: liquid total covers all liquid classes');
}

function test_formatHealthReport() {
  const ok = _formatHealthReport({ healthy: true, hadBaseline: true, regressions: [], brokenLinks: [] });
  Assert.isTrue(ok.indexOf('✅') > -1, 'report: healthy result is affirmative');

  const bad = _formatHealthReport({
    healthy: false, hadBaseline: true,
    regressions: [{ label: 'Exchange Rate column', was: 70, now: 0 }],
    brokenLinks: [{ row: 12, account: 'IBKR' }]
  });
  Assert.isTrue(bad.indexOf('Exchange Rate column') > -1, 'report: names the regressed range');
  Assert.isTrue(bad.indexOf('IBKR') > -1, 'report: names the unlinked brokerage account');
  Assert.isTrue(bad.indexOf('Repair Formulas') > -1, 'report: tells the user what to do next');
}

function test_classifyBrokerageRows() {
  //          A account                B class        ... J status (index 9)
  const mk = (a, c, st) => [a, c, "", "", "", "", "", "", "", st];
  const rows = [
    mk("IBKR",     "Brokerage", "Active"),    // row 7  - linked
    mk("Schwab",   "Brokerage", "Inactive"),  // row 8  - no formula, no holdings
    mk("Fidelity", "Brokerage", "Active"),    // row 9  - no formula, no holdings
    mk("Wealthfront","Brokerage","Active"),   // row 10 - no formula, HAS holdings
    mk("Chase",    "Cash",      "Active")     // row 11 - not brokerage
  ];
  const holdings = { "IBKR": 25, "Wealthfront": 8, "TD Ameritrade": 4 };
  const formulas = ["=SUMPRODUCT(1)", "", "", "", ""];

  const r = _classifyBrokerageRows(rows, holdings, formulas);

  Assert.equal(r.broken.length, 1, 'classify: only the row with holdings counts as broken');
  Assert.equal(r.broken[0].account, "Wealthfront", 'classify: names the genuinely broken account');
  Assert.equal(r.broken[0].holdings, 8, 'classify: reports how many holdings rows are waiting');

  Assert.equal(r.unlinked.length, 2, 'classify: manual-value brokerages are unlinked, not broken');
  Assert.isTrue(r.unlinked.some(u => u.account === "Fidelity"), 'classify: Active manual account is not an error');
  Assert.isTrue(r.unlinked.some(u => u.status === "Inactive"), 'classify: status carried through for review');

  Assert.equal(r.orphaned.length, 1, 'classify: holdings with no ledger row are flagged');
  Assert.equal(r.orphaned[0].account, "TD Ameritrade", 'classify: catches a renamed account');
  Assert.equal(r.orphaned[0].rows, 4, 'classify: reports orphaned holdings row count');
}

function test_healthReportSeparatesErrorsFromNotices() {
  const rep = _formatHealthReport({
    healthy: true, hadBaseline: true, regressions: [], brokenLinks: [],
    unlinked: [{ row: 27, account: "Fidelity", status: "Active" }],
    orphaned: [{ account: "TD Ameritrade", rows: 4 }]
  });
  Assert.isTrue(rep.indexOf('✅') > -1, 'report: notices alone do not make the sheet unhealthy');
  Assert.isTrue(rep.indexOf('Fidelity') > -1, 'report: unlinked account surfaced as a notice');
  Assert.isTrue(rep.indexOf('not counted in your net worth') > -1, 'report: orphaned holdings explained in impact terms');
}
