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
  //          A account                B class      ... E value (idx 4) ... J status (idx 9)
  const mk = (a, c, e, st) => [a, c, "", "", e, "", "", "", "", st];
  const rows = [
    mk("IBKR",        "Brokerage", "",     "Active"),   // row 7  - linked
    mk("Schwab",      "Brokerage", 0,      "Active"),   // row 8  - frozen 0 on a LIVE account
    mk("Fidelity",    "Brokerage", 48250,  "Active"),   // row 9  - real manual number
    mk("Wealthfront", "Brokerage", 0,      "Active"),   // row 10 - no formula, HAS holdings
    mk("Chase",       "Cash",      12000,  "Active")    // row 11 - not holdings-linked
  ];
  const holdings = { "IBKR": 25, "Wealthfront": 8, "TD Ameritrade": 4 };
  const formulas = ["=SUMPRODUCT(1)", "", "", "", ""];

  const r = _classifyBrokerageRows(rows, holdings, formulas, ["Brokerage"]);

  Assert.equal(r.broken.length, 1, 'classify: only the row with holdings counts as broken');
  Assert.equal(r.broken[0].account, "Wealthfront", 'classify: names the genuinely broken account');
  Assert.equal(r.broken[0].holdings, 8, 'classify: reports how many holdings rows are waiting');

  Assert.equal(r.suspect.length, 1, 'classify: a frozen 0 is damage, not a manual balance');
  Assert.equal(r.suspect[0].account, "Schwab", 'classify: names the frozen-zero row');
  Assert.equal(r.suspect[0].status, "Active", 'classify: status carried through for review');

  Assert.equal(r.unlinked.length, 1, 'classify: a real number is an intentional manual balance');
  Assert.equal(r.unlinked[0].account, "Fidelity", 'classify: manual balance left alone');

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

function test_classifyAcrossAssetClasses() {
  const mk = (a, c, e, st) => [a, c, "", "", e, "", "", "", "", st];
  const rows = [
    mk("IBKR",     "Brokerage",  "",     "Active"),   // 7  linked
    mk("Vanguard", "Retirement", 0,      "Active"),   // 8  retirement, HAS holdings -> broken
    mk("Coinbase", "Crypto",     0,      "Active"),   // 9  crypto, HAS holdings -> broken
    mk("Schwab",   "Brokerage",  0,      "Active"),   // 10 frozen 0 on a live account -> suspect
    mk("Fidelity", "Brokerage",  48250,  "Active"),   // 11 real manual number -> unlinked
    mk("Keybank",  "Cash",       9000,   "Inactive")  // 12 cash, never linked -> ignored
  ];
  const holdings = { "IBKR": 25, "Vanguard": 12, "Coinbase": 3, "TD Ameritrade": 4 };
  const formulas = ["=SUMPRODUCT(1)", "", "", "", "", ""];

  const r = _classifyBrokerageRows(rows, holdings, formulas, ["Brokerage", "Retirement", "Crypto"]);

  Assert.equal(r.broken.length, 2, 'classes: Retirement and Crypto rows are monitored, not just Brokerage');
  Assert.isTrue(r.broken.some(b => b.assetClass === "Retirement"), 'classes: Retirement link checked');
  Assert.isTrue(r.broken.some(b => b.assetClass === "Crypto"), 'classes: Crypto link checked');

  Assert.equal(r.suspect.length, 1, 'classes: a literal 0 with no holdings is damage, not intent');
  Assert.equal(r.suspect[0].account, "Schwab", 'classes: names the frozen-zero row');

  Assert.equal(r.unlinked.length, 1, 'classes: a real manual number stays unlinked, not suspect');
  Assert.equal(r.unlinked[0].account, "Fidelity", 'classes: manual balance preserved');

  Assert.isTrue(!r.broken.concat(r.suspect, r.unlinked).some(x => x.account === "Keybank"),
    'classes: Cash rows are never treated as holdings-linked');

  Assert.equal(r.orphaned.length, 1, 'classes: orphaned holdings still detected');
  Assert.equal(r.orphaned[0].account, "TD Ameritrade", 'classes: renamed account surfaced');
}

function test_holdingsPriceFormulaReservesCash() {
  const f = _holdingsRowFormulas(18);
  Assert.isTrue(f.E.indexOf('UPPER(TRIM(C18))="CASH"') > -1, 'holdingsE: CASH is reserved, not sent to GOOGLEFINANCE');
  Assert.isTrue(f.E.indexOf('IFERROR') > -1, 'holdingsE: unquotable tickers degrade to NA() rather than #REF');

  const legacy = _legacyHoldingsPriceFormulas(18);
  Assert.isTrue(legacy.indexOf('=IF(ISBLANK(C18), "", GOOGLEFINANCE(C18, "price"))') > -1,
    'legacy: the original canonical formula is upgradeable');
  Assert.isTrue(legacy.indexOf('=GOOGLEFINANCE("CURRENCY:CADUSD")') === -1,
    'legacy: an FX override is NOT upgradeable and must be preserved');
  Assert.isTrue(legacy.indexOf('=IFERROR(GOOGLEFINANCE(C18,"price"),0.0002)') === -1,
    'legacy: a fallback-price override is NOT upgradeable and must be preserved');
  Assert.isTrue(legacy.indexOf(f.E) === -1, 'legacy: current canonical is not listed as legacy');
}

function test_verifyHeaders() {
  const good = _verifyHeaders(LEDGER_HEADERS.slice(), LEDGER_HEADERS);
  Assert.isTrue(good.ok, 'headers: canonical row passes');

  const shifted = LEDGER_HEADERS.slice();
  shifted[9] = "Net Worth (USD)";                 // J duplicating I
  const bad = _verifyHeaders(shifted, LEDGER_HEADERS);
  Assert.isTrue(!bad.ok, 'headers: a duplicated header is caught');
  Assert.equal(bad.problems.length, 1, 'headers: exactly one problem reported');
  Assert.isTrue(bad.problems[0].indexOf('J should be "Status"') > -1, 'headers: names the column and expected value');

  const inserted = ["Account", "SPURIOUS"].concat(LEDGER_HEADERS.slice(1));
  Assert.isTrue(!_verifyHeaders(inserted, LEDGER_HEADERS).ok, 'headers: an inserted column is caught');

  Assert.isTrue(!_verifyHeaders([], LEDGER_HEADERS).ok, 'headers: an empty row is caught, not treated as valid');
  Assert.isTrue(_verifyHeaders(HOLDINGS_HEADERS.slice(), HOLDINGS_HEADERS).ok, 'headers: holdings row validates too');
}

function test_inactiveZeroIsDormantNotSuspect() {
  const mk = (a, c, e, st) => [a, c, "", "", e, "", "", "", "", st];
  const rows = [
    mk("Schwab",     "Brokerage", 0, "Inactive"),  // row 7  - closed account
    mk("Startengine","Brokerage", 0, "Active")     // row 8  - live account at 0
  ];
  const r = _classifyBrokerageRows(rows, {}, ["", ""], ["Brokerage"]);

  Assert.equal(r.dormant.length, 1, 'dormant: an Inactive zero is not an error');
  Assert.equal(r.dormant[0].account, "Schwab", 'dormant: names the closed account');
  Assert.equal(r.suspect.length, 1, 'dormant: an ACTIVE zero is still flagged as damage');
  Assert.equal(r.suspect[0].account, "Startengine", 'dormant: live account at 0 still surfaces');

  const rep = _formatHealthReport({
    healthy: true, hadBaseline: true, regressions: [], brokenLinks: [], suspect: [],
    unlinked: [], dormant: r.dormant, orphaned: [], headerProblems: []
  });
  Assert.isTrue(rep.indexOf('✅') > -1, 'dormant: does not block a snapshot');
  Assert.isTrue(rep.indexOf('harmless') > -1, 'dormant: explained rather than alarming');
}

function test_headerMismatchBlocksReport() {
  const rep = _formatHealthReport({
    healthy: false, hadBaseline: true, regressions: [], brokenLinks: [], suspect: [],
    unlinked: [], dormant: [], orphaned: [],
    headerProblems: ['Dashboard & Ledger row 6: column J should be "Status" but is "Net Worth (USD)"']
  });
  Assert.isTrue(rep.indexOf('🛑') > -1, 'headerBlock: reported as a hard stop');
  Assert.isTrue(rep.indexOf('BLOCKED') > -1, 'headerBlock: states repair will not run');
  Assert.isTrue(rep.indexOf('wrong column') > -1, 'headerBlock: explains the consequence');
}
