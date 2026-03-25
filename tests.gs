// tests.gs — Baseline Unit Tests for WealthScript
// Run any individual test function from the Apps Script editor.
// These tests are ISOLATED: no live SpreadsheetApp or UrlFetchApp calls.

const Assert = {
  /**
   * Asserts strict equality between actual and expected values.
   * @param {*} actual
   * @param {*} expected
   * @param {string} testName
   */
  equal: (actual, expected, testName) => {
    if (actual !== expected) throw new Error(`[FAIL] ${testName}: Expected ${expected}, got ${actual}`);
    Logger.log(`[PASS] ${testName}`);
  },

  /**
   * Asserts that a condition is truthy.
   * @param {boolean} condition
   * @param {string} testName
   */
  isTrue: (condition, testName) => {
    if (!condition) throw new Error(`[FAIL] ${testName}: Condition was false`);
    Logger.log(`[PASS] ${testName}`);
  },

  /**
   * Asserts that two floats are approximately equal within a delta tolerance.
   * @param {number} actual
   * @param {number} expected
   * @param {number} delta
   * @param {string} testName
   */
  closeTo: (actual, expected, delta, testName) => {
    if (Math.abs(actual - expected) > delta) throw new Error(`[FAIL] ${testName}: Expected ~${expected}, got ${actual}`);
    Logger.log(`[PASS] ${testName}`);
  }
};

// =============================================================================
// HELPERS UNDER TEST
// Mirror the pure business logic from captureSnapshot() so tests are
// fully isolated from any Sheets or network I/O.
// =============================================================================

/**
 * Pure function: calculates the net worth delta and percentage growth.
 * @param {number} currentNet
 * @param {number|null} prevNet
 * @returns {{ dollarDelta: number|string, pctGrowth: number|string }}
 */
const _calcGrowthDelta = (currentNet, prevNet) => {
  if (!prevNet || isNaN(prevNet)) return { dollarDelta: '', pctGrowth: '' };
  const dollarDelta = currentNet - prevNet;
  return { dollarDelta, pctGrowth: dollarDelta / prevNet };
};

/**
 * Pure function: calculates FIRE progress as a ratio of current net worth to target.
 * @param {number} netWorth
 * @param {number} fireTarget
 * @returns {number}
 */
const _calcFireProgress = (netWorth, fireTarget) => netWorth / fireTarget;

/**
 * Pure function: classifies a ledger row into liquid, locked, or skip.
 * @param {string} assetClass
 * @param {string} status
 * @param {number} netVal
 * @returns {'liquid'|'locked'|'skip'}
 */
const _classifyAsset = (assetClass, status, netVal) => {
  if (status !== 'Active' || isNaN(netVal) || netVal === 0) return 'skip';
  return ['Cash', 'Brokerage', 'Crypto', 'Receivable'].includes(assetClass) ? 'liquid' : 'locked';
};

// =============================================================================
// TEST SUITES
// =============================================================================

/** @description Tests for the net worth delta calculation helper. */
function test_calcGrowthDelta() {
  let result;

  result = _calcGrowthDelta(1100000, 1000000);
  Assert.equal(result.dollarDelta, 100000, 'delta: positive growth dollar amount');
  Assert.closeTo(result.pctGrowth, 0.1, 0.0001, 'delta: positive growth pct');

  result = _calcGrowthDelta(950000, 1000000);
  Assert.equal(result.dollarDelta, -50000, 'delta: negative growth dollar amount');
  Assert.closeTo(result.pctGrowth, -0.05, 0.0001, 'delta: negative growth pct');

  result = _calcGrowthDelta(500000, null);
  Assert.equal(result.dollarDelta, '', 'delta: no prev snapshot returns empty string');
  Assert.equal(result.pctGrowth, '', 'delta: no prev snapshot pct returns empty string');

  result = _calcGrowthDelta(500000, NaN);
  Assert.equal(result.dollarDelta, '', 'delta: NaN prevNet returns empty string');
}

/** @description Tests for the FIRE progress ratio helper. */
function test_calcFireProgress() {
  Assert.closeTo(_calcFireProgress(3000000, 3000000), 1.0, 0.0001, 'FIRE: 100% at target');
  Assert.closeTo(_calcFireProgress(1500000, 3000000), 0.5, 0.0001, 'FIRE: 50% at half target');
  Assert.closeTo(_calcFireProgress(0, 3000000), 0.0, 0.0001, 'FIRE: 0% at zero net worth');
  Assert.isTrue(_calcFireProgress(4000000, 3000000) > 1, 'FIRE: exceeding target returns > 1');
}

/** @description Tests for the asset liquidity classification helper. */
function test_classifyAsset() {
  Assert.equal(_classifyAsset('Cash', 'Active', 10000), 'liquid', 'classify: Cash is liquid');
  Assert.equal(_classifyAsset('Brokerage', 'Active', 50000), 'liquid', 'classify: Brokerage is liquid');
  Assert.equal(_classifyAsset('Crypto', 'Active', 5000), 'liquid', 'classify: Crypto is liquid');
  Assert.equal(_classifyAsset('Receivable', 'Active', 2000), 'liquid', 'classify: Receivable is liquid');
  Assert.equal(_classifyAsset('Real Estate', 'Active', 400000), 'locked', 'classify: Real Estate is locked');
  Assert.equal(_classifyAsset('Retirement', 'Active', 120000), 'locked', 'classify: Retirement is locked');
  Assert.equal(_classifyAsset('Cash', 'Inactive', 5000), 'skip', 'classify: Inactive row is skipped');
  Assert.equal(_classifyAsset('Cash', 'Active', 0), 'skip', 'classify: zero-value row is skipped');
  Assert.equal(_classifyAsset('Cash', 'Active', NaN), 'skip', 'classify: NaN value row is skipped');
}

/** @description Master runner — executes all test suites sequentially. */
function runAllTests() {
  Logger.log('=== Running WealthScript Test Suite ===');
  test_calcGrowthDelta();
  test_calcFireProgress();
  test_classifyAsset();
  test_cashFlowKpis();
  test_buildLedgerSnapshot();
  test_driveBackupPruning();
  Logger.log('=== All tests passed ✅ ===');
}

// =============================================================================
// _buildLedgerSnapshot — SHARED BACKUP HELPER UNDER TEST
// =============================================================================

/**
 * Mirrors the pure transformation logic of _buildLedgerSnapshot().
 * @param {Array<Array<*>>} dataRange
 * @returns {Array<Object>}
 */
const _buildLedgerSnapshot_test = (dataRange) => {
  const snapshot = [];
  for (let i = 0; i < dataRange.length; i++) {
    const account = String(dataRange[i][0]);
    if (!account || account === '') continue;
    snapshot.push({
      Account:        account,
      AssetClass:     String(dataRange[i][1]),
      Currency:       String(dataRange[i][2]),
      InitialCapital: Number(dataRange[i][3]) || 0,
      CurrentValue:   Number(dataRange[i][4]) || 0,
      ExchangeRate:   Number(dataRange[i][5]) || 0,
      TaxRate:        Number(dataRange[i][6]) || 0,
      GrossWorthUSD:  Number(dataRange[i][7]) || 0,
      NetWorthUSD:    Number(dataRange[i][8]) || 0,
      Status:  String(dataRange[i][9]),
      Remarks: String(dataRange[i][10]),
    });
  }
  return snapshot;
};

/** @description Tests for the shared _buildLedgerSnapshot helper. */
function test_buildLedgerSnapshot() {
  // Standard valid row
  const mockData = [
    ["Primary Checking", "Cash", "USD", 0, 8500, 1, 0, 8500, 8500, "Active", "Emergency fund"],
    ["Fidelity 401k",    "Retirement", "USD", 0, 609873, 1, 0.15, 609873, 518392, "Active", "Pre-tax"],
    ["",                 "",    "",    0, 0,    0, 0,    0,      0,      "",       ""],  // blank row — skip
    ["AMEX CA Card",     "Liability", "CAD", 0, -1442, 0.73, 0, -1052, -1052, "Active", ""],
  ];

  const result = _buildLedgerSnapshot_test(mockData);

  // Only non-blank rows should be included
  Assert.equal(result.length, 3, 'snapshot: skips blank account rows');
  Assert.equal(result[0].Account, 'Primary Checking', 'snapshot: preserves account name');
  Assert.equal(result[0].AssetClass, 'Cash', 'snapshot: preserves asset class');
  Assert.equal(result[0].CurrentValue, 8500, 'snapshot: preserves current value');
  Assert.equal(result[1].TaxRate, 0.15, 'snapshot: preserves tax rate');
  Assert.equal(result[2].CurrentValue, -1442, 'snapshot: preserves negative liability value');
  Assert.equal(result[2].Currency, 'CAD', 'snapshot: preserves non-USD currency');

  // Empty ledger
  Assert.equal(_buildLedgerSnapshot_test([]).length, 0, 'snapshot: empty ledger returns empty array');

  // All-blank rows
  const allBlanks = [["","","",0,0,0,0,0,0,"",""],["","","",0,0,0,0,0,0,"",""]];
  Assert.equal(_buildLedgerSnapshot_test(allBlanks).length, 0, 'snapshot: all-blank ledger returns empty array');
}

/**
 * Pure function: mirrors the pruning decision from backupToGoogleDrive().
 * Returns how many files would be trashed given fileCount and limit.
 * @param {number} fileCount
 * @param {number} maxCount
 * @returns {number}
 */
const _countFilesToPrune = (fileCount, maxCount) => Math.max(0, fileCount - maxCount);

/** @description Tests for the Drive backup auto-pruning logic. */
function test_driveBackupPruning() {
  Assert.equal(_countFilesToPrune(10, 24), 0,  'prune: under limit, nothing trashed');
  Assert.equal(_countFilesToPrune(24, 24), 0,  'prune: at exact limit, nothing trashed');
  Assert.equal(_countFilesToPrune(25, 24), 1,  'prune: one over limit, one trashed');
  Assert.equal(_countFilesToPrune(30, 24), 6,  'prune: 6 over limit, 6 trashed');
  Assert.equal(_countFilesToPrune(0,  24), 0,  'prune: empty folder, nothing trashed');
}


// =============================================================================
// EPIC 1 — CASH FLOW KPI HELPERS UNDER TEST
// Mirror the pure math from buildCashFlowTab() KPI formulas.
// =============================================================================

/**
 * Pure function: calculates the annualised Safe Withdrawal Rate.
 * Formula mirrors: (avgMonthlyBurn * 12) / currentNetWorth
 * @param {number} avgMonthlyBurn
 * @param {number} netWorth
 * @returns {number} SWR as a decimal (e.g. 0.04 = 4%)
 */
const _calcSafeWithdrawalRate = (avgMonthlyBurn, netWorth) => {
  if (!netWorth || netWorth === 0) return 0;
  return (avgMonthlyBurn * 12) / netWorth;
};

/**
 * Pure function: sums expenses from a static mock ledger that fall within TTM window.
 * @param {Array<{date: Date, amount: number}>} rows
 * @param {Date} today
 * @returns {number}
 */
const _calcTtmExpenses = (rows, today) => {
  const cutoff = new Date(today);
  cutoff.setDate(cutoff.getDate() - 365);
  return rows
    .filter(r => r.date >= cutoff && r.amount > 0)
    .reduce((sum, r) => sum + r.amount, 0);
};

/** @description Tests for Cash Flow KPI calculation helpers (Epic 1). */
function test_cashFlowKpis() {
  // Safe Withdrawal Rate
  Assert.closeTo(_calcSafeWithdrawalRate(5000, 1500000), 0.04, 0.0001, 'SWR: classic 4% rule');
  Assert.closeTo(_calcSafeWithdrawalRate(20000, 3000000), 0.08, 0.0001, 'SWR: aggressive spend scenario');
  Assert.equal(_calcSafeWithdrawalRate(5000, 0), 0, 'SWR: zero net worth returns 0 (no divide-by-zero)');

  // TTM expense aggregation
  const today = new Date('2026-03-24');
  const mockRows = [
    { date: new Date('2026-03-01'), amount: 3500 },  // within TTM
    { date: new Date('2025-04-01'), amount: 800 },   // within TTM (just inside)
    { date: new Date('2025-03-01'), amount: 999 },   // outside TTM (>365 days ago)
    { date: new Date('2026-02-01'), amount: -100 },  // negative — should be excluded
  ];
  Assert.closeTo(_calcTtmExpenses(mockRows, today), 4300, 0.01, 'TTM: sums only positive entries within 365 days');
  Assert.equal(_calcTtmExpenses([], today), 0, 'TTM: empty ledger returns 0');
}

