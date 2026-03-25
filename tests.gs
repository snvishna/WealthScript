// tests.gs — Baseline Unit Tests for Net Worth Tracker v1
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
  Logger.log('=== Running Net Worth Tracker Test Suite ===');
  test_calcGrowthDelta();
  test_calcFireProgress();
  test_classifyAsset();
  Logger.log('=== All tests passed ✅ ===');
}
