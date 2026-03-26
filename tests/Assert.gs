const Assert = {
  equal: (actual, expected, testName) => {
    if (actual !== expected) throw new Error(`[FAIL] ${testName}: Expected ${expected}, got ${actual}`);
    Logger.log(`[PASS] ${testName}`);
  },
  isTrue: (condition, testName) => {
    if (!condition) throw new Error(`[FAIL] ${testName}: Condition was false`);
    Logger.log(`[PASS] ${testName}`);
  },
  closeTo: (actual, expected, delta, testName) => {
    if (Math.abs(actual - expected) > delta) throw new Error(`[FAIL] ${testName}: Expected ~${expected}, got ${actual}`);
    Logger.log(`[PASS] ${testName}`);
  }
};
