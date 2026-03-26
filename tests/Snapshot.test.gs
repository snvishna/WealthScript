const _calcGrowthDelta = (currentNet, prevNet) => {
  if (!prevNet || isNaN(prevNet)) return { dollarDelta: '', pctGrowth: '' };
  const dollarDelta = currentNet - prevNet;
  return { dollarDelta, pctGrowth: dollarDelta / prevNet };
};

const _calcFireProgress = (netWorth, fireTarget) => netWorth / fireTarget;

const _classifyAsset = (assetClass, status, netVal) => {
  if (status !== 'Active' || isNaN(netVal) || netVal === 0) return 'skip';
  return ['Cash', 'Brokerage', 'Crypto', 'Receivable'].includes(assetClass) ? 'liquid' : 'locked';
};

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

function test_calcFireProgress() {
  Assert.closeTo(_calcFireProgress(3000000, 3000000), 1.0, 0.0001, 'FIRE: 100% at target');
  Assert.closeTo(_calcFireProgress(1500000, 3000000), 0.5, 0.0001, 'FIRE: 50% at half target');
  Assert.closeTo(_calcFireProgress(0, 3000000), 0.0, 0.0001, 'FIRE: 0% at zero net worth');
  Assert.isTrue(_calcFireProgress(4000000, 3000000) > 1, 'FIRE: exceeding target returns > 1');
}

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

function test_classifyAsset_extended() {
  Assert.equal(_classifyAsset('Health Savings', 'Active', 15000), 'locked', 'classify: HSA is locked');
  Assert.equal(_classifyAsset('Insurance', 'Active', 5000), 'locked', 'classify: Insurance is locked');
  Assert.equal(_classifyAsset('Commodity', 'Active', 1000), 'locked', 'classify: Commodity is locked');
  Assert.equal(_classifyAsset('Liability', 'Active', -5000), 'locked', 'classify: Liability is locked');
  Assert.equal(_classifyAsset('Cash', 'Active', -100), 'liquid', 'classify: negative Cash is still liquid');
}
