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

function test_buildLedgerSnapshot() {
  const mockData = [
    ["Primary Checking", "Cash", "USD", 0, 8500, 1, 0, 8500, 8500, "Active", "Emergency fund"],
    ["Fidelity 401k",    "Retirement", "USD", 0, 609873, 1, 0.15, 609873, 518392, "Active", "Pre-tax"],
    ["",                 "",    "",    0, 0,    0, 0,    0,      0,      "",       ""],
    ["AMEX CA Card",     "Liability", "CAD", 0, -1442, 0.73, 0, -1052, -1052, "Active", ""],
  ];

  const result = _buildLedgerSnapshot_test(mockData);

  Assert.equal(result.length, 3, 'snapshot: skips blank account rows');
  Assert.equal(result[0].Account, 'Primary Checking', 'snapshot: preserves account name');
  Assert.equal(result[0].AssetClass, 'Cash', 'snapshot: preserves asset class');
  Assert.equal(result[0].CurrentValue, 8500, 'snapshot: preserves current value');
  Assert.equal(result[1].TaxRate, 0.15, 'snapshot: preserves tax rate');
  Assert.equal(result[2].CurrentValue, -1442, 'snapshot: preserves negative liability value');
  Assert.equal(result[2].Currency, 'CAD', 'snapshot: preserves non-USD currency');

  Assert.equal(_buildLedgerSnapshot_test([]).length, 0, 'snapshot: empty ledger returns empty array');

  const allBlanks = [["","","",0,0,0,0,0,0,"",""],["","","",0,0,0,0,0,0,"",""]];
  Assert.equal(_buildLedgerSnapshot_test(allBlanks).length, 0, 'snapshot: all-blank ledger returns empty array');
}

const _countFilesToPrune = (fileCount, maxCount) => Math.max(0, fileCount - maxCount);

function test_driveBackupPruning() {
  Assert.equal(_countFilesToPrune(10, 24), 0,  'prune: under limit, nothing trashed');
  Assert.equal(_countFilesToPrune(24, 24), 0,  'prune: at exact limit, nothing trashed');
  Assert.equal(_countFilesToPrune(25, 24), 1,  'prune: one over limit, one trashed');
  Assert.equal(_countFilesToPrune(30, 24), 6,  'prune: 6 over limit, 6 trashed');
  Assert.equal(_countFilesToPrune(0,  24), 0,  'prune: empty folder, nothing trashed');
}
