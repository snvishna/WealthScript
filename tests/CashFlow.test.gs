const _calcSafeWithdrawalRate = (avgMonthlyBurn, netWorth) => {
  if (!netWorth || netWorth === 0) return 0;
  return (avgMonthlyBurn * 12) / netWorth;
};

const _calcTtmExpenses = (rows, today) => {
  const cutoff = new Date(today);
  cutoff.setDate(cutoff.getDate() - 365);
  return rows
    .filter(r => r.date >= cutoff && r.amount > 0)
    .reduce((sum, r) => sum + r.amount, 0);
};

function test_cashFlowKpis() {
  Assert.closeTo(_calcSafeWithdrawalRate(5000, 1500000), 0.04, 0.0001, 'SWR: classic 4% rule');
  Assert.closeTo(_calcSafeWithdrawalRate(20000, 3000000), 0.08, 0.0001, 'SWR: aggressive spend scenario');
  Assert.equal(_calcSafeWithdrawalRate(5000, 0), 0, 'SWR: zero net worth returns 0 (no divide-by-zero)');
  Assert.equal(_calcSafeWithdrawalRate(0, 1500000), 0, 'SWR: zero burn returns 0');

  const today = new Date('2026-03-24');
  const mockRows = [
    { date: new Date('2026-03-01'), amount: 3500 },  
    { date: new Date('2025-04-01'), amount: 800 },   
    { date: new Date('2025-03-01'), amount: 999 },   
    { date: new Date('2026-02-01'), amount: -100 },  
  ];
  Assert.closeTo(_calcTtmExpenses(mockRows, today), 4300, 0.01, 'TTM: sums only positive entries within 365 days');
  Assert.equal(_calcTtmExpenses([], today), 0, 'TTM: empty ledger returns 0');
}
