function test_buildBrokerageFormulaContract() {
  const f7 = _buildBrokerageFormula(7);
  Assert.isTrue(f7.includes("'Brokerage Holdings'!$A$2:$A$200"), "contract: matches Brokerage account name range");
  Assert.isTrue(f7.includes("A7"), "contract: pulls account name from dashboard row 7");
  Assert.isTrue(f7.includes("N('Brokerage Holdings'!$E$2:$E$200)"), "contract: N() coerces empty strings to 0, preventing #VALUE!");
  Assert.isTrue(f7.includes("SUMPRODUCT"), "contract: uses SUMPRODUCT for cross-sheet volatile formula support");
  Assert.isTrue(!f7.includes("IFERROR"), "contract: no IFERROR wrapper (would silently swallow errors as 0)");
  Assert.equal(f7, "=SUMPRODUCT(('Brokerage Holdings'!$A$2:$A$200=A7)*N('Brokerage Holdings'!$E$2:$E$200))", "contract: exact string match row 7");

  const f50 = _buildBrokerageFormula(50);
  Assert.equal(f50, "=SUMPRODUCT(('Brokerage Holdings'!$A$2:$A$200=A50)*N('Brokerage Holdings'!$E$2:$E$200))", "contract: exact string match row 50");
}
