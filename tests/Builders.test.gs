function test_buildBrokerageFormulaContract() {
  // Test expected formula injection format for dashboard rows
  
  // Row 7 (first data row)
  const f7 = _buildBrokerageFormula(7);
  Assert.isTrue(f7.includes("SUMIF('Brokerage Holdings'!A:A"), "contract: matches Brokerage name col");
  Assert.isTrue(f7.includes("A7"), "contract: pulls the account name from current dashboard row 7");
  Assert.isTrue(f7.includes("'Brokerage Holdings'!E:E"), "contract: aggregates from Brokerage Total Value col");
  Assert.isTrue(f7.startsWith("=IFERROR"), "contract: handles missing tabs or errors gracefully");
  Assert.equal(f7, "=IFERROR(SUMIF('Brokerage Holdings'!A:A,A7,'Brokerage Holdings'!E:E),0)", "contract: exact string match row 7");

  // Row 50
  const f50 = _buildBrokerageFormula(50);
  Assert.equal(f50, "=IFERROR(SUMIF('Brokerage Holdings'!A:A,A50,'Brokerage Holdings'!E:E),0)", "contract: exact string match row 50");
}
