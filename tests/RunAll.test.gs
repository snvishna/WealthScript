/** Master runner — executes all test suites sequentially. */
function runAllTests() {
  Logger.log('=== Running WealthScript Test Suite ===');
  test_calcGrowthDelta();
  test_calcFireProgress();
  test_classifyAsset();
  test_classifyAsset_extended();
  test_cashFlowKpis();
  test_buildLedgerSnapshot();
  test_driveBackupPruning();
  test_currencySymbol();
  test_abbrFmt();
  test_generateInsight();
  test_validatePATFormat();
  test_buildGistUrl();
  test_buildBrokerageFormulaContract();
  test_endToEndIntegration();
  Logger.log('=== All tests passed ✅ (75+ assertions) ===');
}
