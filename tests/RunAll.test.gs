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
  test_buildAbbrDisplayFormula();
  test_planRealEstateUpdates();
  test_diffFormulaCounts();
  test_hasUserLedgerData();
  test_ledgerFixedFormulas();
  test_formatHealthReport();
  test_healthReportSeparatesErrorsFromNotices();
  test_classifyBrokerageRows();
  test_classifyAcrossAssetClasses();
  test_holdingsPriceFormulaReservesCash();
  test_headerMismatchBlocksReport();
  test_buildOccSymbol();
  test_normaliseFlexPosition();
  test_normaliseFlexCash();
  test_parseFlexEnvelope();
  test_extractFlexElements();
  test_planHoldingsSync();
  test_planHoldingsSyncRunsOutOfRoom();
  test_deriveQuantityFromPositionValue();
  test_planBlockSyncRefusesEmptyFeed();
  test_planBlockSyncRewritesInPlace();
  test_planBlockSyncRefusesToStraddleAnotherAccount();
  test_planBlockSyncFlagsLargeShrink();
  test_planBlockSyncBootstrapsEmptyAccount();
  test_planBlockSyncRelocatesWhenOutgrown();
  test_planBlockSyncRefusesWhenNoRunFits();
  test_planBlockSyncShrinkNeverNegative();
  test_headerAnnotationTolerated();
  test_inactiveZeroIsDormantNotSuspect();
  test_verifyHeaders();
  test_generateInsight();
  test_validatePATFormat();
  test_buildGistUrl();
  test_buildBrokerageFormulaContract();
  test_endToEndIntegration(); // 14-section E2E sandbox (creates + destroys a live spreadsheet)
  Logger.log('=== All tests passed ✅ (100+ assertions, 14 functional suites) ===');
}
