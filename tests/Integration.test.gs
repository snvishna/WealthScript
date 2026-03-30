/**
 * Native E2E Integration Test Suite
 * Generates an ephemeral Google Spreadsheet, runs the entire
 * setup pipeline, injects mock financial data, executes a snapshot,
 * validates the cross-tab mathematical integrity, and deletes the sheet.
 */

function test_endToEndIntegration() {
  Logger.log("--- Starting E2E Integration Suite ---");
  const testName = "WealthScript_E2E_Test_" + Date.now();
  let testSs;

  try {
    // 1. Create Ephemeral Sandbox
    testSs = SpreadsheetApp.create(testName);
    Logger.log("✅ Sandbox created: " + testSs.getId());
    
    // 2. Headless First Time Setup
    runFirstTimeSetup(testSs, true);
    
    // 3. Tab Architecture Validation
    const dashboard = testSs.getSheetByName("Dashboard & Ledger");
    const holdings = testSs.getSheetByName("Brokerage Holdings");
    const snapshots = testSs.getSheetByName("Snapshots");
    
    Assert.isTrue(dashboard !== null, "E2E: Dashboard tab is built");
    Assert.isTrue(holdings !== null, "E2E: Holdings tab is built");
    Assert.isTrue(snapshots !== null, "E2E: Snapshots tab is built");
    
    // 4. Mock Financial Data Injection
    // Set 'Primary Checking' (Row 7) current value to 100,000
    dashboard.getRange("E7").setValue(100000); // Static entry
    
    // Set Brokerage Holdings for 'Taxable Brokerage'
    // Row 2: Taxable Brokerage | TICKER | 1000 shares | $500/share
    holdings.getRange("A2").setValue("Taxable Brokerage");
    holdings.getRange("C2").setValue(1000);
    // Overwrite the GOOGLEFINANCE formula in D2 so tests don't rely on network latency
    holdings.getRange("D2").setValue(500); 
    // Column E (Total) is C2 * D2 = 500,000. Wait for sheets engine to calculate.
    SpreadsheetApp.flush();
    
    // 5. Cross-Tab Contract Validation
    // 'Taxable Brokerage' is at Row 10 in DEFAULT_PORTFOLIO_DATA
    const brokerageVal = dashboard.getRange("E10").getValue();
    Assert.equal(brokerageVal, 500000, "E2E: Dashboard SUMIF formula successfully aggregates live Brokerage Holdings");
    
    // Total Net Worth (B2) should be 100k + 500k = 600k (plus standard layout data)
    // Actually, DEFAULT_PORTFOLIO_DATA has other accounts with initial non-zero balances.
    // Let's zero out all other balances in E7:E12 to isolate our test math
    const zeroRange = [[100000], [0], [0], [null], [0], [0]]; // Null implies skipping overwrite for row 10 (brokerage)
    dashboard.getRange("E7").setValue(100000);
    dashboard.getRange("E8").setValue(0);
    dashboard.getRange("E9").setValue(0);
    // E10 is Brokerage, driven by formula, do not overwrite
    dashboard.getRange("E11").setValue(0);
    dashboard.getRange("E12").setValue(0);
    
    SpreadsheetApp.flush();
    const netUSD = dashboard.getRange("B2").getValue();
    Assert.equal(netUSD, 600000, "E2E: Dashboard Net Worth KPI accurately rolls up all ledger classes");
    
    // 6. Snapshot Execution
    captureSnapshot(testSs, true);
    
    // 7. Snapshot State Validation
    // Row 2 is the newly inserted snapshot
    const snapNet = snapshots.getRange("B2").getValue();
    const snapLiquid = snapshots.getRange("C2").getValue();
    
    Assert.equal(snapNet, 600000, "E2E: Snapshot correctly persists accurate Net Worth");
    Assert.equal(snapLiquid, 600000, "E2E: Snapshot correctly classifies Checking and Brokerage as Liquid Worth");
    
    Logger.log("✅ E2E Integration Suite Passed");

  } catch (error) {
    Logger.log("❌ E2E Failed: " + error.message);
    throw error;
  } finally {
    // 8. Silent Teardown
    if (testSs) {
      try {
        DriveApp.getFileById(testSs.getId()).setTrashed(true);
        Logger.log("✅ Sandbox deleted securely");
      } catch(e) {
        Logger.log("⚠️ Teardown warning: Could not delete sandbox: " + e.message);
      }
    }
  }
}
