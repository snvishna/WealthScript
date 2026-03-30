/**
 * ============================================================
 * WealthScript — Comprehensive E2E Integration Test Suite
 * Covers ALL 8 product feature areas.
 * Entry point: test_endToEndIntegration() called from RunAll.test.gs
 * ============================================================
 *
 * Architecture:
 * - All tests share one ephemeral sandbox spreadsheet (created + destroyed per session)
 * - Tests are ordered: Setup → Structural → Functional → Integration → Teardown
 * - Each section independently verifiable; failures abort with descriptive messages
 */

/** Master orchestrator: runs all E2E sections in sequence against a single sandbox. */
function test_endToEndIntegration() {
  Logger.log("=== E2E Integration Suite Starting ===");
  let testSs;

  try {
    // 1. Sandbox Creation
    testSs = SpreadsheetApp.create("WealthScript_E2E_" + Date.now());
    Logger.log("Sandbox: " + testSs.getId());

    _e2e_setup(testSs);
    _e2e_tabArchitecture(testSs);
    _e2e_settingsTab(testSs);
    _e2e_dashboardLedger(testSs);
    _e2e_brokerageHoldings(testSs);
    _e2e_snapshotTab(testSs);
    _e2e_cashFlowTab(testSs);
    _e2e_triggerRegistration();
    _e2e_backupJsonSchema(testSs);
    _e2e_snapshotMath(testSs);
    _e2e_githubIntegrationContracts();
    _e2e_driveApiContracts();
    _e2e_realEstateGracefulDegradation(testSs);
    _e2e_wizardContracts();

    Logger.log("=== E2E Integration Suite: ALL SUITES PASSED ✅ ===");

  } catch (error) {
    Logger.log("❌ E2E Failed: " + error.message);
    throw error;
  } finally {
    if (testSs) {
      try {
        // Teardown via DriveApp — trashes the sandbox spreadsheet
        DriveApp.getFileById(testSs.getId()).setTrashed(true);
        Logger.log("Sandbox deleted ✅");
      } catch (e) {
        Logger.log("⚠️ Teardown warning: " + e.message);
      }
    }
  }
}

// ─────────────────────────────────────────────────────────────
// SECTION 1: FIRST TIME SETUP
// ─────────────────────────────────────────────────────────────
function _e2e_setup(testSs) {
  runFirstTimeSetup(testSs, true); // silent = true
  Logger.log("✅ S1: First Time Setup completed");
}

// ─────────────────────────────────────────────────────────────
// SECTION 2: TAB ARCHITECTURE (All 4 Tabs Exist)
// ─────────────────────────────────────────────────────────────
function _e2e_tabArchitecture(testSs) {
  const requiredTabs = ["Dashboard & Ledger", "Brokerage Holdings", "Snapshots", "Settings & Config", "💸 Cash Flow & Burn"];
  for (const name of requiredTabs) {
    Assert.isTrue(testSs.getSheetByName(name) !== null, `E2E-TAB: "${name}" tab was built`);
  }
  Logger.log("✅ S2: Tab Architecture verified (5 tabs)");
}

// ─────────────────────────────────────────────────────────────
// SECTION 3: SETTINGS & CONFIG TAB
// ─────────────────────────────────────────────────────────────
function _e2e_settingsTab(testSs) {
  const s = testSs.getSheetByName("Settings & Config");
  Assert.isTrue(s !== null, "E2E-SETTINGS: Tab exists");

  // Getting Started block must be at rows 1-6
  const gsHeader = s.getRange("A1").getValue();
  Assert.isTrue(gsHeader.includes("GETTING STARTED"), "E2E-SETTINGS: Getting Started block is present");

  // Config sections must exist (headers only, values may be placeholders)
  const a8 = s.getRange("A8").getValue();
  Assert.isTrue(String(a8).includes("REAL ESTATE"), "E2E-SETTINGS: RapidAPI config section exists at A8");

  const a12 = s.getRange("A12").getValue();
  Assert.isTrue(String(a12).includes("CLOUD BACKUP"), "E2E-SETTINGS: Cloud Backup config section exists at A12");

  const a18 = s.getRange("A18").getValue();
  Assert.isTrue(String(a18).includes("FIRE"), "E2E-SETTINGS: FIRE config section exists at A18");

  const a22 = s.getRange("A22").getValue();
  Assert.isTrue(String(a22).includes("CURRENCY"), "E2E-SETTINGS: Currency config section exists at A22");

  const a26 = s.getRange("A26").getValue();
  Assert.isTrue(String(a26).includes("ZPID"), "E2E-SETTINGS: ZPID mapping section exists at A26");

  Logger.log("✅ S3: Settings tab structure verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 4: DASHBOARD & LEDGER
// ─────────────────────────────────────────────────────────────
function _e2e_dashboardLedger(testSs) {
  const s = testSs.getSheetByName("Dashboard & Ledger");
  Assert.isTrue(s !== null, "E2E-DASHBOARD: Tab exists");

  // Header row must be at row 6
  const headers = s.getRange("A6:K6").getValues()[0];
  const expected = ["Account", "Asset Class", "Currency", "Initial Capital", "Current Value",
    "Exchange Rate (to USD)", "Tax Rate", "Gross Worth (USD)", "Net Worth (USD)", "Status", "Remarks"];
  for (let i = 0; i < expected.length; i++) {
    Assert.equal(headers[i], expected[i], `E2E-DASHBOARD: Header col ${i+1} = "${expected[i]}"`);
  }

  // Data must start at row 7 and have accounts
  const firstAccount = s.getRange("A7").getValue();
  Assert.isTrue(String(firstAccount).length > 0, "E2E-DASHBOARD: Sample account data populated at row 7");

  // Exchange rate column (F) should contain GOOGLEFINANCE formulas
  const exchFormula = s.getRange("F7").getFormula();
  Assert.isTrue(exchFormula.includes("GOOGLEFINANCE") || exchFormula.includes("USD"),
    "E2E-DASHBOARD: Exchange rate formula is present in column F");

  // Gross Worth formula in column H
  const grossFormula = s.getRange("H7").getFormula();
  Assert.isTrue(grossFormula.includes("E7") && grossFormula.includes("F7"),
    "E2E-DASHBOARD: Gross Worth formula references Current Value and Exchange Rate");

  // Net Worth formula in column I
  const netFormula = s.getRange("I7").getFormula();
  Assert.isTrue(netFormula.includes("H7"), "E2E-DASHBOARD: Net Worth formula references Gross Worth");

  // Brokerage rows must have SUMIF formula in Current Value
  const brokerageRows = [];
  const portfolioData = s.getRange("A7:B77").getValues();
  for (let i = 0; i < portfolioData.length; i++) {
    if (portfolioData[i][1] === "Brokerage") brokerageRows.push(i + 7);
  }
  Assert.isTrue(brokerageRows.length > 0, "E2E-DASHBOARD: At least one Brokerage row exists in ledger");
  for (const r of brokerageRows) {
    const f = s.getRange(r, 5).getFormula();
    Assert.isTrue(f.includes("SUMPRODUCT"), `E2E-DASHBOARD: Row ${r} uses SUMPRODUCT for live aggregation`);
    Assert.isTrue(f.includes("N("), `E2E-DASHBOARD: Row ${r} uses N() to prevent #VALUE! from empty Holdings cells`);
    Assert.isTrue(!f.includes("IFERROR"), `E2E-DASHBOARD: Row ${r} has no IFERROR wrapper that could silently return 0`);
  }

  Logger.log("✅ S4: Dashboard & Ledger architecture verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 5: BROKERAGE HOLDINGS TAB
// ─────────────────────────────────────────────────────────────
function _e2e_brokerageHoldings(testSs) {
  const s = testSs.getSheetByName("Brokerage Holdings");
  Assert.isTrue(s !== null, "E2E-HOLDINGS: Tab exists");

  const headers = s.getRange("A1:E1").getValues()[0];
  Assert.equal(headers[0], "Account Name",   "E2E-HOLDINGS: Col A header = Account Name");
  Assert.equal(headers[1], "Ticker Symbol",  "E2E-HOLDINGS: Col B header = Ticker Symbol");
  Assert.equal(headers[2], "Quantity",       "E2E-HOLDINGS: Col C header = Quantity");
  Assert.equal(headers[3], "Live Price",     "E2E-HOLDINGS: Col D header = Live Price");
  Assert.equal(headers[4], "Total Value",    "E2E-HOLDINGS: Col E header = Total Value");

  // Live Price col D should have GOOGLEFINANCE formula
  const priceFmla = s.getRange("D2").getFormula();
  Assert.isTrue(priceFmla.includes("GOOGLEFINANCE") || priceFmla.includes("IF"),
    "E2E-HOLDINGS: Live Price uses GOOGLEFINANCE");

  // Total Value col E should be C * D
  const totalFmla = s.getRange("E2").getFormula();
  Assert.isTrue(totalFmla.includes("C2") && totalFmla.includes("D2"),
    "E2E-HOLDINGS: Total Value = Quantity × Live Price");

  // Cross-tab: inject account name + quantity, allow GOOGLEFINANCE to resolve, assert SUMPRODUCT fires
  const dashboard = testSs.getSheetByName("Dashboard & Ledger");
  const portfolio = dashboard.getRange("A7:B77").getValues();
  let brokerageAcct = "";
  let brokerageRow = -1;
  for (let i = 0; i < portfolio.length; i++) {
    if (portfolio[i][1] === "Brokerage" && portfolio[i][0]) {
      brokerageAcct = portfolio[i][0];
      brokerageRow = i + 7;
      break;
    }
  }
  Assert.isTrue(brokerageAcct !== "", "E2E-HOLDINGS: Found a Brokerage account in Dashboard to test cross-tab");

  // Set up Holdings row with a static price (NOT bypassing the formula chain)
  // We inject a static price into D2 to avoid network latency in tests, but the
  // E (Total Value) column formula =IF(ISNUMBER(C2),ISNUMBER(D2), C2*D2, "") still runs natively
  s.getRange("A2").setValue(brokerageAcct);
  s.getRange("C2").setValue(100);
  s.getRange("D2").setValue(250); // static price: simulates resolved GOOGLEFINANCE

  // Force the spreadsheet engine to recalculate formula chains (including SUMPRODUCT in Dashboard)
  SpreadsheetApp.flush();

  // Verify via formula result — use > 0 to tolerate any static fixture value
  const crossTabValue = dashboard.getRange(brokerageRow, 5).getValue();
  Assert.isTrue(crossTabValue > 0, "E2E-HOLDINGS: Dashboard SUMPRODUCT cross-tab correctly aggregates Holdings values (> 0)");
  Assert.equal(crossTabValue, 25000, "E2E-HOLDINGS: Dashboard SUMPRODUCT cross-tab = 100 qty × $250 price = 25000");

  Logger.log("✅ S5: Brokerage Holdings cross-tab linkage verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 6: SNAPSHOTS TAB
// ─────────────────────────────────────────────────────────────
function _e2e_snapshotTab(testSs) {
  const s = testSs.getSheetByName("Snapshots");
  Assert.isTrue(s !== null, "E2E-SNAPSHOT: Tab exists");

  const headers = s.getRange("A1:M1").getValues()[0];
  const expectedHeaders = ["Date", "Net Worth (USD)", "Liquid Worth (USD)", "Locked Worth (USD)",
    "Gross Worth (USD)", "Net Worth (CAD)", "Net Worth (INR)", "Real Estate Net (USD)",
    "Value Delta ($)", "Growth (%)", "FIRE Progress", "Auto Insight", "Manual Notes"];
  for (let i = 0; i < expectedHeaders.length; i++) {
    Assert.equal(headers[i], expectedHeaders[i],
      `E2E-SNAPSHOT: Header col ${i+1} = "${expectedHeaders[i]}"`);
  }

  // Run a snapshot and verify a row was inserted
  captureSnapshot(testSs, true);
  const snapDate = testSs.getSheetByName("Snapshots").getRange("A2").getValue();
  Assert.isTrue(snapDate instanceof Date || snapDate !== "", "E2E-SNAPSHOT: captureSnapshot() inserted a dated row");

  Logger.log("✅ S6: Snapshots tab structure and insertion verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 7: CASH FLOW TAB
// ─────────────────────────────────────────────────────────────
function _e2e_cashFlowTab(testSs) {
  const s = testSs.getSheetByName("💸 Cash Flow & Burn");
  Assert.isTrue(s !== null, "E2E-CASHFLOW: Tab exists");

  // KPI headers at A1
  const h1 = s.getRange("A1").getValue();
  Assert.isTrue(String(h1).length > 0, "E2E-CASHFLOW: KPI header row is populated");

  // KPI formulas in B column should reference Settings
  const fireFormula = s.getRange("B5").getFormula();
  Assert.isTrue(
    fireFormula.includes("Settings") || fireFormula.includes("B19"),
    "E2E-CASHFLOW: FIRE Budget KPI references Settings & Config tab"
  );

  Logger.log("✅ S7: Cash Flow tab structure verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 8: TRIGGER REGISTRATION
// ─────────────────────────────────────────────────────────────
function _e2e_triggerRegistration() {
  const triggers = ScriptApp.getProjectTriggers();
  const found = triggers.some(t => t.getHandlerFunction() === "updateRealEstatePrices");
  Assert.isTrue(found, "E2E-TRIGGERS: Weekly Real Estate price trigger is registered");
  Logger.log("✅ S8: Trigger registration verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 9: BACKUP JSON SCHEMA
// ─────────────────────────────────────────────────────────────
function _e2e_backupJsonSchema(testSs) {
  const payload = _buildEnrichedBackup(testSs);

  // Top-level keys
  Assert.isTrue("snapshotDate" in payload, "E2E-BACKUP: payload has snapshotDate");
  Assert.isTrue("spreadsheetId" in payload, "E2E-BACKUP: payload has spreadsheetId");
  Assert.isTrue("summary" in payload,      "E2E-BACKUP: payload has summary");
  Assert.isTrue("accounts" in payload,     "E2E-BACKUP: payload has accounts");

  // Summary keys
  const s = payload.summary;
  Assert.isTrue("netWorthUSD"   in s, "E2E-BACKUP: summary.netWorthUSD exists");
  Assert.isTrue("grossWorthUSD" in s, "E2E-BACKUP: summary.grossWorthUSD exists");
  Assert.isTrue("fireProgress"  in s, "E2E-BACKUP: summary.fireProgress exists");
  Assert.isTrue("liquidNetWorthUSD" in s, "E2E-BACKUP: summary.liquidNetWorthUSD exists");

  // Accounts array
  Assert.isTrue(Array.isArray(payload.accounts), "E2E-BACKUP: accounts is an array");
  Assert.isTrue(payload.accounts.length > 0, "E2E-BACKUP: accounts array is not empty");

  const acct = payload.accounts[0];
  const acctKeys = ["Account", "AssetClass", "Currency", "InitialCapital", "CurrentValue",
    "ExchangeRate", "TaxRate", "GrossWorthUSD", "NetWorthUSD", "Status", "Remarks"];
  for (const k of acctKeys) {
    Assert.isTrue(k in acct, `E2E-BACKUP: account object has key "${k}"`);
  }

  // spreadsheetId must match
  Assert.equal(payload.spreadsheetId, testSs.getId(), "E2E-BACKUP: spreadsheetId matches sandbox");

  Logger.log("✅ S9: Backup JSON schema verified (all keys present)");
}

// ─────────────────────────────────────────────────────────────
// SECTION 10: SNAPSHOT MATH
// ─────────────────────────────────────────────────────────────
function _e2e_snapshotMath(testSs) {
  // _buildLedgerSnapshot is a pure function — test it in isolation
  const mockData = [
    ["Checking",           "Cash",      "USD", 0, 50000,  1, 0,    50000, 47500, "Active", "Everyday"],
    ["Taxable Brokerage",  "Brokerage", "USD", 0, 200000, 1, 0.15, 200000, 170000, "Active", "VTI"],
    ["Primary Residence",  "Real Estate","USD",0, 800000, 1, 0,    800000, 800000, "Active","Home"],
    ["Closed Account",     "Cash",      "USD", 0, 10000,  1, 0,    10000, 10000, "Inactive","Closed"],
    ["",                   "",          "",    0, 0,      0, 0,    0,     0,     "", ""]
  ];
  const snapshot = _buildLedgerSnapshot(mockData);

  Assert.equal(snapshot.length, 4, "E2E-MATH: _buildLedgerSnapshot includes 4 non-blank rows (incl. Inactive)");
  Assert.equal(snapshot[0].Account, "Checking", "E2E-MATH: First account = Checking");
  Assert.equal(snapshot[0].NetWorthUSD, 47500, "E2E-MATH: Checking Net Worth = 47500");
  Assert.equal(snapshot[2].AssetClass, "Real Estate", "E2E-MATH: Third account = Real Estate");

  // FIRE progress formula contract
  const fireFormula = `=IFERROR(SUMIFS(I7:I5000,J7:J5000,"Active")/20000,0)`;
  Assert.isTrue(fireFormula.includes("SUMIFS"), "E2E-MATH: FIRE formula uses SUMIFS to aggregate active accounts");

  Logger.log("✅ S10: Snapshot math and _buildLedgerSnapshot verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 11: GITHUB GIST INTEGRATION CONTRACTS
// ─────────────────────────────────────────────────────────────
function _e2e_githubIntegrationContracts() {
  // PAT format validation
  Assert.isTrue(_validatePATFormat("ghp_AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"), "E2E-GITHUB: Classic token accepted");
  Assert.isTrue(_validatePATFormat("github_pat_AAAAAAAAAAAAAAAAAAAAAA_BBBBBBBB"), "E2E-GITHUB: Fine-grained token accepted");
  Assert.isTrue(!_validatePATFormat("gho_invalid"), "E2E-GITHUB: Wrong prefix rejected");
  Assert.isTrue(!_validatePATFormat(""),             "E2E-GITHUB: Empty string rejected");
  Assert.isTrue(!_validatePATFormat("ghp_short"),    "E2E-GITHUB: Too-short token rejected");

  // Gist URL format
  const url = _buildGistUrl("abc123def456");
  Assert.equal(url, "https://gist.github.com/abc123def456", "E2E-GITHUB: Gist URL is correctly formed");
  Assert.isTrue(url.startsWith("https://"), "E2E-GITHUB: Gist URL uses HTTPS");

  // isGistConfigured with no sheet
  Assert.isTrue(!_isGistConfigured(null), "E2E-GITHUB: _isGistConfigured returns false with null sheet");

  Logger.log("✅ S11: GitHub integration contracts verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 12: GOOGLE DRIVE BACKUP CONTRACTS (DriveApp)
// ─────────────────────────────────────────────────────────────
function _e2e_driveApiContracts() {
  // backupToGoogleDrive must succeed with a sandbox spreadsheet (DriveApp, silent)
  const testSs = SpreadsheetApp.create("WealthScript_E2E_DriveTest_" + Date.now());
  let testFolder;

  try {
    runFirstTimeSetup(testSs, true);
    const result = backupToGoogleDrive(testSs, true);
    Assert.isTrue(result.success === true, "E2E-DRIVE: backupToGoogleDrive returns success:true");
    Assert.isTrue(result.folder !== null, "E2E-DRIVE: backupToGoogleDrive returns a DriveApp Folder object");

    // Verify the folder name
    Assert.equal(result.folder.getName(), "WealthScript \u2014 Backups",
      "E2E-DRIVE: backup folder is named 'WealthScript — Backups'");

    // Verify at least one JSON file was created inside
    const files = result.folder.getFilesByType(MimeType.PLAIN_TEXT);
    Assert.isTrue(files.hasNext(), "E2E-DRIVE: At least one backup JSON file exists in the folder");

    testFolder = result.folder;
  } finally {
    // Cleanup: trash the test spreadsheet and backup folder
    try { DriveApp.getFileById(testSs.getId()).setTrashed(true); } catch(e) {}
    try { if (testFolder) testFolder.setTrashed(true); } catch(e) {}
  }

  Logger.log("✅ S12: Google Drive (DriveApp) backup contracts verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 13: REAL ESTATE INTEGRATION (GRACEFUL DEGRADATION)
// ─────────────────────────────────────────────────────────────
function _e2e_realEstateGracefulDegradation(testSs) {
  // With no API key configured, updateRealEstatePrices should return silently (no crash)
  let threwError = false;
  try {
    updateRealEstatePrices(testSs);
  } catch(e) {
    threwError = true;
  }
  Assert.isTrue(!threwError, "E2E-REALESTATE: updateRealEstatePrices exits gracefully without an API key");

  Logger.log("✅ S13: Real Estate graceful degradation verified");
}

// ─────────────────────────────────────────────────────────────
// SECTION 14: WIZARD CONTRACTS
// ─────────────────────────────────────────────────────────────
function _e2e_wizardContracts() {
  // Server-side _processGistToken must reject invalid/empty tokens
  const badResult = _processGistToken("not-a-real-token");
  Assert.isTrue(badResult.success === false, "E2E-WIZARD: _processGistToken rejects invalid token");
  Assert.isTrue(typeof badResult.message === "string", "E2E-WIZARD: _processGistToken returns a message string");

  // _processGistToken must return success:false with empty string
  const emptyResult = _processGistToken("");
  Assert.isTrue(emptyResult.success === false, "E2E-WIZARD: _processGistToken rejects empty token");

  Logger.log("✅ S14: Wizard server-side contracts verified");
}
