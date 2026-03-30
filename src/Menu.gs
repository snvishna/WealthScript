/**
 * Creates the custom menu in the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WealthScript')
      .addItem('🚀 Run First Time Setup', 'runFirstTimeSetup')
      .addSeparator()
      .addItem('📸 Log Snapshot & Cloud Sync', 'captureSnapshot')
      .addItem('🔄 Refresh Real Estate Prices', 'updateRealEstatePrices')
      .addItem('📊 Update Visual Dashboards', 'updateVisualDashboards')
      .addItem('💸 Rebuild Cash Flow Tab', 'buildCashFlowTab')
      .addSeparator()
      .addItem('🔐 Setup GitHub Backup', 'setupGistWizard')
      .addItem('📁 Setup Google Drive Backup', 'setupDriveBackup')
      .addItem('☁️ Force Cloud Backup', 'forceBackup')
      .addToUi();
}

/**
 * MASTER SETUP: Builds all tabs and sets up automated cron jobs.
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Optional target spreadsheet
 * @param {boolean} [silent=false] - Whether to suppress UI alerts
 */
function runFirstTimeSetup(ss_inject, silent = false) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  
  buildSettingsTab(ss);
  buildHoldingsTab(ss);      // must exist BEFORE buildPortfolioTracker injects SUMPRODUCT formulas
  buildPortfolioTracker(ss); // references 'Brokerage Holdings' tab — tab must already exist
  buildSnapshotTab(ss);
  buildCashFlowTab(ss);

  // Setup the Weekly Real Estate Trigger
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateRealEstatePrices') {
      triggerExists = true;
      break;
    }
  }

  if (!triggerExists) {
    ScriptApp.newTrigger('updateRealEstatePrices')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(8)
      .create();
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "✅ Setup Complete!\n\n" +
      "Your dashboard and all tabs are ready.\n\n" +
      "📋 Next Steps:\n" +
      "1. Review the 'Settings & Config' tab\n" +
      "2. Highlight rows 7+ on your Dashboard → Format > Convert to Table\n\n" +
      "☁️ Secure Your Data (Optional):\n" +
      "• WealthScript > 🔐 Setup GitHub Backup\n" +
      "• WealthScript > 📁 Setup Google Drive Backup"
    );
  }
}
