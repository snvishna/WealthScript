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
 */
function runFirstTimeSetup() {
  buildSettingsTab();
  buildPortfolioTracker();
  buildHoldingsTab();
  buildSnapshotTab();
  buildCashFlowTab();

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

  SpreadsheetApp.getUi().alert("Setup Complete!\n\n1. Review the 'Settings & Config' tab.\n2. Highlight rows 7+ on your Dashboard and click Format > Convert to Table.");
}
