/**
 * Creates the custom menu in the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Destructive builders are hidden once the workbook is live, so a stray click
  // can never wipe a populated ledger. They stay reachable under Danger Zone.
  const isSetUp = !!(ss.getSheetByName("Dashboard & Ledger") && ss.getSheetByName("Brokerage Holdings"));
  const menu = ui.createMenu('WealthScript');

  if (!isSetUp) {
    menu.addItem('🚀 Run First Time Setup', 'runFirstTimeSetup').addSeparator();
  }

  menu.addItem('📸 Log Snapshot & Cloud Sync', 'captureSnapshot')
      .addItem('🔄 Refresh Real Estate Prices', 'updateRealEstatePrices')
      .addItem('📊 Update Visual Dashboards', 'updateVisualDashboards')
      .addSeparator()
      .addItem('🩺 Check Formula Health', 'checkFormulaHealth')
      .addItem('🛠 Repair Formulas', 'repairFormulas')
      .addItem('🧭 Migrate Sheet Layout', 'migrateSheetLayout')
      .addItem('🔒 Secure Stored Credentials', 'migrateSecretsToProperties')
      .addSeparator();

  // Broker sync is opt-in. The sync action only appears once a provider is
  // configured, so users who track holdings by hand never see it.
  if (isSetUp && isIbkrFlexConfigured()) {
    menu.addItem('🔗 Sync IBKR Positions', 'syncIbkrPositions')
        .addItem('⚙️ Reconfigure IBKR Flex', 'setupIbkrFlexWizard');
  } else if (isSetUp) {
    menu.addItem('🔗 Connect IBKR (optional)', 'setupIbkrFlexWizard');
  }

  menu.addSeparator()
      .addItem('🔐 Setup GitHub Backup', 'setupGistWizard')
      .addItem('📁 Setup Google Drive Backup', 'setupDriveBackup')
      .addItem('☁️ Force Cloud Backup', 'forceBackup');

  if (isSetUp) {
    menu.addSeparator().addSubMenu(
      ui.createMenu('⚠️ Danger Zone')
        .addItem('💥 Rebuild ALL tabs (erases your data)', 'rebuildEverythingDestructive')
        .addItem('💥 Rebuild Cash Flow tab (erases expenses)', 'rebuildCashFlowDestructive')
    );
  }

  menu.addToUi();
}

/**
 * MASTER SETUP: Builds all tabs and sets up automated cron jobs.
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Optional target spreadsheet
 * @param {boolean} [silent=false] - Whether to suppress UI alerts
 * @param {boolean} [force=false] - Bypass the populated-ledger guard
 */
function runFirstTimeSetup(ss_inject, silent = false, force = false) {
  const ss = _resolveSpreadsheet(ss_inject);

  // Every builder below calls sheet.clear(). Refuse to run against a populated
  // ledger unless explicitly forced via the Danger Zone menu.
  if (!force) {
    const existing = ss.getSheetByName("Dashboard & Ledger");
    if (existing) {
      const rows = existing.getRange(LEDGER_FIRST_ROW, 1, LEDGER_NUM_ROWS, 1).getValues();
      if (_hasUserLedgerData(rows)) {
        if (!silent) {
          SpreadsheetApp.getUi().alert(
            "🛑 Setup aborted — this workbook already contains your data.\n\n" +
            "Running setup would erase every tab.\n\n" +
            "You probably want one of these instead:\n" +
            "  • 🛠 Repair Formulas — restore damaged formulas, keep your data\n" +
            "  • 🧭 Migrate Sheet Layout — apply layout changes from a new version\n\n" +
            "If you really do want to start over, use ⚠️ Danger Zone > Rebuild ALL tabs."
          );
        }
        return;
      }
    }
  }
  
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
