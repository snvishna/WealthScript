/**
 * ==========================================
 * GLOBAL CONFIGURATION: THEME & ACCOUNTS
 * ==========================================
 */

const THEME = {
  canvas: "#F8FAFC",
  headerBg: "#2563EB",
  headerText: "#FFFFFF",
  kpiCardBg: "#FFFFFF",
  mutedText: "#64748B",
  accentBlue: "#2563EB",
  accentEmerald: "#059669",
  accentViolet: "#7C3AED",
  quickStats: {
    liquidBg: "#E0F2FE", liquidFg: "#0369A1",
    lockedBg: "#FEF2F2", lockedFg: "#BE123C",
    fireBg: "#FAF5FF",   fireFg: "#7E22CE"
  },
  assetRows: {
    "Cash":           "#ECFDF5", 
    "Brokerage":      "#EFF6FF",
    "Retirement":     "#EEF2FF", 
    "Health Savings": "#F0FDFA",
    "Real Estate":    "#FFF7ED", 
    "Crypto":         "#FDF2F8",
    "Commodity":      "#FEFCE8", 
    "Insurance":      "#FAF5FF",
    "Receivable":     "#ECFEFF", 
    "Liability":      "#FEF2F2"
  },
  assetText: "#0F172A",
  negativeValueBg: "#fff1f2",
  negativeValueFg: "#be123c",
  accentBar: "#E2E8F0",
  titleBanner: { bg: "#1E293B", text: "#F8FAFC" },
  charts: {
    donut: ['#3B82F6', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#14B8A6', '#F43F5E', '#9CA3AF'],
    area: ['#3B82F6'],
    stacked: ['#10B981', '#F43F5E'],
    gridlines: "transparent",
    axisText: "#64748B",
    legendText: "#0F172A"
  }
};

/** Edit this array to change the default accounts generated during First Time Setup. */
const DEFAULT_PORTFOLIO_DATA = [
  // --- Cash & Checking ---
  ["Primary Checking",        "Cash",          "USD", 0,          8000,       "", 0.00, "", "", "Active", "Everyday expenses account"],
  ["High Yield Savings",      "Cash",          "USD", 0,          40000,      "", 0.00, "", "", "Active", "Emergency fund (6 months)"],
  ["International Bank",      "Cash",          "CAD", 0,          10000,      "", 0.00, "", "", "Active", "Canadian bank account"],

  // --- Brokerage ---
  ["Taxable Brokerage",       "Brokerage",     "USD", 0,          5000,       "", 0.15, "", "", "Active", "Index funds (VTI / VXUS)"],
  ["Angel Investing",         "Brokerage",     "USD", 2000,       1500,       "", 0.30, "", "", "Active", "Private equity / startup investing"],

  // --- Crypto ---
  ["Crypto Exchange",         "Crypto",        "USD", 0,          8000,       "", 0.30, "", "", "Active", "BTC / ETH"],

  // --- Retirement ---
  ["401k (Employer Plan)",    "Retirement",    "USD", 0,          120000,     "", 0.20, "", "", "Active", "Pre-tax employer 401k"],
  ["Roth IRA",                "Retirement",    "USD", 0,          45000,      "", 0.00, "", "", "Active", "Tax-free Roth IRA (no tax on withdrawal)"],

  // --- Health Savings ---
  ["HSA Account",             "Health Savings","USD", 0,          15000,      "", 0.00, "", "", "Active", "Triple tax-advantaged HSA"],

  // --- Real Estate ---
  ["Primary Residence",       "Real Estate",   "USD", 400000,     450000,     "", 0.20, "", "", "Active", "Primary home"],
  ["Investment Property",     "Real Estate",   "USD", 300000,     350000,     "", 0.20, "", "", "Active", "Rental property"],

  // --- Insurance ---
  ["Endowment Policy",        "Insurance",     "INR", 0,          0,          "", 0.20, "", "", "Active", "Maturity value estimate"],

  // --- Liabilities (Enter Current Value as Negative) ---
  ["Credit Card 1",           "Liability",     "USD", 0,          0,          "", 0.00, "", "", "Active", "Paid in full monthly"],
  ["Credit Card 2 (CAD)",     "Liability",     "CAD", 0,          -1500,      "", 0.00, "", "", "Active", "Canadian credit card"],
  ["Credit Card 3",           "Liability",     "USD", 0,          -2000,      "", 0.00, "", "", "Active", ""],
  ["Credit Card 4",           "Liability",     "USD", 0,          -1200,      "", 0.00, "", "", "Active", ""],
  ["Auto Loan",               "Liability",     "USD", 0,          -8000,      "", 0.00, "", "", "Active", "Vehicle loan"],
  ["Primary Mortgage",        "Liability",     "USD", 0,          -380000,    "", 0.00, "", "", "Active", "Home mortgage — 30yr fixed"],
];

/**
 * CLOUD DISASTER RECOVERY CONFIGURATION
 * Provide your GitHub PAT here before running First Time Setup to auto-create your backup Gist.
 */
const CLOUD_SYNC_CONFIG = {
  githubPAT: "", // Enter your GitHub Personal Access Token (must have 'gist' scope)
  gistId: ""     // Leave blank to auto-create a new Secret Gist during setup, OR paste an existing ID.
};

/**
 * DASHBOARD CONFIGURATION
 * Edit secondaryCurrencies to show your currencies in the top KPI cards.
 * Supports any valid GOOGLEFINANCE code: EUR, GBP, AUD, JPY, MXN, SGD, etc.
 * Only the first two entries are rendered (layout: USD + 2 secondary cards).
 */
const DASHBOARD_CONFIG = {
  secondaryCurrencies: ["CAD", "INR"], // ← Change these to your currencies
  fireTargetUSD: 3000000,              // ← Your FIRE / net worth target in USD
};
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
/**
 * ==========================================
 * BACKUP SETUP WIZARD
 * Guided flows for GitHub Gist & Google Drive
 * ==========================================
 */

/**
 * Pure helper: validates that a string looks like a GitHub PAT.
 * Accepts classic tokens (ghp_*) and fine-grained tokens (github_pat_*).
 * @param {string} token - The raw token string from user input.
 * @returns {boolean}
 */
function _validatePATFormat(token) {
  if (!token || typeof token !== 'string') return false;
  const trimmed = token.trim();
  return /^ghp_[A-Za-z0-9]{36,}$/.test(trimmed) || /^github_pat_[A-Za-z0-9_]{20,}$/.test(trimmed);
}

/**
 * Pure helper: builds Gist URL from an ID.
 * @param {string} gistId
 * @returns {string}
 */
function _buildGistUrl(gistId) {
  return `https://gist.github.com/${gistId}`;
}

/**
 * Guided GitHub Gist Backup Wizard.
 * Uses an inline clickable link (popup-blocker safe) instead of window.open().
 * Validates token, creates Gist, populates Settings with credentials + hyperlink.
 */
function setupGistWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  // Step 1: Show instructions with a clickable link (no popup blocker issues)
  const patUrl = "https://github.com/settings/tokens/new?scopes=gist&description=WealthScript+Backup";
  const htmlOutput = HtmlService
    .createHtmlOutput(
      `<style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; color: #1a1a1a; }
        .step { margin-bottom: 12px; }
        .step-num { display: inline-block; background: #2563EB; color: white; border-radius: 50%;
                    width: 24px; height: 24px; text-align: center; line-height: 24px; font-size: 13px; margin-right: 8px; }
        a.btn { display: inline-block; background: #2563EB; color: white !important; padding: 10px 20px;
                border-radius: 6px; text-decoration: none; font-weight: bold; margin: 12px 0; }
        a.btn:hover { background: #1d4ed8; }
        code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
      </style>
      <div class="step"><span class="step-num">1</span> Click the button below to open GitHub's token page:</div>
      <a class="btn" href="${patUrl}" target="_blank">🔐 Open GitHub Token Page →</a>
      <div class="step"><span class="step-num">2</span> The <code>gist</code> scope is pre-selected — just click <b>"Generate token"</b></div>
      <div class="step"><span class="step-num">3</span> Copy the token (starts with <code>ghp_</code>)</div>
      <div class="step"><span class="step-num">4</span> Close this dialog, then paste it in the next prompt</div>`
    )
    .setWidth(460)
    .setHeight(280);
  ui.showModalDialog(htmlOutput, "🔐 Step 1 of 2: Create GitHub Token");

  // Step 2: Prompt for token
  const response = ui.prompt(
    "🔐 Step 2 of 2: Paste Your Token",
    "Paste the GitHub Personal Access Token you just created (starts with ghp_):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Wizard cancelled. No changes were made.");
    return;
  }

  const pat = response.getResponseText().trim();

  if (!_validatePATFormat(pat)) {
    ui.alert("❌ Invalid Token Format\n\nExpected a token starting with 'ghp_' or 'github_pat_'.\nPlease try the wizard again from the WealthScript menu.");
    return;
  }

  // Step 3: Validate token against GitHub API
  try {
    const testResponse = UrlFetchApp.fetch("https://api.github.com/user", {
      headers: { "Authorization": "Bearer " + pat, "Accept": "application/vnd.github.v3+json" },
      muteHttpExceptions: true
    });
    if (testResponse.getResponseCode() !== 200) {
      ui.alert("❌ Token Validation Failed\n\nGitHub rejected the token. Please ensure it has 'gist' scope and try again.");
      return;
    }
  } catch (e) {
    ui.alert("❌ Network Error\n\n" + e.message);
    return;
  }

  // Step 4: Create Gist and populate Settings
  const gistId = autoCreateGist(pat);
  if (!gistId) {
    ui.alert("❌ Gist Creation Failed\n\nThe token is valid but Gist creation failed. Check Apps Script logs for details.");
    return;
  }

  configSheet.getRange("B6").setValue(pat);
  configSheet.getRange("B7").setValue(gistId);

  const gistUrl = _buildGistUrl(gistId);
  const richGistLink = SpreadsheetApp.newRichTextValue()
    .setText(gistUrl)
    .setLinkUrl(gistUrl)
    .build();
  configSheet.getRange("B8").setRichTextValue(richGistLink);

  ui.alert(`✅ GitHub Backup Connected!\n\nYour private Gist has been created and linked.\nGist ID: ${gistId}\n\nEvery snapshot will now auto-sync to GitHub.`);
}

/**
 * Guided Google Drive Backup Setup.
 * Creates the backup folder (if needed), retrieves its URL,
 * and populates the Settings tab with a clickable hyperlink.
 */
function setupDriveBackup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");
  const FOLDER_NAME = "WealthScript \u2014 Backups";

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  try {
    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(FOLDER_NAME);
    const folderUrl = folder.getUrl();

    const richDriveLink = SpreadsheetApp.newRichTextValue()
      .setText(folderUrl)
      .setLinkUrl(folderUrl)
      .build();
    configSheet.getRange("B9").setRichTextValue(richDriveLink);

    ui.alert(`✅ Google Drive Backup Connected!\n\nFolder: "${FOLDER_NAME}"\n\nA clickable link has been added to your Settings tab.\nEvery snapshot will now auto-sync a dated JSON file here.`);
  } catch (e) {
    ui.alert("❌ Drive Setup Failed:\n" + e.message);
  }
}
/**
 * Fetches Zestimates using config from the Settings tab.
 */
function updateRealEstatePrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");
  
  if(!configSheet) return SpreadsheetApp.getUi().alert("Settings tab missing. Run First Time Setup.");

  const apiKey = configSheet.getRange("B2").getValue();
  const apiHost = configSheet.getRange("B3").getValue();
  
  if (!apiKey || apiKey === "PASTE_KEY_HERE") return; 

  const propData = configSheet.getRange("A21:B37").getValues();
  const properties = [];
  for (let i = 0; i < propData.length; i++) {
    if (propData[i][0] && propData[i][1]) {
      properties.push({ name: String(propData[i][0]), zpid: String(propData[i][1]) });
    }
  }

  const accountNames = sheet.getRange("A1:A100").getValues().flat();
  const currentValues = sheet.getRange("E1:E100").getValues();
  let updated = false;

  properties.forEach(prop => {
    try {
      const url = `https://${apiHost}/api/property-details/byzpid?zpid=${prop.zpid}`;
      const options = {
        method: 'GET',
        headers: {
          'X-RapidAPI-Key': apiKey,
          'X-RapidAPI-Host': apiHost
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const zestimate = data.property.zestimate; 
        const targetRowIdx = accountNames.indexOf(prop.name); 
        
        if (targetRowIdx >= 0 && zestimate) {
          currentValues[targetRowIdx][0] = zestimate;
          updated = true;
        }
      }
    } catch (e) {
      Logger.log(`Script crashed on ${prop.name}: ${e}`);
    }
  });

  if (updated) {
    sheet.getRange("E1:E100").setValues(currentValues);
  }
}
function buildSettingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings & Config");
  if (!sheet) sheet = ss.insertSheet("Settings & Config");
  else sheet.clear();

  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:C100").setBackground(THEME.canvas);

  let pat = CLOUD_SYNC_CONFIG.githubPAT || "PASTE_GITHUB_TOKEN_HERE";
  let gistId = CLOUD_SYNC_CONFIG.gistId;

  if (pat !== "" && pat !== "PASTE_GITHUB_TOKEN_HERE" && gistId === "") {
    gistId = autoCreateGist(pat);
  }
  if (!gistId) gistId = "PASTE_GIST_ID_HERE";

  const styleRow = (range, bg) => range.setBackground(bg).setVerticalAlignment("middle");

  sheet.getRange("A1").setValue("REAL ESTATE API CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  styleRow(sheet.getRange("A2:B2"), THEME.kpiCardBg).setValues([["RapidAPI Key", "PASTE_KEY_HERE"]]);
  styleRow(sheet.getRange("A3:B3"), THEME.kpiCardBg).setValues([["RapidAPI Host", "real-estate101.p.rapidapi.com"]]);

  sheet.getRange("A5").setValue("CLOUD BACKUP CONFIG (DISASTER RECOVERY)").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  styleRow(sheet.getRange("A6:B6"), THEME.kpiCardBg).setValues([["GitHub PAT (gist scope)", pat]]);
  styleRow(sheet.getRange("A7:B7"), THEME.kpiCardBg).setValues([["GitHub Gist ID", gistId]]);
  styleRow(sheet.getRange("A8:B8"), THEME.kpiCardBg).setValues([["GitHub Gist URL", "Run '🔐 Setup GitHub Backup' from the menu"]]);
  sheet.getRange("A8").setFontColor(THEME.mutedText);
  sheet.getRange("B8").setFontColor(THEME.accentBlue);
  styleRow(sheet.getRange("A9:B9"), THEME.kpiCardBg).setValues([["Google Drive Backup Folder", "Run '📁 Setup Google Drive Backup' from the menu"]]);
  sheet.getRange("A9").setFontColor(THEME.mutedText);
  sheet.getRange("B9").setFontColor(THEME.accentBlue);

  sheet.getRange("A11").setValue("FIRE & CASH FLOW CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  const fireConfig = [
    ["Target Monthly FIRE Budget (USD)", 20000],
    ["Estimated Monthly Rental Income (USD)", 0],
    ["Annual Portfolio Return Rate", 0.07]
  ];
  const fireRange = sheet.getRange(12, 1, fireConfig.length, 2);
  fireRange.setValues(fireConfig);
  styleRow(fireRange, THEME.kpiCardBg);
  sheet.getRange(12, 2).setNumberFormat("$#,##0");
  sheet.getRange(13, 2).setNumberFormat("$#,##0");
  sheet.getRange(14, 2).setNumberFormat("0.00%");

  sheet.getRange("A15").setValue("DASHBOARD CURRENCY CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  const currencyConfig = [
    ["Secondary Currency (Card 2)", (DASHBOARD_CONFIG.secondaryCurrencies[0] || "CAD")],
    ["Secondary Currency (Card 3)", (DASHBOARD_CONFIG.secondaryCurrencies[1] || "INR")]
  ];
  const currRange = sheet.getRange(16, 1, currencyConfig.length, 2);
  currRange.setValues(currencyConfig);
  styleRow(currRange, THEME.kpiCardBg);
  sheet.getRange("B16").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");
  sheet.getRange("B17").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");

  sheet.getRange("A19").setValue("REAL ESTATE ZPID MAPPING").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  sheet.getRange("A20:B20")
    .setValues([["Account Name (Must match Dashboard exactly)", "ZPID"]])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleMapping = [
    ["Primary Residence", "12345678"],
    ["Investment Property 1", "87654321"]
  ];
  sheet.getRange(21, 1, sampleMapping.length, 2).setValues(sampleMapping);

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 350);
}

/**
 * 2. Builds the Dashboard & Ledger with full professional formatting.
 */
function buildPortfolioTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) sheet = ss.insertSheet("Dashboard & Ledger");
  else sheet.clear();

  sheet.setHiddenGridlines(true);
  sheet.getRange("A2:K5").setBackground(THEME.canvas); 
  
  sheet.getRange("A1:K1").merge()
    .setValue("💰  NET WORTH DASHBOARD")
    .setBackground(THEME.titleBanner.bg).setFontColor(THEME.titleBanner.text)
    .setFontWeight("bold").setFontSize(15)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 44);

  const USD_ABBR_FMT = '[>999999]"$"0.00,,"M";[>999]"$"0,"K";"$"0';
  const PLAIN_ABBR_FMT = '[>999999]0.00,,"M";[>999]0,"K";0.00';

  const currencySymbol = (code) => {
    const SYM = { USD:'$', EUR:'€', GBP:'£', INR:'₹', JPY:'¥',
                  CAD:'CA$', AUD:'A$', SGD:'S$', CHF:'Fr', MXN:'MX$' };
    return SYM[code.toUpperCase()] || code;
  };

  const abbrFmt = (code) => {
    const s = currencySymbol(code);
    return `[>999999]"${s}"0.00,,"M";[>999]"${s}"0,"K";"${s}"0`;
  };

  const CARD_STYLES = [
    { bg:THEME.kpiCardBg, labelFg:THEME.mutedText, valueFg:THEME.accentBlue, subFg:THEME.mutedText }, 
    { bg:THEME.kpiCardBg, labelFg:THEME.mutedText, valueFg:THEME.accentEmerald, subFg:THEME.mutedText }, 
    { bg:THEME.kpiCardBg, labelFg:THEME.mutedText, valueFg:THEME.accentViolet, subFg:THEME.mutedText }, 
  ];
  const CARD_LAYOUT = [
    { bg:"A2:C3", lbl:"A", val:"B" },
    { bg:"D2:G3", lbl:"D", val:"E" },
    { bg:"H2:K3", lbl:"H", val:"I" },
  ];

  const s0 = CARD_STYLES[0]; const c0 = CARD_LAYOUT[0];
  sheet.getRange(c0.bg).setBackground(s0.bg);
  sheet.getRange(`${c0.lbl}2`).setValue("Net Worth (USD)").setFontColor(s0.labelFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange(`${c0.val}2`).setFormula('=SUMIFS(I7:I5000,J7:J5000,"Active")')
    .setNumberFormat(USD_ABBR_FMT).setFontColor(s0.valueFg).setFontSize(14).setFontWeight("bold");
  sheet.getRange(`${c0.lbl}3`).setValue("Gross Worth (USD)").setFontColor(s0.labelFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange(`${c0.val}3`).setFormula('=SUMIFS(H7:H5000,J7:J5000,"Active")')
    .setNumberFormat(USD_ABBR_FMT).setFontColor(s0.subFg).setFontSize(11);

  const SETTINGS_CURRENCY_CELLS = ["'Settings & Config'!B16", "'Settings & Config'!B17"];
  SETTINGS_CURRENCY_CELLS.slice(0, 2).forEach((settingsCell, idx) => {
    const sn = CARD_STYLES[idx + 1]; const cn = CARD_LAYOUT[idx + 1];
    sheet.getRange(cn.bg).setBackground(sn.bg);
    sheet.getRange(`${cn.lbl}2`)
      .setFormula(`="Net Worth ("&${settingsCell}&")"`)  
      .setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}2`)
      .setFormula(`=B2*IFERROR(GOOGLEFINANCE("CURRENCY:USD"&${settingsCell}),1)`)
      .setNumberFormat(PLAIN_ABBR_FMT).setFontColor(sn.valueFg).setFontSize(14).setFontWeight("bold");
    sheet.getRange(`${cn.lbl}3`)
      .setFormula(`="Gross Worth ("&${settingsCell}&")"`)  
      .setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}3`)
      .setFormula(`=B3*IFERROR(GOOGLEFINANCE("CURRENCY:USD"&${settingsCell}),1)`)
      .setNumberFormat(PLAIN_ABBR_FMT).setFontColor(sn.subFg).setFontSize(11);
  });

  sheet.setRowHeight(2, 38); sheet.setRowHeight(3, 28);

  const LIQUID_CLASSES = ['"Cash"','"Brokerage"','"Crypto"','"Receivable"'];
  const liquidParts = LIQUID_CLASSES.map(c => `SUMIFS(I7:I5000,J7:J5000,"Active",B7:B5000,${c})`).join('+');
  const fireTarget  = DASHBOARD_CONFIG.fireTargetUSD || 3000000;
  const fireLabel   = `🔥 FIRE Progress ($${(fireTarget / 1e6).toFixed(0)}M)`;

  sheet.getRange("A4:C4").setBackground(THEME.quickStats.liquidBg);
  sheet.getRange("A4").setValue("🌊 Liquid Net Worth").setFontColor(THEME.quickStats.liquidFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("B4").setFormula(`=${liquidParts}`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.liquidFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("D4:G4").setBackground(THEME.quickStats.lockedBg);
  sheet.getRange("D4").setValue("🔒 Locked Net Worth").setFontColor(THEME.quickStats.lockedFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("E4").setFormula(`=SUMIFS(I7:I5000,J7:J5000,"Active")-(${liquidParts})`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.lockedFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("H4:K4").setBackground(THEME.quickStats.fireBg);
  sheet.getRange("H4").setValue(fireLabel).setFontColor(THEME.quickStats.fireFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("I4").setFormula(`=IFERROR(SUMIFS(I7:I5000,J7:J5000,"Active")/${fireTarget},0)`)
    .setNumberFormat("0.0%").setFontColor(THEME.quickStats.fireFg).setFontSize(11).setFontWeight("bold");

  sheet.setRowHeight(4, 28);

  sheet.getRange("A5:K5").setBackground(THEME.accentBar);
  sheet.setRowHeight(5, 3);

  const headers = ["Account","Asset Class","Currency","Initial Capital","Current Value","Exchange Rate (to USD)","Tax Rate","Gross Worth (USD)","Net Worth (USD)","Status","Remarks"];
  sheet.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(6, 36);

  sheet.getRange(7, 1, DEFAULT_PORTFOLIO_DATA.length, headers.length).setValues(DEFAULT_PORTFOLIO_DATA);

  const NUM_ROWS = 70;
  const exch = [], gross = [], net = [];
  for (let i = 0; i < NUM_ROWS; i++) {
    const r = i + 7;
    exch.push([`=IF(ISBLANK(C${r}),"",IF(TRIM(UPPER(C${r}))="USD",1,IFERROR(GOOGLEFINANCE("CURRENCY:"&TRIM(UPPER(C${r}))&"USD"),"Error")))` ]);
    gross.push([`=IF(AND(ISNUMBER(E${r}),ISNUMBER(F${r})),E${r}*F${r},"")` ]);
    net.push([`=IF(AND(ISNUMBER(H${r}),ISNUMBER(G${r})),H${r}-(MAX(0,E${r}-D${r})*F${r}*G${r}),"")` ]);
  }
  sheet.getRange(7, 6, NUM_ROWS, 1).setFormulas(exch);
  sheet.getRange(7, 8, NUM_ROWS, 1).setFormulas(gross);
  sheet.getRange(7, 9, NUM_ROWS, 1).setFormulas(net);

  const lastDataRow = 6 + NUM_ROWS;
  sheet.getRange(7, 4, NUM_ROWS, 1).setNumberFormat("#,##0.00");       
  sheet.getRange(7, 5, NUM_ROWS, 1).setNumberFormat("#,##0.00");       
  sheet.getRange(7, 6, NUM_ROWS, 1).setNumberFormat("0.0000");          
  sheet.getRange(7, 7, NUM_ROWS, 1).setNumberFormat("0.00%");           
  sheet.getRange(7, 8, NUM_ROWS, 1).setNumberFormat('"$"#,##0.00');    
  sheet.getRange(7, 9, NUM_ROWS, 1).setNumberFormat('"$"#,##0.00');    

  const assetClassRange = sheet.getRange(7, 2, NUM_ROWS, 1);
  const cfRules = Object.entries(THEME.assetRows).map(([cls, bg]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(cls).setBackground(bg).setFontColor(THEME.assetText)
      .setRanges([assetClassRange]).build()
  );
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(THEME.negativeValueBg).setFontColor(THEME.negativeValueFg)
      .setRanges([sheet.getRange(7, 9, NUM_ROWS, 1)]).build()
  );
  sheet.setConditionalFormatRules(cfRules);

  sheet.setColumnWidth(1, 220);  
  sheet.setColumnWidth(2, 135);  
  sheet.setColumnWidth(3, 90);   
  sheet.setColumnWidth(4, 130);  
  sheet.setColumnWidth(5, 130);  
  sheet.setColumnWidth(6, 160);  
  sheet.setColumnWidth(7, 90);   
  sheet.setColumnWidth(8, 150);  
  sheet.setColumnWidth(9, 150);  
  sheet.setColumnWidth(10, 80);  
  sheet.setColumnWidth(11, 260); 
  sheet.setFrozenRows(6);
}

/**
 * 3. Builds the Brokerage Holdings Tab 
 */
function buildHoldingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Brokerage Holdings");
  if (!sheet) sheet = ss.insertSheet("Brokerage Holdings");
  else sheet.clear(); 

  sheet.setHiddenGridlines(true);
  
  const headers = ["Account Name", "Ticker Symbol", "Quantity", "Live Price", "Total Value"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleData = [
    ["Taxable Brokerage", "VTI", 150],
    ["Taxable Brokerage", "AAPL", 50],
    ["401k / RRSP", "QQQ", 200]
  ];
  sheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);

  const numRows = 99;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    let rowNum = i + 2;
    formulas.push([
      `=IF(ISBLANK(B${rowNum}), "", GOOGLEFINANCE(B${rowNum}, "price"))`, 
      `=IF(AND(ISNUMBER(C${rowNum}), ISNUMBER(D${rowNum})), C${rowNum} * D${rowNum}, "")` 
    ]);
  }
  sheet.getRange(2, 4, numRows, 1).setFormulas(formulas.map(row => [row[0]]));
  sheet.getRange(2, 5, numRows, 1).setFormulas(formulas.map(row => [row[1]]));

  sheet.getRange("D2:E100").setNumberFormat("$#,##0.00");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(2, 1, numRows, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

/**
 * 4. Builds the Snapshots Tab
 */
function buildSnapshotTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Snapshots");
  if (!sheet) sheet = ss.insertSheet("Snapshots");
  else sheet.clear();

  const headers = [
    "Date", "Net (USD)", "Liquid (USD)", "Locked (USD)", "Gross (USD)", 
    "Net (CAD)", "Net (INR)", "Total RE (USD)", "Value Δ (USD)", "% Growth", "FIRE Progress", "Auto-Insights", "Manual Notes"
  ];
  sheet.setHiddenGridlines(true);
  
  sheet.getRange("A1:M1").setValues([headers]).setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");
  sheet.getRange("A2:M100").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.setColumnWidth(1, 150); 
  sheet.setColumnWidth(12, 350); 
  sheet.setColumnWidth(13, 200); 
  sheet.setFrozenRows(1);
}

/**
 * 5. Builds the Cash Flow & Burn Tab
 */
function buildCashFlowTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("💸 Cash Flow & Burn");
  if (!sheet) sheet = ss.insertSheet("💸 Cash Flow & Burn");
  else sheet.clear();

  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:F100").setBackground(THEME.canvas);

  sheet.getRange("A1").setValue("CASH FLOW & BURN RATE SUMMARY")
    .setFontWeight("bold").setFontSize(14).setFontColor(THEME.assetText);

  const kpiLabels = [
    "Average Monthly Burn (USD)",
    "TTM (Trailing 12-Month) Expenses (USD)",
    "Target Monthly FIRE Budget (USD)",
    "Current Safe Withdrawal Rate"
  ];
  sheet.getRange(2, 1, kpiLabels.length, 1).setValues(kpiLabels.map(l => [l]))
    .setFontWeight("bold").setFontColor(THEME.mutedText);

  const kpiFormulas = [
    [`=IFERROR(AVERAGEIF(C9:C10000,">0"),0)`],
    [`=IFERROR(SUMPRODUCT((A9:A10000>=TODAY()-365)*(C9:C10000>0)*(C9:C10000)),0)`],
    [`=IFERROR('Settings & Config'!B12, 20000)`],
    [`=IFERROR((B2*12)/'Dashboard & Ledger'!B2, 0)`]
  ];
  sheet.getRange(2, 2, kpiFormulas.length, 1).setFormulas(kpiFormulas);

  sheet.getRange("B2:B4").setNumberFormat("$#,##0.00");
  sheet.getRange("B5").setNumberFormat("0.00%");

  const kpiCardRange = sheet.getRange(2, 1, kpiLabels.length, 2);
  kpiCardRange.setBackground(THEME.kpiCardBg).setBorder(true, true, true, true, false, false, THEME.accentBar, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A7").setValue("EXPENSE LEDGER")
    .setFontWeight("bold").setFontSize(12).setFontColor(THEME.assetText);

  const headers = ["Date", "Category", "Amount (USD)", "Notes"];
  const headerRange = sheet.getRange(8, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleExpenses = [
    [new Date(), "Housing", 3500, "Mortgage / Rent"],
    [new Date(), "Groceries", 800, "Monthly groceries"],
    [new Date(), "Utilities", 250, "Electricity, internet"],
    [new Date(), "Transport", 400, "Gas, insurance"],
    [new Date(), "Dining & Entertainment", 600, "Restaurants, subscriptions"]
  ];
  const dataRange = sheet.getRange(9, 1, sampleExpenses.length, headers.length);
  dataRange.setValues(sampleExpenses);

  sheet.getRange(9, 1, 200, 1).setNumberFormat("mm/dd/yyyy");
  sheet.getRange(9, 3, 200, 1).setNumberFormat("$#,##0.00");

  sheet.getRange(9, 1, 200, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  sheet.setFrozenRows(8);
  sheet.setColumnWidth(1, 120);  
  sheet.setColumnWidth(2, 200);  
  sheet.setColumnWidth(3, 160);  
  sheet.setColumnWidth(4, 300);  
}
/**
 * Generates professional, full-screen visual dashboards on a dedicated tab.
 */
function updateVisualDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Dashboard & Ledger");
  const snapSheet = ss.getSheetByName("Snapshots");

  if (!ledgerSheet || !snapSheet) return;

  // 1. Setup the Dedicated UI Tab
  let dashboardName = "📊 Insights & Analytics";
  let uiSheet = ss.getSheetByName(dashboardName);
  
  if (!uiSheet) {
    uiSheet = ss.insertSheet(dashboardName, 1); 
  } else {
    const existingCharts = uiSheet.getCharts();
    for (let i = 0; i < existingCharts.length; i++) {
      uiSheet.removeChart(existingCharts[i]);
    }
    uiSheet.clear(); 
  }

  // Format the Dashboard Canvas
  uiSheet.getRange("A1:Z100").setBackground(THEME.canvas); 
  uiSheet.getRange("B2").setValue("PORTFOLIO ANALYTICS & TRENDS").setFontWeight("bold").setFontSize(16).setFontColor(THEME.assetText);
  
  uiSheet.setHiddenGridlines(true);

  // 2. Setup the Hidden Data Backend
  let dataSheet = ss.getSheetByName("ChartData");
  if (!dataSheet) {
    dataSheet = ss.insertSheet("ChartData");
    dataSheet.hideSheet(); 
  } else {
    dataSheet.clear();
  }

  // 3. Aggregate Live Asset Allocation
  const dataRange = ledgerSheet.getRange("A7:J60").getValues();
  const allocMap = {};
  
  for (let i = 0; i < dataRange.length; i++) {
    let assetClass = String(dataRange[i][1]);
    let netVal = Number(dataRange[i][8]); 
    let status = String(dataRange[i][9]); 
    
    if (status === "Active" && !isNaN(netVal) && netVal > 0 && assetClass !== "") { 
      if (!allocMap[assetClass]) allocMap[assetClass] = 0;
      allocMap[assetClass] += netVal;
    }
  }

  const pieData = [["Asset Class", "Net Value"]];
  for (const [key, value] of Object.entries(allocMap)) {
    pieData.push([key, value]);
  }
  dataSheet.getRange(1, 1, pieData.length, 2).setValues(pieData);

  // --- CHART BUILDERS ---

  // A. Asset Allocation (Modern Minimalist Donut)
  const lastAssetRow = ledgerSheet.getLastRow();
  if (lastAssetRow > 6) {
    const pieChart = uiSheet.newChart()
      .asPieChart()
      .addRange(ledgerSheet.getRange(7, 2, lastAssetRow - 6, 1)) // Asset Class
      .addRange(ledgerSheet.getRange(7, 9, lastAssetRow - 6, 1)) // Net Worth (USD)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
      .setOption('title', 'Asset Allocation Framework')
      .setOption('pieHole', 0.45)
      .setOption('colors', THEME.charts.donut)
      .setOption('pieSliceBorderColor', "transparent")
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '10%', top: '15%', width: '80%', height: '70%'})
      .setOption('legend', {position: 'labeled', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('pieSliceText', 'percentage')
      .setOption('pieSliceTextStyle', {color: 'white', fontSize: 11})
      .setPosition(4, 2, 0, 0) // Row 4, Col B
      .build();

    uiSheet.insertChart(pieChart);
  }

  // B. Historical Net Worth (Smooth Area Chart)
  const lastSnapRow = snapSheet.getLastRow();
  if (lastSnapRow > 1) {
    const areaChart = uiSheet.newChart()
      .asAreaChart()
      .addRange(snapSheet.getRange(1, 1, lastSnapRow, 1)) // X: Dates
      .addRange(snapSheet.getRange(1, 2, lastSnapRow, 1)) // Y1: Net USD
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Net Worth Trajectory (USD)')
      .setOption('colors', THEME.charts.area) 
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('curveType', 'function') 
      .setOption('chartArea', {left: '15%', top: '15%', width: '80%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: 'short' })
      .setOption('hAxis', { textStyle: {color: THEME.charts.axisText}, format: 'MMM yyyy' })
      .setOption('legend', {position: 'none'}) 
      .setPosition(4, 7, 0, 0) // Row 4, Col G (Next to Donut)
      .build();

    uiSheet.insertChart(areaChart);

    // C. Liquid vs Locked (Stacked Column Chart)
    const stackedBar = uiSheet.newChart()
      .asColumnChart()
      .addRange(snapSheet.getRange(1, 1, lastSnapRow, 1)) // X: Dates
      .addRange(snapSheet.getRange(1, 3, lastSnapRow, 1)) // Y1: Liquid
      .addRange(snapSheet.getRange(1, 4, lastSnapRow, 1)) // Y2: Locked
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Liquidity Profile: Liquid vs. Locked Assets')
      .setOption('isStacked', true)
      .setOption('colors', THEME.charts.stacked) 
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '10%', top: '15%', width: '85%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: 'short' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setPosition(22, 2, 0, 0) // Row 22, Col B (Below the others)
      .build();

    uiSheet.insertChart(stackedBar);
  }
}
/**
 * Execute a manual snapshot. Calculates deltas, generates insights,
 * then chains cloud backups with transparent status reporting.
 */
function captureSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const logSheet = ss.getSheetByName("Snapshots");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!mainSheet || !logSheet) return;

  const netUSD = mainSheet.getRange("B2").getValue();
  const grossUSD = mainSheet.getRange("B3").getValue();
  const netCAD = mainSheet.getRange("E2").getValue();
  const netINR = mainSheet.getRange("H2").getValue();
  const dataRange = mainSheet.getRange("A7:J80").getValues(); 
  
  let liquidUSD = 0, lockedUSD = 0, totalReUSD = 0;

  for (let i = 0; i < dataRange.length; i++) {
    let assetClass = String(dataRange[i][1]);
    let netVal = Number(dataRange[i][8]); 
    let status = String(dataRange[i][9]); 

    if (status === "Active" && !isNaN(netVal) && netVal !== 0) {
      if (["Cash", "Brokerage", "Crypto", "Receivable"].includes(assetClass)) {
        liquidUSD += netVal;
      } else {
        lockedUSD += netVal; 
      }
      if (assetClass === "Real Estate") totalReUSD += netVal;
    }
  }

  const prevNetUSD = logSheet.getRange(2, 2).getValue(); 
  const prevLiquid = logSheet.getRange(2, 3).getValue();
  let dollarDelta = "", pctGrowth = "", autoInsight = "Initial baseline snapshot established.", fireProgress = netUSD / 3000000; 

  if (prevNetUSD && !isNaN(prevNetUSD)) {
    dollarDelta = netUSD - prevNetUSD;
    pctGrowth = dollarDelta / prevNetUSD;
    let liquidDelta = liquidUSD - prevLiquid;
    
    let formatVal = (val) => "$" + Math.abs(val).toLocaleString('en-US', {maximumFractionDigits:0});
    let trend = dollarDelta >= 0 ? "Increased" : "Decreased";
    let sign = (val) => val >= 0 ? "+" : "-";

    autoInsight = `Net worth ${trend.toLowerCase()} by ${formatVal(dollarDelta)}. Liquid pool ${sign(liquidDelta)}${formatVal(liquidDelta)}.`;
  }

  logSheet.insertRowBefore(2);
  const rowData = [new Date(), netUSD, liquidUSD, lockedUSD, grossUSD, netCAD, netINR, totalReUSD, dollarDelta, pctGrowth, fireProgress, autoInsight, ""];
  logSheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);

  logSheet.getRange(2, 2, 1, 7).setNumberFormat("$#,##0.00"); 
  logSheet.getRange(2, 9).setNumberFormat("[Color10]+$#,##0.00;[Color3]-$#,##0.00"); 
  logSheet.getRange(2, 10).setNumberFormat("[Color10]+0.00%;[Color3]-0.00%"); 
  logSheet.getRange(2, 11).setNumberFormat("0.00%"); 
  
  // --- Cloud Backup Chain with Transparent Status ---
  const gistConfigured = _isGistConfigured(configSheet);
  const gistOk    = gistConfigured ? backupToGitHub(true) : false;
  const driveOk   = backupToGoogleDrive(true);

  // Build transparent status message
  const statusParts = ["✅ Snapshot captured successfully!"];
  
  if (gistConfigured && gistOk) {
    statusParts.push("☁️ GitHub Gist — Synced");
  } else if (gistConfigured && !gistOk) {
    statusParts.push("⚠️ GitHub Gist — Sync failed (check logs)");
  } else {
    statusParts.push("💤 GitHub Gist — Not configured");
  }

  if (driveOk) {
    statusParts.push("📁 Google Drive — Synced");
  } else {
    statusParts.push("💤 Google Drive — Not synced");
  }

  if (!gistConfigured && !driveOk) {
    statusParts.push("\n💡 Tip: Set up cloud backups from the WealthScript menu:\n• 🔐 Setup GitHub Backup\n• 📁 Setup Google Drive Backup");
  }

  SpreadsheetApp.getUi().alert(statusParts.join("\n"));

  updateVisualDashboards(); 
}
/**
 * Builds an enriched backup payload including accounts, dashboard KPIs, and latest snapshot.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Object} Complete backup payload
 */
function _buildEnrichedBackup(ss) {
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const snapSheet = ss.getSheetByName("Snapshots");

  // Account-level data
  const dataRange = mainSheet.getRange("A7:K80").getValues();
  const accounts = _buildLedgerSnapshot(dataRange);

  // Dashboard KPI summary
  const summary = {
    netWorthUSD:   mainSheet.getRange("B2").getValue() || 0,
    grossWorthUSD: mainSheet.getRange("B3").getValue() || 0,
    netWorthSecondary1:  mainSheet.getRange("E2").getValue() || 0,
    grossWorthSecondary1: mainSheet.getRange("E3").getValue() || 0,
    netWorthSecondary2:  mainSheet.getRange("H2").getValue() || 0,
    grossWorthSecondary2: mainSheet.getRange("H3").getValue() || 0,
    liquidNetWorthUSD: mainSheet.getRange("B4").getValue() || 0,
    lockedNetWorthUSD: mainSheet.getRange("E4").getValue() || 0,
    fireProgress:      mainSheet.getRange("I4").getValue() || 0,
  };

  // Latest snapshot row (if exists)
  let latestSnapshot = null;
  if (snapSheet && snapSheet.getLastRow() > 1) {
    const snapRow = snapSheet.getRange(2, 1, 1, 13).getValues()[0];
    latestSnapshot = {
      date:         snapRow[0],
      netUSD:       snapRow[1],
      liquidUSD:    snapRow[2],
      lockedUSD:    snapRow[3],
      grossUSD:     snapRow[4],
      valueDelta:   snapRow[8],
      pctGrowth:    snapRow[9],
      fireProgress: snapRow[10],
      autoInsight:  snapRow[11],
      manualNotes:  snapRow[12]
    };
  }

  return {
    snapshotDate: new Date().toISOString(),
    spreadsheetId: ss.getId(),
    summary,
    latestSnapshot,
    accounts
  };
}

/** Manual trigger: runs both Gist and Drive backups with UI alerts. */
function forceBackup() {
  backupToGitHub(false);
  backupToGoogleDrive(false);
}

/**
 * Pure helper: transforms a raw 2D ledger data range into a structured JSON array.
 * Used by both backupToGitHub() and backupToGoogleDrive().
 * @param {Array<Array<*>>} dataRange - 2D array from sheet.getRange().getValues()
 * @returns {Array<Object>} Structured account objects
 */
function _buildLedgerSnapshot(dataRange) {
  const snapshot = [];
  for (let i = 0; i < dataRange.length; i++) {
    const account = String(dataRange[i][0]);
    if (!account) continue;
    snapshot.push({
      Account:      account,
      AssetClass:   String(dataRange[i][1]),
      Currency:     String(dataRange[i][2]),
      InitialCapital: Number(dataRange[i][3]) || 0,
      CurrentValue:   Number(dataRange[i][4]) || 0,
      ExchangeRate:   Number(dataRange[i][5]) || 0,
      TaxRate:        Number(dataRange[i][6]) || 0,
      GrossWorthUSD:  Number(dataRange[i][7]) || 0,
      NetWorthUSD:    Number(dataRange[i][8]) || 0,
      Status:   String(dataRange[i][9]),
      Remarks:  String(dataRange[i][10])
    });
  }
  return snapshot;
}

/**
 * Checks if GitHub Gist backup is configured in Settings.
 * @param {SpreadsheetApp.Sheet} configSheet
 * @returns {boolean}
 */
function _isGistConfigured(configSheet) {
  if (!configSheet) return false;
  const pat = configSheet.getRange("B6").getValue();
  const gistId = configSheet.getRange("B7").getValue();
  return pat && pat !== "PASTE_GITHUB_TOKEN_HERE" && gistId && gistId !== "PASTE_GIST_ID_HERE";
}

/**
 * Disaster Recovery: Serializes live ledger into enriched JSON and pushes to a private GitHub Gist.
 * Silently skips if not configured (no errors thrown).
 * @param {boolean} silent - If true, suppresses UI alerts on success.
 * @returns {boolean} Whether the backup was attempted and succeeded.
 */
function backupToGitHub(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!_isGistConfigured(configSheet)) {
    // Silently skip — not configured
    return false;
  }

  const githubToken = configSheet.getRange("B6").getValue();
  const gistId = configSheet.getRange("B7").getValue();
  const backupData = _buildEnrichedBackup(ss);

  const payload = {
    "description": "WealthScript Automated Backup",
    "files": {
      "net_worth_backup.json": {
        "content": JSON.stringify(backupData, null, 2)
      }
    }
  };

  const options = {
    "method": "PATCH",
    "headers": {
      "Authorization": "Bearer " + githubToken,
      "Accept": "application/vnd.github.v3+json",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.github.com/gists/" + gistId, options);
    if (response.getResponseCode() === 200) {
      if(!silent) SpreadsheetApp.getUi().alert("✅ GitHub Backup Successful!\n\nYour enriched ledger data has been securely versioned in your private GitHub Gist.");
      return true;
    } else {
      if(!silent) SpreadsheetApp.getUi().alert("❌ GitHub API Error:\n" + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log("GitHub backup error: " + e.message);
    if(!silent) SpreadsheetApp.getUi().alert("❌ Script crashed:\n" + e.message);
    return false;
  }
}

/**
 * Google Drive Backup: serializes the enriched ledger to a dated JSON file.
 * Creates a "WealthScript — Backups" folder in Drive automatically.
 * Silently skips if Drive access fails.
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 * @returns {boolean} Whether the backup succeeded.
 */
function backupToGoogleDrive(silent = false) {
  const FOLDER_NAME = "WealthScript \u2014 Backups";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const backupData = _buildEnrichedBackup(ss);
    const jsonContent = JSON.stringify(backupData, null, 2);

    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(FOLDER_NAME);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
    const fileName = `net_worth_${timestamp}.json`;
    folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `\u2705 Google Drive Backup Successful!\n\nFolder: "${FOLDER_NAME}"\nFile: ${fileName}`
      );
    }
    return true;
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
    return false;
  }
}

/**
 * Helper: Creates a Secret Gist via GitHub API
 */
function autoCreateGist(pat) {
  const payload = {
    "description": "WealthScript Automated Backup",
    "public": false,
    "files": { "net_worth_backup.json": { "content": "{\n  \"status\": \"Initialized\"\n}" } }
  };
  const options = {
    "method": "POST",
    "headers": {
      "Authorization": "Bearer " + pat,
      "Accept": "application/vnd.github.v3+json",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.github.com/gists", options);
    if (response.getResponseCode() === 201) {
      return JSON.parse(response.getContentText()).id;
    } else {
      Logger.log("Failed to create Gist: " + response.getContentText());
      return "";
    }
  } catch (e) {
    Logger.log("Error creating Gist: " + e.message);
    return "";
  }
}
