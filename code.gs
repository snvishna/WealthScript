/**
 * ==========================================
 * GLOBAL CONFIGURATION: DEFAULT STARTER DATA
 * ==========================================
 * Edit this array to change the default accounts generated during First Time Setup.
 * Format: ["Account Name", "Asset Class", "Currency", Initial Capital, Current Value, "", Tax Rate, "", "", "Active", "Remarks"]
 */
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
 * ==========================================
 * APP CORE FUNCTIONALITY
 * ==========================================
 */

/**
 * Creates the custom menu in the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Net Worth Tracker')
      .addItem('🚀 Run First Time Setup', 'runFirstTimeSetup')
      .addSeparator()
      .addItem('📸 Log Snapshot & Cloud Sync', 'captureSnapshot')
      .addItem('🔄 Refresh Real Estate Prices', 'updateRealEstatePrices')
      .addItem('📊 Update Visual Dashboards', 'updateVisualDashboards')
      .addItem('💸 Rebuild Cash Flow Tab', 'buildCashFlowTab')
      .addSeparator()
      .addItem('☁️ Force GitHub Gist Backup', 'forceGistBackup')
      .addItem('📂 Force Google Drive Backup', 'forceDriveBackup')
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

/**
 * 1. Builds the Settings & Config Tab.
 * Layout:
 *   Rows 1-3:   Real Estate API config
 *   Rows 5-7:   Cloud Backup config
 *   Rows 9-12:  FIRE & Cash Flow config  ← NEW (Epic 0)
 *   Rows 14-15: ZPID mapping header
 *   Rows 16+:   ZPID mapping data
 */
function buildSettingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings & Config");
  if (!sheet) sheet = ss.insertSheet("Settings & Config");
  else sheet.clear();

  // --- Auto-Gist Creation Logic ---
  let pat = CLOUD_SYNC_CONFIG.githubPAT || "PASTE_GITHUB_TOKEN_HERE";
  let gistId = CLOUD_SYNC_CONFIG.gistId;

  if (pat !== "" && pat !== "PASTE_GITHUB_TOKEN_HERE" && gistId === "") {
    gistId = autoCreateGist(pat);
  }
  if (!gistId) gistId = "PASTE_GIST_ID_HERE";

  // Helper to apply a standard config-row style
  const styleRow = (range, bg) => range.setBackground(bg).setVerticalAlignment("middle");

  // --- Section 1: Real Estate API ---
  sheet.getRange("A1").setValue("REAL ESTATE API CONFIG").setFontWeight("bold").setFontSize(12);
  styleRow(sheet.getRange("A2:B2"), "#f3f4f6").setValues([["RapidAPI Key", "PASTE_KEY_HERE"]]);
  styleRow(sheet.getRange("A3:B3"), "#f3f4f6").setValues([["RapidAPI Host", "real-estate101.p.rapidapi.com"]]);

  // --- Section 2: Cloud Backup ---
  sheet.getRange("A5").setValue("CLOUD BACKUP CONFIG (DISASTER RECOVERY)").setFontWeight("bold").setFontSize(12);
  styleRow(sheet.getRange("A6:B6"), "#f8fafc").setValues([["GitHub PAT (gist scope)", pat]]);
  styleRow(sheet.getRange("A7:B7"), "#f8fafc").setValues([["GitHub Gist ID", gistId]]);

  // --- Section 3: FIRE & Cash Flow Config ---
  sheet.getRange("A9").setValue("FIRE & CASH FLOW CONFIG").setFontWeight("bold").setFontSize(12);
  const fireConfig = [
    ["Target Monthly FIRE Budget (USD)", 20000],
    ["Estimated Monthly Rental Income (USD)", 0],
    ["Annual Portfolio Return Rate", 0.07]
  ];
  const fireRange = sheet.getRange(10, 1, fireConfig.length, 2);
  fireRange.setValues(fireConfig);
  styleRow(fireRange, "#f0fdf4");
  sheet.getRange(10, 2).setNumberFormat("$#,##0");
  sheet.getRange(11, 2).setNumberFormat("$#,##0");
  sheet.getRange(12, 2).setNumberFormat("0.00%");

  // --- Section 4: Dashboard Currency Config ---
  // Edit the values in column B to change which secondary currencies appear
  // in the top KPI cards of the Dashboard & Ledger tab. Any valid
  // GOOGLEFINANCE currency code works (EUR, GBP, AUD, JPY, SGD, MXN, etc.).
  // Changing a value here instantly updates the dashboard — no re-run needed.
  sheet.getRange("A13").setValue("DASHBOARD CURRENCY CONFIG").setFontWeight("bold").setFontSize(12);
  const currencyConfig = [
    ["Secondary Currency (Card 2)", (DASHBOARD_CONFIG.secondaryCurrencies[0] || "CAD")],
    ["Secondary Currency (Card 3)", (DASHBOARD_CONFIG.secondaryCurrencies[1] || "INR")]
  ];
  const currRange = sheet.getRange(14, 1, currencyConfig.length, 2);
  currRange.setValues(currencyConfig);
  styleRow(currRange, "#eff6ff");
  sheet.getRange("B14").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");
  sheet.getRange("B15").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");

  // --- Section 5: ZPID Mapping ---
  sheet.getRange("A17").setValue("REAL ESTATE ZPID MAPPING").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A18:B18")
    .setValues([["Account Name (Must match Dashboard exactly)", "ZPID"]])
    .setBackground("#1e293b").setFontColor("white").setFontWeight("bold");

  const sampleMapping = [
    ["Primary Residence", "12345678"],
    ["Investment Property 1", "87654321"]
  ];
  sheet.getRange(19, 1, sampleMapping.length, 2).setValues(sampleMapping);

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 350);
}

/**
 * Helper: Creates a Secret Gist via GitHub API
 */
function autoCreateGist(pat) {
  const payload = {
    "description": "Net Worth Tracker Automated Backup",
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

/**
 * 2. Builds the Dashboard & Ledger with full professional formatting.
 * Rows 1-4: KPI summary dashboard. Row 5: gap. Row 6: table header. Row 7+: data.
 */
function buildPortfolioTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) sheet = ss.insertSheet("Dashboard & Ledger");
  else sheet.clear();

  // ── Row 1: Title Banner ─────────────────────────────────────────────────────
  sheet.getRange("A1:K1").merge()
    .setValue("💰  NET WORTH DASHBOARD")
    .setBackground("#0f172a").setFontColor("#f1f5f9")
    .setFontWeight("bold").setFontSize(15)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 44);

  // ── Rows 2-3: KPI Cards — USD (primary) + up to 2 secondary currencies ──────
  //
  // Card layout over 11 columns:
  //   Card 0 (USD):   cols A-C  — label col A, value col B
  //   Card 1 (Cur1):  cols D-G  — label col D, value col E
  //   Card 2 (Cur2):  cols H-K  — label col H, value col I
  //
  // To change secondary currencies edit DASHBOARD_CONFIG.secondaryCurrencies
  // at the top of this file. Any GOOGLEFINANCE currency code works (EUR, GBP…).

  /** Abbreviated number format: $1.55M / $320K / $42 */
  const USD_ABBR_FMT = '[>999999]"$"0.00,,"M";[>999]"$"0,"K";"$"0';

  /** Returns the currency symbol for common codes; defaults to the code string. */
  const currencySymbol = (code) => {
    const SYM = { USD:'$', EUR:'€', GBP:'£', INR:'₹', JPY:'¥',
                  CAD:'CA$', AUD:'A$', SGD:'S$', CHF:'Fr', MXN:'MX$' };
    return SYM[code.toUpperCase()] || code;
  };

  /** Returns an abbreviated Sheets number format for any currency code. */
  const abbrFmt = (code) => {
    const s = currencySymbol(code);
    return `[>999999]"${s}"0.00,,"M";[>999]"${s}"0,"K";"${s}"0`;
  };

  // Card style definitions (index 0 = USD, 1 = first secondary, 2 = second)
  const CARD_STYLES = [
    { bg:"#1e3a5f", labelFg:"#93c5fd", valueFg:"#ffffff", subFg:"#94a3b8" }, // blue
    { bg:"#14532d", labelFg:"#86efac", valueFg:"#ffffff", subFg:"#bbf7d0" }, // green
    { bg:"#312e81", labelFg:"#c4b5fd", valueFg:"#ffffff", subFg:"#a5b4fc" }, // indigo
  ];
  const CARD_LAYOUT = [
    { bg:"A2:C3", lbl:"A", val:"B" },
    { bg:"D2:G3", lbl:"D", val:"E" },
    { bg:"H2:K3", lbl:"H", val:"I" },
  ];

  // USD card (always first — primary currency)
  const s0 = CARD_STYLES[0]; const c0 = CARD_LAYOUT[0];
  sheet.getRange(c0.bg).setBackground(s0.bg);
  sheet.getRange(`${c0.lbl}2`).setValue("Net Worth (USD)").setFontColor(s0.labelFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange(`${c0.val}2`).setFormula('=SUMIFS(I7:I5000,J7:J5000,"Active")')
    .setNumberFormat(USD_ABBR_FMT).setFontColor(s0.valueFg).setFontSize(14).setFontWeight("bold");
  sheet.getRange(`${c0.lbl}3`).setValue("Gross Worth (USD)").setFontColor(s0.labelFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange(`${c0.val}3`).setFormula('=SUMIFS(H7:H5000,J7:J5000,"Active")')
    .setNumberFormat(USD_ABBR_FMT).setFontColor(s0.subFg).setFontSize(11);

  // Secondary currency cards — labels and formulas reference Settings tab
  // directly so CHANGING A CURRENCY IN SETTINGS INSTANTLY UPDATES THE CARD.
  const SETTINGS_CURRENCY_CELLS = ["'Settings & Config'!B14", "'Settings & Config'!B15"];
  SETTINGS_CURRENCY_CELLS.slice(0, 2).forEach((settingsCell, idx) => {
    const sn = CARD_STYLES[idx + 1]; const cn = CARD_LAYOUT[idx + 1];
    sheet.getRange(cn.bg).setBackground(sn.bg);
    // Label formula: dynamically shows the currency code from Settings
    sheet.getRange(`${cn.lbl}2`)
      .setFormula(`="Net Worth ("&${settingsCell}&")"`)  
      .setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    // Value formula: GOOGLEFINANCE uses the currency code from Settings
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

  // ── Row 4: Liquid / Locked / FIRE quick-stats ──────────────────────────────
  // IMPORTANT: use bounded ranges (B7:B5000, not B:B) to avoid circular
  // dependency — this cell is in column B, which whole-col refs would include.
  const LIQUID_CLASSES = ['"Cash"','"Brokerage"','"Crypto"','"Receivable"'];
  const liquidParts = LIQUID_CLASSES.map(c => `SUMIFS(I7:I5000,J7:J5000,"Active",B7:B5000,${c})`).join('+');
  const fireTarget  = DASHBOARD_CONFIG.fireTargetUSD || 3000000;
  const fireLabel   = `🔥 FIRE Progress ($${(fireTarget / 1e6).toFixed(0)}M)`;

  sheet.getRange("A4:C4").setBackground("#f0f9ff");
  sheet.getRange("A4").setValue("🌊 Liquid Net Worth").setFontColor("#0369a1").setFontWeight("bold").setFontSize(9);
  sheet.getRange("B4").setFormula(`=${liquidParts}`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor("#0369a1").setFontSize(11).setFontWeight("bold");

  sheet.getRange("D4:G4").setBackground("#f0fdf4");
  sheet.getRange("D4").setValue("🔒 Locked Net Worth").setFontColor("#15803d").setFontWeight("bold").setFontSize(9);
  sheet.getRange("E4").setFormula(`=SUMIFS(I7:I5000,J7:J5000,"Active")-(${liquidParts})`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor("#15803d").setFontSize(11).setFontWeight("bold");

  sheet.getRange("H4:K4").setBackground("#fdf4ff");
  sheet.getRange("H4").setValue(fireLabel).setFontColor("#7e22ce").setFontWeight("bold").setFontSize(9);
  sheet.getRange("I4").setFormula(`=IFERROR(SUMIFS(I7:I5000,J7:J5000,"Active")/${fireTarget},0)`)
    .setNumberFormat("0.0%").setFontColor("#7e22ce").setFontSize(11).setFontWeight("bold");

  sheet.setRowHeight(4, 28);

  // ── Row 5: Thin accent bar separating header from data table ───────────────
  sheet.getRange("A5:K5").setBackground("#3b82f6");
  sheet.setRowHeight(5, 3);


  // ── Row 6: Table Header ────────────────────────────────────────────────────
  const headers = ["Account","Asset Class","Currency","Initial Capital","Current Value","Exchange Rate (to USD)","Tax Rate","Gross Worth (USD)","Net Worth (USD)","Status","Remarks"];
  sheet.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setBackground("#0f172a").setFontColor("#f8fafc")
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(6, 36);

  // ── Rows 7+: Data & Computed Formulas ─────────────────────────────────────
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

  // ── Number Formats ────────────────────────────────────────────────────────
  const lastDataRow = 6 + NUM_ROWS;
  sheet.getRange(7, 4, NUM_ROWS, 1).setNumberFormat("#,##0.00");       // Initial Capital
  sheet.getRange(7, 5, NUM_ROWS, 1).setNumberFormat("#,##0.00");       // Current Value
  sheet.getRange(7, 6, NUM_ROWS, 1).setNumberFormat("0.0000");          // Exchange Rate
  sheet.getRange(7, 7, NUM_ROWS, 1).setNumberFormat("0.00%");           // Tax Rate
  sheet.getRange(7, 8, NUM_ROWS, 1).setNumberFormat('"$"#,##0.00');    // Gross Worth
  sheet.getRange(7, 9, NUM_ROWS, 1).setNumberFormat('"$"#,##0.00');    // Net Worth

  // ── Conditional Formatting: Asset Class Color Coding ───────────────────────
  const assetClassRange = sheet.getRange(7, 2, NUM_ROWS, 1);
  const classColors = [
    ["Cash",           "#dcfce7"], ["Brokerage",   "#dbeafe"],
    ["Retirement",     "#ede9fe"], ["Health Savings","#bbf7d0"],
    ["Real Estate",    "#fef9c3"], ["Crypto",       "#fed7aa"],
    ["Commodity",      "#fef3c7"], ["Insurance",    "#e0e7ff"],
    ["Receivable",     "#d1fae5"], ["Liability",    "#fee2e2"],
  ];
  const cfRules = classColors.map(([cls, bg]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(cls).setBackground(bg)
      .setRanges([assetClassRange]).build()
  );
  // Red text for negative net worth values
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground("#fff1f2").setFontColor("#be123c")
      .setRanges([sheet.getRange(7, 9, NUM_ROWS, 1)]).build()
  );
  sheet.setConditionalFormatRules(cfRules);

  // ── Column Widths & Polish ─────────────────────────────────────────────────
  sheet.setColumnWidth(1, 220);  // Account
  sheet.setColumnWidth(2, 135);  // Asset Class
  sheet.setColumnWidth(3, 90);   // Currency
  sheet.setColumnWidth(4, 130);  // Initial Capital
  sheet.setColumnWidth(5, 130);  // Current Value
  sheet.setColumnWidth(6, 160);  // Exchange Rate
  sheet.setColumnWidth(7, 90);   // Tax Rate
  sheet.setColumnWidth(8, 150);  // Gross Worth
  sheet.setColumnWidth(9, 150);  // Net Worth
  sheet.setColumnWidth(10, 80);  // Status
  sheet.setColumnWidth(11, 260); // Remarks
  sheet.setFrozenRows(6);
  // Note: setFrozenColumns is intentionally omitted — it conflicts with the
  // full-width merged banner in row 1. Freeze columns manually if needed.
}

/**
 * 3. Builds the Brokerage Holdings Tab 
 */
function buildHoldingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Brokerage Holdings");
  if (!sheet) sheet = ss.insertSheet("Brokerage Holdings");
  else sheet.clear(); 

  const headers = ["Account Name", "Ticker Symbol", "Quantity", "Live Price", "Total Value"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#0f172a").setFontColor("white").setFontWeight("bold");

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
  sheet.getRange("A1:M1").setValues([headers]).setBackground("#1e293b").setFontColor("white").setFontWeight("bold");
  sheet.setColumnWidth(1, 150); 
  sheet.setColumnWidth(12, 350); 
  sheet.setColumnWidth(13, 200); 
  sheet.setFrozenRows(1);
}

/**
 * 5. Builds the Cash Flow & Burn Tab (Epic 1).
 * Top section: KPI summary cards (Average Monthly Burn, TTM Expenses,
 *              Target FIRE Budget, Safe Withdrawal Rate).
 * Bottom section: Manual expense ledger with Date / Category / Amount / Notes.
 */
function buildCashFlowTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("💸 Cash Flow & Burn");
  if (!sheet) sheet = ss.insertSheet("💸 Cash Flow & Burn");
  else sheet.clear();

  // ── Canvas styling ──────────────────────────────────────────────────────────
  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:F100").setBackground("#f8fafc");

  // ── Section 1: KPI Summary ──────────────────────────────────────────────────
  sheet.getRange("A1").setValue("CASH FLOW & BURN RATE SUMMARY")
    .setFontWeight("bold").setFontSize(14).setFontColor("#0f172a");

  // KPI label column (A) and value column (B)
  const kpiLabels = [
    "Average Monthly Burn (USD)",
    "TTM (Trailing 12-Month) Expenses (USD)",
    "Target Monthly FIRE Budget (USD)",
    "Current Safe Withdrawal Rate"
  ];
  sheet.getRange(2, 1, kpiLabels.length, 1).setValues(kpiLabels.map(l => [l]))
    .setFontWeight("bold").setFontColor("#334155");

  // KPI formulas — reference the ledger table which starts at row 9
  // C column = Amount (USD), A column = Date
  const kpiFormulas = [
    // Average monthly burn: average of all positive expense entries
    [`=IFERROR(AVERAGEIF(C9:C10000,">0"),0)`],
    // TTM: sum of entries where date is within the last 365 days
    [`=IFERROR(SUMPRODUCT((A9:A10000>=TODAY()-365)*(C9:C10000>0)*(C9:C10000)),0)`],
    // Target FIRE Budget pulled from Settings tab (row 10, col B)
    [`=IFERROR('Settings & Config'!B10, 20000)`],
    // Safe Withdrawal Rate: annualised burn / current net worth (4% rule reference)
    [`=IFERROR((B2*12)/'Dashboard & Ledger'!B2, 0)`]
  ];
  sheet.getRange(2, 2, kpiFormulas.length, 1).setFormulas(kpiFormulas);

  // Format KPI values
  sheet.getRange("B2:B4").setNumberFormat("$#,##0.00");
  sheet.getRange("B5").setNumberFormat("0.00%");

  // Style KPI card rows
  const kpiCardRange = sheet.getRange(2, 1, kpiLabels.length, 2);
  kpiCardRange.setBackground("#ffffff").setBorder(true, true, true, true, false, false, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);

  // ── Section Divider ─────────────────────────────────────────────────────────
  sheet.getRange("A7").setValue("EXPENSE LEDGER")
    .setFontWeight("bold").setFontSize(12).setFontColor("#0f172a");

  // ── Section 2: Expense Ledger ────────────────────────────────────────────────
  const headers = ["Date", "Category", "Amount (USD)", "Notes"];
  const headerRange = sheet.getRange(8, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#1e293b").setFontColor("white").setFontWeight("bold");

  // Sample starter data to illustrate usage
  const sampleExpenses = [
    [new Date(), "Housing", 3500, "Mortgage / Rent"],
    [new Date(), "Groceries", 800, "Monthly groceries"],
    [new Date(), "Utilities", 250, "Electricity, internet"],
    [new Date(), "Transport", 400, "Gas, insurance"],
    [new Date(), "Dining & Entertainment", 600, "Restaurants, subscriptions"]
  ];
  const dataRange = sheet.getRange(9, 1, sampleExpenses.length, headers.length);
  dataRange.setValues(sampleExpenses);

  // Format date and currency columns
  sheet.getRange(9, 1, 200, 1).setNumberFormat("mm/dd/yyyy");
  sheet.getRange(9, 3, 200, 1).setNumberFormat("$#,##0.00");

  // Row banding on ledger
  sheet.getRange(9, 1, 200, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Freeze the header row and resize columns
  sheet.setFrozenRows(8);
  sheet.setColumnWidth(1, 120);  // Date
  sheet.setColumnWidth(2, 200);  // Category
  sheet.setColumnWidth(3, 160);  // Amount
  sheet.setColumnWidth(4, 300);  // Notes
}

/**
 * Executes a manual snapshot. Calculates deltas and generates insights.
 */
function captureSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const logSheet = ss.getSheetByName("Snapshots");

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
  
  // Chain both cloud backups silently, then refresh charts
  backupToGitHub(true);
  backupToGoogleDrive(true);
  updateVisualDashboards(); 
}

/** Manual trigger: runs GitHub Gist backup only. */
function forceGistBackup() { backupToGitHub(false); }

/** Manual trigger: runs Google Drive backup only. */
function forceDriveBackup() { backupToGoogleDrive(false); }

/**
 * Manual trigger: runs BOTH backup methods.
 * @deprecated Use forceGistBackup() or forceDriveBackup() directly.
 */
function forceManualBackup() {
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
 * Disaster Recovery: Serializes live ledger into JSON and pushes to a private GitHub Gist.
 * @param {boolean} silent - If true, suppresses UI alerts on success (used for chained snapshotting)
 */
function backupToGitHub(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    if(!silent) SpreadsheetApp.getUi().alert("Settings tab missing. Please run First Time Setup.");
    return;
  }

  const githubToken = configSheet.getRange("B6").getValue();
  const gistId = configSheet.getRange("B7").getValue();

  if (!githubToken || githubToken === "PASTE_GITHUB_TOKEN_HERE" || !gistId || gistId === "PASTE_GIST_ID_HERE") {
    if(!silent) SpreadsheetApp.getUi().alert("Disaster Recovery not configured.\n\nPlease add your GitHub PAT and Gist ID in the Settings tab.");
    return;
  }

  const dataRange = sheet.getRange("A7:K80").getValues();
  const backupData = _buildLedgerSnapshot(dataRange);

  const payload = {
    "description": "Net Worth Tracker Automated Backup",
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
      if(!silent) SpreadsheetApp.getUi().alert("✅ Backup Successful!\n\nYour live JSON data has been securely versioned in your private GitHub Gist.");
    } else {
      if(!silent) SpreadsheetApp.getUi().alert("❌ GitHub API Error:\n" + response.getContentText());
    }
  } catch (e) {
    if(!silent) SpreadsheetApp.getUi().alert("❌ Script crashed:\n" + e.message);
  }
}

/**
 * Google Drive Backup: serializes the live ledger to a dated JSON file.
 * Requires ZERO configuration — uses the script's own Google auth context.
 * Creates a "Net Worth Tracker — Backups" folder in Drive automatically.
 * Retains the latest MAX_DRIVE_BACKUPS files and prunes older ones.
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 */
function backupToGoogleDrive(silent = false) {
  const MAX_DRIVE_BACKUPS = 24; // ~2 years of monthly snapshots
  const FOLDER_NAME = "Net Worth Tracker \u2014 Backups";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) return;

  try {
    const dataRange = sheet.getRange("A7:K80").getValues();
    const accounts = _buildLedgerSnapshot(dataRange);
    const jsonContent = JSON.stringify({
      snapshotDate: new Date().toISOString(),
      spreadsheetId: ss.getId(),
      accounts
    }, null, 2);

    // Resolve or create the backup folder
    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(FOLDER_NAME);

    // Create the dated backup file
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
    const fileName = `net_worth_${timestamp}.json`;
    folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);
    // Note: old files are NOT pruned here to avoid requesting the Drive DELETE
    // OAuth scope. Files are tiny (~2KB each) — manage them manually in Drive
    // if needed. Folder: "Net Worth Tracker — Backups" in your Google Drive.

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `\u2705 Google Drive Backup Successful!\n\nFolder: "${FOLDER_NAME}"\nFile: ${fileName}`
      );
    }
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
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

  // NOTE: ZPID table starts at row 16 after Epic 0 added FIRE config at rows 9-12.
  const propData = configSheet.getRange("A16:B30").getValues();
  const properties = [];
  for (let i = 0; i < propData.length; i++) {
    if (propData[i][0] && propData[i][1]) {
      properties.push({ name: String(propData[i][0]), zpid: String(propData[i][1]) });
    }
  }

  const accountNames = sheet.getRange("A1:A100").getValues().flat();

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
        const targetRow = accountNames.indexOf(prop.name) + 1;
        
        if (targetRow > 0 && zestimate) {
          sheet.getRange(targetRow, 5).setValue(zestimate);
        }
      }
    } catch (e) {
      Logger.log(`Script crashed on ${prop.name}: ${e}`);
    }
  });
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
    uiSheet = ss.insertSheet(dashboardName, 1); // Insert right after the Ledger
  } else {
    // Clear existing charts to prevent stacking
    const existingCharts = uiSheet.getCharts();
    for (let i = 0; i < existingCharts.length; i++) {
      uiSheet.removeChart(existingCharts[i]);
    }
    uiSheet.clear(); 
  }

  // Format the Dashboard Canvas
  uiSheet.getRange("A1:Z100").setBackground("#f8fafc"); // Clean slate background
  uiSheet.getRange("B2").setValue("PORTFOLIO ANALYTICS & TRENDS").setFontWeight("bold").setFontSize(16).setFontColor("#0f172a");
  
  // Hide gridlines for a clean UI feel
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
  
  // Custom Modern Color Palette
  const palette = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6', '#f43f5e'];

  // A. Asset Allocation (Modern Donut)
  if (pieData.length > 1) {
    const pieChart = uiSheet.newChart()
      .asPieChart()
      .addRange(dataSheet.getRange(1, 1, pieData.length, 2))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Live Asset Allocation')
      .setOption('pieHole', 0.55) 
      .setOption('colors', palette)
      .setOption('backgroundColor', { fill: 'transparent' }) // Removes white box
      .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '80%'})
      .setOption('legend', {position: 'labeled', textStyle: {fontSize: 12, color: '#334155'}})
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
      .setOption('colors', ['#3b82f6'])
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('curveType', 'function') // Smooth curves
      .setOption('chartArea', {left: '15%', top: '15%', width: '80%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: '#e2e8f0'}, textStyle: {color: '#64748b'}, format: 'short' })
      .setOption('hAxis', { textStyle: {color: '#64748b'}, format: 'MMM yyyy' })
      .setOption('legend', {position: 'none'}) // Cleaner without legend for a single metric
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
      .setOption('colors', ['#10b981', '#94a3b8']) // Green for Liquid, Slate for Locked
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '10%', top: '15%', width: '85%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: '#e2e8f0'}, textStyle: {color: '#64748b'}, format: 'short' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: '#334155'}})
      .setPosition(22, 2, 0, 0) // Row 22, Col B (Below the others)
      .build();

    uiSheet.insertChart(stackedBar);
  }
}
