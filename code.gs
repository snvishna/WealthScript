/**
 * ==========================================
 * GLOBAL CONFIGURATION: DEFAULT STARTER DATA
 * ==========================================
 * Edit this array to change the default accounts generated during First Time Setup.
 * Format: ["Account Name", "Asset Class", "Currency", Initial Capital, Current Value, "", Tax Rate, "", "", "Active", "Remarks"]
 */
const DEFAULT_PORTFOLIO_DATA = [
  // --- Cash & Checking ---
  ["Primary Checking", "Cash", "USD", 0, 5000, "", 0.00, "", "", "Active", "Everyday expenses"],
  ["High Yield Savings", "Cash", "USD", 0, 25000, "", 0.00, "", "", "Active", "Emergency Fund"],
  ["International Bank", "Cash", "CAD", 0, 10000, "", 0.00, "", "", "Active", "Travel funds"],
  
  // --- Investments & Retirement ---
  ["Taxable Brokerage", "Brokerage", "USD", 40000, 55000, "", 0.15, "", "", "Active", "Index Funds"],
  ["401k / RRSP", "Retirement", "USD", 0, 120000, "", 0.20, "", "", "Active", "Pre-tax retirement"],
  ["Crypto Exchange", "Crypto", "USD", 5000, 8000, "", 0.15, "", "", "Active", "BTC/ETH"],
  
  // --- Real Estate ---
  ["Primary Residence", "Real Estate", "USD", 400000, 550000, "", 0.00, "", "", "Active", "House"],
  ["Investment Property 1", "Real Estate", "USD", 250000, 310000, "", 0.15, "", "", "Active", "Rental"],
  
  // --- Liabilities (Enter Current Value as Negative) ---
  ["Primary Credit Card", "Liability", "USD", 0, -1500, "", 0.00, "", "", "Active", "Paid monthly"],
  ["Primary Mortgage", "Liability", "USD", 0, -380000, "", 0.00, "", "", "Active", "Home Loan"]
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
      .addSeparator()
      .addItem('☁️ Force Cloud Backup', 'forceManualBackup')
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
 * 1. Builds the Settings & Config Tab
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

  // --- Draw UI ---
  sheet.getRange("A1").setValue("REAL ESTATE API CONFIG").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A2:B2").setValues([["RapidAPI Key", "PASTE_KEY_HERE"]]).setBackground("#f3f4f6");
  sheet.getRange("A3:B3").setValues([["RapidAPI Host", "real-estate101.p.rapidapi.com"]]).setBackground("#f3f4f6");
  
  sheet.getRange("A5").setValue("CLOUD BACKUP CONFIG (DISASTER RECOVERY)").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A6:B6").setValues([["GitHub PAT (gist scope)", pat]]).setBackground("#f8fafc");
  sheet.getRange("A7:B7").setValues([["GitHub Gist ID", gistId]]).setBackground("#f8fafc");

  sheet.getRange("A10").setValue("REAL ESTATE ZPID MAPPING").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A11:B11").setValues([["Account Name (Must match Dashboard exactly)", "ZPID"]]).setBackground("#1e293b").setFontColor("white").setFontWeight("bold");
  
  const sampleMapping = [
    ["Primary Residence", "12345678"],
    ["Investment Property 1", "87654321"]
  ];
  sheet.getRange(12, 1, sampleMapping.length, 2).setValues(sampleMapping);
  
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
 * 2. Builds the Dashboard & Ledger (Ready for Native Tables)
 */
function buildPortfolioTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) sheet = ss.insertSheet("Dashboard & Ledger");
  else sheet.clear();

  sheet.getRange("A1").setValue("TOTAL NET WORTH DASHBOARD").setFontWeight("bold").setFontSize(14);
  
  sheet.getRange("A2").setValue("Total Net Worth (USD):").setFontWeight("bold");
  sheet.getRange("B2").setFormula('=SUMIFS(I:I, J:J, "Active")').setNumberFormat("$#,##0.00");
  sheet.getRange("A3").setValue("Total Gross Worth (USD):").setFontWeight("bold");
  sheet.getRange("B3").setFormula('=SUMIFS(H:H, J:J, "Active")').setNumberFormat("$#,##0.00");
  
  sheet.getRange("D2").setValue("Total Net Worth (CAD):").setFontWeight("bold");
  sheet.getRange("E2").setFormula('=B2 * IFERROR(GOOGLEFINANCE("CURRENCY:USDCAD"), 1)').setNumberFormat("$#,##0.00");
  sheet.getRange("D3").setValue("Total Gross Worth (CAD):").setFontWeight("bold");
  sheet.getRange("E3").setFormula('=B3 * IFERROR(GOOGLEFINANCE("CURRENCY:USDCAD"), 1)').setNumberFormat("$#,##0.00");
  
  sheet.getRange("G2").setValue("Total Net Worth (INR):").setFontWeight("bold");
  sheet.getRange("H2").setFormula('=B2 * IFERROR(GOOGLEFINANCE("CURRENCY:USDINR"), 1)').setNumberFormat("₹#,##0.00");
  sheet.getRange("G3").setValue("Total Gross Worth (INR):").setFontWeight("bold");
  sheet.getRange("H3").setFormula('=B3 * IFERROR(GOOGLEFINANCE("CURRENCY:USDINR"), 1)').setNumberFormat("₹#,##0.00");
  
  sheet.getRange("A1:H4").setBackground("#f3f4f6");

  const headers = ["Account", "Asset Class", "Currency", "Initial Capital", "Current Value", "Exchange Rate (to USD)", "Tax Rate", "Gross Worth (USD)", "Net Worth (USD)", "Status", "Remarks"];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);

  // Inject Global Data
  sheet.getRange(7, 1, DEFAULT_PORTFOLIO_DATA.length, headers.length).setValues(DEFAULT_PORTFOLIO_DATA);

  const numRowsToFill = 45; 
  const formulas = [];
  for (let i = 0; i < numRowsToFill; i++) {
    let rowNum = i + 7; 
    formulas.push([
      `=IF(ISBLANK(C${rowNum}), "", IF(TRIM(UPPER(C${rowNum}))="USD", 1, IFERROR(GOOGLEFINANCE("CURRENCY:"&TRIM(UPPER(C${rowNum}))&"USD"), "Error")))`, 
      `=IF(AND(ISNUMBER(E${rowNum}), ISNUMBER(F${rowNum})), E${rowNum} * F${rowNum}, "")`, 
      `=IF(AND(ISNUMBER(H${rowNum}), ISNUMBER(G${rowNum})), H${rowNum} - (MAX(0, E${rowNum} - D${rowNum}) * F${rowNum} * G${rowNum}), "")`  
    ]);
  }
  
  sheet.getRange(7, 6, numRowsToFill, 1).setFormulas(formulas.map(row => [row[0]])); 
  sheet.getRange(7, 8, numRowsToFill, 1).setFormulas(formulas.map(row => [row[1]])); 
  sheet.getRange(7, 9, numRowsToFill, 1).setFormulas(formulas.map(row => [row[2]])); 

  sheet.getRange("D7:E55").setNumberFormat("#,##0.00"); 
  sheet.getRange("H7:I55").setNumberFormat("$#,##0.00"); 
  sheet.getRange("G7:G55").setNumberFormat("0.00%"); 
  sheet.getRange("F7:F55").setNumberFormat("0.0000"); 
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(11, 250);
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
  const dataRange = mainSheet.getRange("A7:J60").getValues(); 
  
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
  
  // Chain the Cloud Sync silently
  backupToGitHub(true);
  
  // Update the visual charts with the new data
  updateVisualDashboards(); 
}

/**
 * Manual trigger wrapper for UI
 */
function forceManualBackup() {
  backupToGitHub(false);
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

  const dataRange = sheet.getRange("A7:K60").getValues();
  const backupData = [];

  for (let i = 0; i < dataRange.length; i++) {
    let account = String(dataRange[i][0]);
    if (!account) continue; 

    backupData.push({
      "Account": account,
      "AssetClass": String(dataRange[i][1]),
      "Currency": String(dataRange[i][2]),
      "InitialCapital": Number(dataRange[i][3]) || 0,
      "CurrentValue": Number(dataRange[i][4]) || 0,
      "ExchangeRate": Number(dataRange[i][5]) || 0,
      "TaxRate": Number(dataRange[i][6]) || 0,
      "GrossWorthUSD": Number(dataRange[i][7]) || 0,
      "NetWorthUSD": Number(dataRange[i][8]) || 0,
      "Status": String(dataRange[i][9]),
      "Remarks": String(dataRange[i][10])
    });
  }

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

  const propData = configSheet.getRange("A11:B25").getValues();
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
