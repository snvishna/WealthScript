function buildSettingsTab(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings & Config");
  if (!sheet) sheet = ss.insertSheet("Settings & Config");
  else sheet.clear();

  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:C100").setBackground(THEME.canvas);

  // --- Persistent Getting Started Instructions ---
  sheet.getRange("A1:B1").merge()
    .setValue("📋 GETTING STARTED")
    .setFontWeight("bold").setFontSize(13).setFontColor(THEME.titleBanner.bg);
  const steps = [
    ["Step 1", "Replace the sample accounts in 'Dashboard & Ledger' with your real assets"],
    ["Step 2", "Select rows 7+ on Dashboard → Format > Convert to Table"],
    ["Step 3", "Set up cloud backups: WealthScript menu > 🔐 Setup GitHub Backup"],
    ["Step 4", "Set up cloud backups: WealthScript menu > 📁 Setup Google Drive Backup"],
    ["Step 5", "Take your first snapshot: WealthScript menu > 📸 Log Snapshot & Cloud Sync"],
  ];
  const stepsRange = sheet.getRange(2, 1, steps.length, 2);
  stepsRange.setValues(steps);
  stepsRange.setBackground("#EFF6FF").setFontColor(THEME.assetText).setFontSize(10);
  sheet.getRange(2, 1, steps.length, 1).setFontWeight("bold").setFontColor(THEME.accentBlue);

  let pat = CLOUD_SYNC_CONFIG.githubPAT || "PASTE_GITHUB_TOKEN_HERE";
  let gistId = CLOUD_SYNC_CONFIG.gistId;

  if (pat !== "" && pat !== "PASTE_GITHUB_TOKEN_HERE" && gistId === "") {
    gistId = autoCreateGist(pat);
  }
  if (!gistId) gistId = "PASTE_GIST_ID_HERE";

  const styleRow = (range, bg) => range.setBackground(bg).setVerticalAlignment("middle");

  sheet.getRange("A8").setValue("REAL ESTATE API CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  styleRow(sheet.getRange("A9:B9"), THEME.kpiCardBg).setValues([["RapidAPI Key", "PASTE_KEY_HERE"]]);
  styleRow(sheet.getRange("A10:B10"), THEME.kpiCardBg).setValues([["RapidAPI Host", "real-estate101.p.rapidapi.com"]]);

  sheet.getRange("A12").setValue("CLOUD BACKUP CONFIG (DISASTER RECOVERY)").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  styleRow(sheet.getRange("A13:B13"), THEME.kpiCardBg).setValues([["GitHub PAT (gist scope)", pat]]);
  styleRow(sheet.getRange("A14:B14"), THEME.kpiCardBg).setValues([["GitHub Gist ID", gistId]]);
  styleRow(sheet.getRange("A15:B15"), THEME.kpiCardBg).setValues([["GitHub Gist URL", "Run '🔐 Setup GitHub Backup' from the menu"]]);
  sheet.getRange("A15").setFontColor(THEME.mutedText);
  sheet.getRange("B15").setFontColor(THEME.accentBlue);
  styleRow(sheet.getRange("A16:B16"), THEME.kpiCardBg).setValues([["Google Drive Backup Folder", "Run '📁 Setup Google Drive Backup' from the menu"]]);
  sheet.getRange("A16").setFontColor(THEME.mutedText);
  sheet.getRange("B16").setFontColor(THEME.accentBlue);

  sheet.getRange("A18").setValue("FIRE & CASH FLOW CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  const fireConfig = [
    ["Target Monthly FIRE Budget (USD)", 20000],
    ["Estimated Monthly Rental Income (USD)", 0],
    ["Annual Portfolio Return Rate", 0.07]
  ];
  const fireRange = sheet.getRange(19, 1, fireConfig.length, 2);
  fireRange.setValues(fireConfig);
  styleRow(fireRange, THEME.kpiCardBg);
  sheet.getRange(19, 2).setNumberFormat("$#,##0");
  sheet.getRange(20, 2).setNumberFormat("$#,##0");
  sheet.getRange(21, 2).setNumberFormat("0.00%");

  styleRow(sheet.getRange("A22:B22"), THEME.kpiCardBg)
    .setValues([["FIRE Target Net Worth (USD)", (DASHBOARD_CONFIG.fireTargetUSD || 3000000)]]);
  sheet.getRange("B22").setNumberFormat("$#,##0");
  sheet.getRange("B22").setNote("The dashboard FIRE Progress card and the Time-to-FIRE chart both read this cell live. Change it any time — no rebuild needed.");

  sheet.getRange("A23").setValue("DASHBOARD CURRENCY CONFIG").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  const currencyConfig = [
    ["Secondary Currency (Card 2)", (DASHBOARD_CONFIG.secondaryCurrencies[0] || "CAD")],
    ["Secondary Currency (Card 3)", (DASHBOARD_CONFIG.secondaryCurrencies[1] || "INR")]
  ];
  const currRange = sheet.getRange(24, 1, currencyConfig.length, 2);
  currRange.setValues(currencyConfig);
  styleRow(currRange, THEME.kpiCardBg);
  sheet.getRange("B24").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");
  sheet.getRange("B25").setNote("Examples: CAD, EUR, GBP, AUD, JPY, SGD, INR, MXN, CHF");

  sheet.getRange("A27").setValue("REAL ESTATE ZPID MAPPING").setFontWeight("bold").setFontSize(12).setFontColor(THEME.headerBg);
  sheet.getRange("A28:B28")
    .setValues([["Account Name (Must match Dashboard exactly)", "ZPID"]])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleMapping = [
    ["Primary Residence", "12345678"],
    ["Investment Property 1", "87654321"]
  ];
  sheet.getRange(29, 1, sampleMapping.length, 2).setValues(sampleMapping);

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 350);
}

/**
 * 2. Builds the Dashboard & Ledger with full professional formatting.
 */
/**
 * Pure helper: Generates the SUMPRODUCT formula linking a dashboard row to the Brokerage Holdings tab.
 * Uses N() to coerce empty strings in the Total Value column to 0, preventing #VALUE! errors
 * that IFERROR would silently swallow (returning 0 instead of the actual sum).
 * @param {number} rowNum - The 1-indexed row number on the Dashboard sheet
 * @returns {string} The formula string
 */
function _buildBrokerageFormula(rowNum) {
  return `=SUMPRODUCT(('Brokerage Holdings'!$A$2:$A$200=A${rowNum})*N('Brokerage Holdings'!$F$2:$F$200))`;
}

/**
 * Pure helper: builds a locale-aware abbreviated DISPLAY string formula for a KPI card.
 *
 * Google Sheets number formats can only scale by powers of 1,000 (each trailing comma
 * divides by 1000). Crore (10^7) and Lakh (10^5) are NOT powers of 1000, so they are
 * impossible to express as a number format. The scaling therefore has to happen inside
 * the formula, which means these cells resolve to TEXT, not numbers.
 *
 * Consumers that need the numeric value must read the hidden helper cells (N2/N3/O2/O3)
 * instead of the display cells. See _buildCardHelperFormulas().
 *
 * @param {string} settingsCell - A1 ref holding the currency code, e.g. "'Settings & Config'!B24"
 * @param {string} valueCell - A1 ref holding the numeric value in that currency, e.g. "$N$2"
 * @returns {string} The formula string
 */
function _buildAbbrDisplayFormula(settingsCell, valueCell) {
  const SYM = { USD:'$', EUR:'€', GBP:'£', INR:'₹', JPY:'¥',
                CAD:'CA$', AUD:'A$', SGD:'S$', CHF:'Fr', MXN:'MX$' };
  const symIfs = Object.keys(SYM).map(c => `cur="${c}","${SYM[c]}"`).join(',');
  // Currencies that use the Indian numbering system (lakh / crore).
  const indianIfs = ['INR','PKR','LKR','NPR','BDT'].map(c => `cur="${c}"`).join(',');

  return '=IFERROR(LET('
    + `cur,UPPER(TRIM(${settingsCell})),`
    + `v,${valueCell},`
    + `sym,IFS(${symIfs},TRUE,cur&" "),`
    + 'a,ABS(v),'
    + `IF(OR(${indianIfs}),`
    +   'IFS(a>=10000000,sym&TEXT(v/10000000,"0.#")&"Cr",'
    +      'a>=100000,sym&TEXT(v/100000,"0.#")&"L",'
    +      'TRUE,sym&TEXT(v,"#,##0")),'
    +   'IFS(a>=1000000000,sym&TEXT(v/1000000000,"0.00")&"B",'
    +      'a>=1000000,sym&TEXT(v/1000000,"0.00")&"M",'
    +      'a>=1000,sym&TEXT(v/1000,"0")&"K",'
    +      'TRUE,sym&TEXT(v,"0"))'
    + ')),"—")';
}

function buildPortfolioTracker(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
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

  // Hidden numeric backing cells (columns M/N/O). The visible KPI cards render TEXT
  // because Indian-numbering abbreviations cannot be expressed as a number format;
  // Snapshot.gs and Backup.gs read these numeric cells instead of the display cells.
  const SETTINGS_CURRENCY_CELLS = ["'Settings & Config'!B24", "'Settings & Config'!B25"];
  const HELPER_COLS = ["N", "O"];   // N = card 2, O = card 3
  sheet.getRange("M1").setValue("⚙ INTERNAL — do not edit or delete")
    .setFontColor(THEME.mutedText).setFontSize(8);

  SETTINGS_CURRENCY_CELLS.forEach((settingsCell, idx) => {
    const sn = CARD_STYLES[idx + 1]; const cn = CARD_LAYOUT[idx + 1];
    const hc = HELPER_COLS[idx];
    const rateCell = `$M$${idx + 2}`;

    sheet.getRange(`M${idx + 2}`)
      .setFormula(`=IFERROR(GOOGLEFINANCE("CURRENCY:USD"&TRIM(UPPER(${settingsCell}))),1)`)
      .setNumberFormat("0.000000");
    sheet.getRange(`${hc}2`).setFormula(`=B2*${rateCell}`).setNumberFormat("#,##0.00");
    sheet.getRange(`${hc}3`).setFormula(`=B3*${rateCell}`).setNumberFormat("#,##0.00");

    sheet.getRange(cn.bg).setBackground(sn.bg);
    sheet.getRange(`${cn.lbl}2`)
      .setFormula(`="Net Worth ("&${settingsCell}&")"`)
      .setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}2`)
      .setFormula(_buildAbbrDisplayFormula(settingsCell, `$${hc}$2`))
      .setNumberFormat("@").setFontColor(sn.valueFg).setFontSize(14).setFontWeight("bold");
    sheet.getRange(`${cn.lbl}3`)
      .setFormula(`="Gross Worth ("&${settingsCell}&")"`)
      .setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}3`)
      .setFormula(_buildAbbrDisplayFormula(settingsCell, `$${hc}$3`))
      .setNumberFormat("@").setFontColor(sn.subFg).setFontSize(11);
  });
  sheet.hideColumns(13, 3);   // M, N, O

  sheet.setRowHeight(2, 38); sheet.setRowHeight(3, 28);

  const LIQUID_CLASSES = ['"Cash"','"Brokerage"','"Crypto"','"Receivable"'];
  const liquidParts = LIQUID_CLASSES.map(c => `SUMIFS(I7:I5000,J7:J5000,"Active",B7:B5000,${c})`).join('+');
  const FIRE_TARGET_CELL = "'Settings & Config'!$B$22";

  sheet.getRange("A4:C4").setBackground(THEME.quickStats.liquidBg);
  sheet.getRange("A4").setValue("🌊 Liquid Net Worth").setFontColor(THEME.quickStats.liquidFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("B4").setFormula(`=${liquidParts}`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.liquidFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("D4:G4").setBackground(THEME.quickStats.lockedBg);
  sheet.getRange("D4").setValue("🔒 Locked Net Worth").setFontColor(THEME.quickStats.lockedFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("E4").setFormula(`=SUMIFS(I7:I5000,J7:J5000,"Active")-(${liquidParts})`)
    .setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.lockedFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("H4:K4").setBackground(THEME.quickStats.fireBg);
  sheet.getRange("H4").setFormula(
      `="🔥 FIRE Progress ("&IF(${FIRE_TARGET_CELL}>=1000000,"$"&TEXT(${FIRE_TARGET_CELL}/1000000,"0.##")&"M","$"&TEXT(${FIRE_TARGET_CELL},"#,##0"))&")"`)
    .setFontColor(THEME.quickStats.fireFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("I4").setFormula(`=IFERROR(SUMIFS(I7:I5000,J7:J5000,"Active")/${FIRE_TARGET_CELL},0)`)
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
  for (let i = 0; i < DEFAULT_PORTFOLIO_DATA.length; i++) {
    if (DEFAULT_PORTFOLIO_DATA[i][1] === "Brokerage") {
      const r = i + 7;
      const cell = sheet.getRange(r, 5);
      cell.setFormula(_buildBrokerageFormula(r));
      // Warning-only protection: native Google Sheets dialog fires when user tries to edit
      cell.protect()
        .setWarningOnly(true)
        .setDescription("Auto-calculated from Brokerage Holdings tab. Editing will override live market data.");
    }
  }

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

  // --- Current Value UX: Formula cells only get a visual lock indicator ---
  // Only formula-driven cells (e.g., Brokerage SUMIF rows) are styled as muted/italic
  // to signal "do not override." Manual input cells keep default formatting.
  const currentValueRange = sheet.getRange(7, 5, NUM_ROWS, 1);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISFORMULA(E7)')
      .setFontColor(THEME.mutedText)
      .setItalic(true)
      .setRanges([currentValueRange]).build()
  );
  sheet.setConditionalFormatRules(cfRules);

  // Add an instructional note to the 'Current Value' header (Row 6, Col 5)
  sheet.getRange(6, 5).setNote("💡 Muted italic = Auto-calculated from Holdings tab (do not manually edit).\n\nAll other rows: type your current balance here.");

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
function buildHoldingsTab(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Brokerage Holdings");
  if (!sheet) sheet = ss.insertSheet("Brokerage Holdings");
  else sheet.clear(); 

  sheet.setHiddenGridlines(true);
  
  const headers = ["Account Name", "Asset Category", "Ticker Symbol", "Quantity", "Live Price", "Total Value"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleData = [
    ["Taxable Brokerage", "US Equity", "VTI", 150],
    ["Taxable Brokerage", "Individual Stocks", "AAPL", 50],
    ["401k / RRSP", "Technology", "QQQ", 200]
  ];
  sheet.getRange(2, 1, sampleData.length, 4).setValues(sampleData);

  const numRows = 99;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    let rowNum = i + 2;
    formulas.push([
      `=IF(ISBLANK(C${rowNum}), "", GOOGLEFINANCE(C${rowNum}, "price"))`, 
      `=IF(AND(ISNUMBER(D${rowNum}), ISNUMBER(E${rowNum})), D${rowNum} * E${rowNum}, "")` 
    ]);
  }
  sheet.getRange(2, 5, numRows, 1).setFormulas(formulas.map(row => [row[0]]));
  sheet.getRange(2, 6, numRows, 1).setFormulas(formulas.map(row => [row[1]]));

  sheet.getRange("E2:F100").setNumberFormat("$#,##0.00");
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
    "Net (CAD)", "Net (INR)", "Total RE (USD)", 
    "Cash (USD)", "Brokerage (USD)", "Retirement (USD)", "Liabilities (USD)", 
    "Value Δ (USD)", "% Growth", "FIRE Progress", "Auto-Insights", "Manual Notes"
  ];
  sheet.setHiddenGridlines(true);
  
  sheet.getRange("A1:Q1").setValues([headers]).setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");
  sheet.getRange("A2:Q100").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.setColumnWidth(1, 150); 
  sheet.setColumnWidth(16, 350); 
  sheet.setColumnWidth(17, 200); 
  sheet.setFrozenRows(1);
}

/**
 * 5. Builds the Cash Flow & Burn Tab
 */
function buildCashFlowTab(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
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
    [`=IFERROR('Settings & Config'!B19, 20000)`],
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
