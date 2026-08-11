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
 * fireTargetUSD only SEEDS the Settings tab (B22) during First Time Setup.
 * After setup, change the target in Settings & Config!B22 — the dashboard reads it live.
 * Supports any valid GOOGLEFINANCE code: EUR, GBP, AUD, JPY, MXN, SGD, etc.
 * Only the first two entries are rendered (layout: USD + 2 secondary cards).
 */
const DASHBOARD_CONFIG = {
  secondaryCurrencies: ["CAD", "INR"], // ← Change these to your currencies
  fireTargetUSD: 10000000,             // ← Seeds Settings!B22 on first setup; edit B22 live thereafter

  // Asset classes whose Current Value is computed from the Brokerage Holdings
  // tab. Any class listed here is monitored for broken links and repaired.
  // Add your own (e.g. "HSA", "529") if you track them the same way.
  holdingsLinkedClasses: ["Brokerage", "Retirement", "Crypto", "Health Savings"],
};


/**
 * Credentials that must not live in a spreadsheet cell.
 *
 * A cell is visible to every collaborator, including view-only ones, and is
 * carried into CSV/XLSX exports. These move to DocumentProperties, which is
 * attached to the bound script project rather than to sheet content.
 *
 * `cell` is retained only so an un-migrated workbook keeps working and so
 * migrateSecretsToProperties() knows where to look.
 */
const SECRET_SPECS = [
  { name: "rapidApiKey", prop: "wealthscript.rapidapi.key", cell: "B9",
    label: "RapidAPI Key", placeholder: "PASTE_KEY_HERE" },
  { name: "githubPat", prop: "wealthscript.github.pat", cell: "B13",
    label: "GitHub PAT (gist scope)", placeholder: "PASTE_GITHUB_TOKEN_HERE" }
];

/** Text left in a cell once its secret has moved to secure storage. */
const SECRET_MOVED_NOTICE = "🔒 Stored securely (not in this sheet)";


/* ------------------------------------------------------------------ *
 * EXECUTION CONTEXT
 * ------------------------------------------------------------------ */

/**
 * Resolves the spreadsheet to operate on.
 *
 * Time-driven and installable triggers invoke their handler with an EVENT
 * OBJECT as the first argument. Functions written as `f(ss_inject)` then did
 * `_resolveSpreadsheet(ss_inject)`, took the truthy event,
 * and died on `ss.getSheetByName is not a function`. Duck-type instead of
 * trusting truthiness.
 *
 * @param {*} maybe - a Spreadsheet, a trigger event object, or undefined
 * @returns {SpreadsheetApp.Spreadsheet}
 */
function _resolveSpreadsheet(maybe) {
  return (maybe && typeof maybe.getSheetByName === "function")
    ? maybe
    : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * True when a UI is attached. SpreadsheetApp.getUi() throws in a trigger
 * context, so any alert() there is an uncaught exception rather than a dialog.
 * @returns {boolean}
 */
function _isInteractive() {
  try { return !!SpreadsheetApp.getUi(); } catch (e) { return false; }
}

/**
 * Pure helper: should this run suppress UI?
 * Explicitly-silent callers stay silent; everything else follows the context.
 * @param {boolean} requested
 * @param {boolean} interactive
 * @returns {boolean}
 */
function _resolveSilent(requested, interactive) {
  return requested === true || interactive === false;
}
/**
 * ==========================================
 * FORMULAS — SINGLE SOURCE OF TRUTH
 * ==========================================
 * Every formula WealthScript injects is defined here, exactly once.
 *
 * Both buildPortfolioTracker() (fresh setup) and repairFormulas() (recovery)
 * read from these functions. Defining a formula in two places is how a builder
 * and a repair routine silently drift apart, so don't inline formulas elsewhere.
 *
 * All functions here are PURE — no SpreadsheetApp calls — so they are unit
 * testable without a live spreadsheet.
 */

const LEDGER_FIRST_ROW = 7;
const LEDGER_NUM_ROWS = 70;
const LEDGER_LAST_ROW = LEDGER_FIRST_ROW + LEDGER_NUM_ROWS - 1;   // 76

const HOLDINGS_FIRST_ROW = 2;
const HOLDINGS_NUM_ROWS = 99;
const HOLDINGS_LAST_ROW = HOLDINGS_FIRST_ROW + HOLDINGS_NUM_ROWS - 1;  // 100

const CFG = "'Settings & Config'";
const FIRE_TARGET_CELL = `${CFG}!$B$22`;
const CURRENCY_CELLS = [`${CFG}!B24`, `${CFG}!B25`];

/**
 * Canonical header rows. repairFormulas() and migrateSheetLayout() write by
 * hardcoded A1 reference (I4, M2, ...), so an inserted or deleted column would
 * send every formula to the wrong cell while still reporting success. These
 * are asserted before any write.
 */
const LEDGER_HEADERS = ["Account", "Asset Class", "Currency", "Initial Capital",
  "Current Value", "Exchange Rate (to USD)", "Tax Rate", "Gross Worth (USD)",
  "Net Worth (USD)", "Status", "Remarks"];

const HOLDINGS_HEADERS = ["Account Name", "Asset Category", "Ticker Symbol",
  "Quantity", "Live Price", "Total Value (USD)"];

/**
 * Pure helper: compares an actual header row against the canonical one.
 * @param {Array<*>} actual - values read from the sheet's header row
 * @param {Array<string>} expected - canonical headers
 * @returns {{ok: boolean, problems: Array<string>}}
 */
function _verifyHeaders(actual, expected) {
  const problems = [];
  const col = n => String.fromCharCode(65 + n);

  // A parenthetical annotation is a labelling choice, not a structural change:
  // "Total Value" and "Total Value (USD)" describe the same column. Compare on
  // the text before " (" so users can annotate headers freely, while a genuine
  // mismatch ("Net Worth (USD)" sitting in the Status column) still fails.
  const base = t => String(t === undefined || t === null ? "" : t)
    .trim().replace(/\s*\(.*\)\s*$/, "").trim().toLowerCase();

  for (let i = 0; i < expected.length; i++) {
    const got = String((actual || [])[i] === undefined ? "" : actual[i]).trim();
    if (base(got) !== base(expected[i])) {
      problems.push(`${col(i)} should be "${expected[i]}" but is "${got}"`);
    }
  }
  return { ok: problems.length === 0, problems: problems };
}

/** Asset classes counted as liquid on the dashboard quick-stats row. */
const LIQUID_CLASSES = ["Cash", "Brokerage", "Crypto", "Receivable"];

/**
 * Pure helper: SUMIFS chain totalling the liquid asset classes.
 * @returns {string} formula fragment (no leading '=')
 */
function _liquidSumifs() {
  return LIQUID_CLASSES
    .map(c => `SUMIFS(I7:I5000,J7:J5000,"Active",B7:B5000,"${c}")`)
    .join('+');
}

/**
 * Pure helper: links a dashboard row to the Brokerage Holdings tab.
 * N() coerces empty strings in the Total Value column to 0, preventing #VALUE!
 * errors that IFERROR would silently swallow (returning 0 instead of the sum).
 * @param {number} rowNum - 1-indexed Dashboard row
 * @returns {string}
 */
function _buildBrokerageFormula(rowNum) {
  return `=SUMPRODUCT(('Brokerage Holdings'!$A$2:$A$200=A${rowNum})*N('Brokerage Holdings'!$F$2:$F$200))`;
}

/**
 * Pure helper: locale-aware abbreviated DISPLAY string for a KPI card.
 *
 * Google Sheets number formats scale only by powers of 1,000 (each trailing
 * comma divides by 1000). Crore (10^7) and Lakh (10^5) are NOT powers of 1000,
 * so they cannot be expressed as a number format. Scaling therefore happens
 * inside the formula, which means these cells resolve to TEXT, not numbers.
 *
 * Consumers needing the numeric value must read the hidden helper cells
 * (N2/N3/O2/O3), never the display cells.
 *
 * @param {string} settingsCell - A1 ref holding the currency code
 * @param {string} valueCell - A1 ref holding the numeric value in that currency
 * @returns {string}
 */
function _buildAbbrDisplayFormula(settingsCell, valueCell) {
  const SYM = { USD:'$', EUR:'€', GBP:'£', INR:'₹', JPY:'¥',
                CAD:'CA$', AUD:'A$', SGD:'S$', CHF:'Fr', MXN:'MX$' };
  const symIfs = Object.keys(SYM).map(c => `curr="${c}","${SYM[c]}"`).join(',');
  // Currencies using the Indian numbering system (lakh / crore).
  const indianIfs = ['INR','PKR','LKR','NPR','BDT'].map(c => `curr="${c}"`).join(',');

  return '=IFERROR(LET('
    + `curr,UPPER(TRIM(${settingsCell})),val,${valueCell},`
    + `symb,IFS(${symIfs},TRUE,curr&" "),mag,ABS(val),`
    + `IF(OR(${indianIfs}),`
    +   'IFS(mag>=10000000,symb&TEXT(val/10000000,"0.#")&"Cr",'
    +      'mag>=100000,symb&TEXT(val/100000,"0.#")&"L",'
    +      'TRUE,symb&TEXT(val,"#,##0")),'
    +   'IFS(mag>=1000000000,symb&TEXT(val/1000000000,"0.00")&"B",'
    +      'mag>=1000000,symb&TEXT(val/1000000,"0.00")&"M",'
    +      'mag>=1000,symb&TEXT(val/1000,"0")&"K",'
    +      'TRUE,symb&TEXT(val,"0"))'
    + ')),"—")';
}

/**
 * Every fixed (non-repeating) formula cell on Dashboard & Ledger.
 * @returns {Object<string,string>} map of A1 notation -> formula
 */
function _ledgerFixedFormulas() {
  const liquid = _liquidSumifs();
  const out = {
    "B2": '=SUMIFS(I7:I5000,J7:J5000,"Active")',
    "B3": '=SUMIFS(H7:H5000,J7:J5000,"Active")',
    "B4": `=${liquid}`,
    "E4": `=SUMIFS(I7:I5000,J7:J5000,"Active")-(${liquid})`,
    "H4": `="🔥 FIRE Progress ("&IF(${FIRE_TARGET_CELL}>=1000000,"$"&TEXT(${FIRE_TARGET_CELL}/1000000,"0.##")&"M","$"&TEXT(${FIRE_TARGET_CELL},"#,##0"))&")"`,
    "I4": `=IFERROR(SUMIFS(I7:I5000,J7:J5000,"Active")/${FIRE_TARGET_CELL},0)`
  };

  // Secondary currency cards: label, hidden FX rate, hidden numeric, display text.
  const CARDS = [
    { lbl: "D", val: "E", helper: "N", rate: "M2", rateRow: 2 },
    { lbl: "H", val: "I", helper: "O", rate: "M3", rateRow: 3 }
  ];

  CARDS.forEach((c, idx) => {
    const cur = CURRENCY_CELLS[idx];
    out[c.rate] = `=IFERROR(GOOGLEFINANCE("CURRENCY:USD"&TRIM(UPPER(${cur}))),1)`;
    out[`${c.helper}2`] = `=B2*$M$${c.rateRow}`;
    out[`${c.helper}3`] = `=B3*$M$${c.rateRow}`;
    out[`${c.lbl}2`] = `="Net Worth ("&${cur}&")"`;
    out[`${c.lbl}3`] = `="Gross Worth ("&${cur}&")"`;
    out[`${c.val}2`] = _buildAbbrDisplayFormula(cur, `$${c.helper}$2`);
    out[`${c.val}3`] = _buildAbbrDisplayFormula(cur, `$${c.helper}$3`);
  });

  return out;
}

/**
 * Per-row formulas for the Dashboard ledger grid.
 * Column E is deliberately excluded — it is manual entry except on Brokerage
 * rows, which are handled separately via _buildBrokerageFormula().
 * @param {number} r - 1-indexed row
 * @returns {Object<string,string>} map of column letter -> formula
 */
function _ledgerRowFormulas(r) {
  return {
    F: `=IF(ISBLANK(C${r}),"",IF(TRIM(UPPER(C${r}))="USD",1,IFERROR(GOOGLEFINANCE("CURRENCY:"&TRIM(UPPER(C${r}))&"USD"),"Error")))`,
    H: `=IF(AND(ISNUMBER(E${r}),ISNUMBER(F${r})),E${r}*F${r},"")`,
    I: `=IF(AND(ISNUMBER(H${r}),ISNUMBER(G${r})),H${r}-(MAX(0,E${r}-D${r})*F${r}*G${r}),"")`
  };
}

/**
 * Per-row formulas for the Brokerage Holdings grid.
 * @param {number} r - 1-indexed row
 * @returns {Object<string,string>} map of column letter -> formula
 */
function _holdingsRowFormulas(r) {
  return {
    // "CASH" is itself a live ticker on US markets, so a bare
    // GOOGLEFINANCE(C, "price") on a cash row returns a real equity quote
    // instead of erroring — silently multiplying the balance by ~87x. Reserve
    // it explicitly. Non-USD cash rows override this cell with an FX lookup;
    // repairFormulas() preserves such overrides.
    E: `=IF(ISBLANK(C${r}), "", IF(UPPER(TRIM(C${r}))="CASH", 1, IFERROR(GOOGLEFINANCE(C${r}, "price"), NA())))`,
    F: `=IF(AND(ISNUMBER(D${r}), ISNUMBER(E${r})), D${r} * E${r}, "")`
  };
}

/**
 * Previous canonical price formulas, safe for repairFormulas() to upgrade.
 *
 * Anything NOT in this list and not the current canonical is treated as a
 * deliberate user override and left untouched — an FX lookup on a foreign
 * currency row, a fallback price for an unquotable ticker, a pinned option mark.
 *
 * @param {number} r - 1-indexed row
 * @returns {Array<string>}
 */
function _legacyHoldingsPriceFormulas(r) {
  return [
    `=IF(ISBLANK(C${r}), "", GOOGLEFINANCE(C${r}, "price"))`,
    `=IF(ISBLANK(C${r}),"",GOOGLEFINANCE(C${r},"price"))`,
    `=IFERROR(GOOGLEFINANCE(C${r},"price"),NA())`
  ];
}

/**
 * Ranges whose formula count is tracked for integrity checking.
 * Column E on the ledger is excluded — its formula count legitimately varies
 * with how many Brokerage accounts exist, so it gets a dedicated structural
 * check instead (see _auditBrokerageLinks).
 */
function _managedRanges() {
  return [
    { sheet: "Dashboard & Ledger", a1: "B2:B4",  label: "KPI totals (USD)" },
    { sheet: "Dashboard & Ledger", a1: "D2:E4",  label: "KPI card 2 + Locked Net Worth" },
    { sheet: "Dashboard & Ledger", a1: "H2:I4",  label: "KPI card 3 + FIRE progress" },
    { sheet: "Dashboard & Ledger", a1: "M2:O3",  label: "Hidden FX helper cells" },
    { sheet: "Dashboard & Ledger", a1: `F${LEDGER_FIRST_ROW}:F${LEDGER_LAST_ROW}`, label: "Exchange Rate column" },
    { sheet: "Dashboard & Ledger", a1: `H${LEDGER_FIRST_ROW}:I${LEDGER_LAST_ROW}`, label: "Gross / Net Worth columns" },
    { sheet: "Brokerage Holdings", a1: `F${HOLDINGS_FIRST_ROW}:F${HOLDINGS_LAST_ROW}`, label: "Holdings Total Value column" }
  ];
}
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
 * Uses a SINGLE modal dialog with inline link + embedded token input
 * to avoid the dual-popup overlap issue.
 */
function setupGistWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  const patUrl = "https://github.com/settings/tokens/new?scopes=gist&description=WealthScript+Backup";
  const htmlContent = `
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; color: #1a1a1a; margin: 0; }
      .step { margin-bottom: 10px; display: flex; align-items: flex-start; gap: 10px; }
      .step-num { flex-shrink: 0; background: #2563EB; color: white; border-radius: 50%;
                  width: 24px; height: 24px; text-align: center; line-height: 24px; font-size: 13px; }
      a.btn { display: inline-block; background: #2563EB; color: white !important; padding: 10px 20px;
              border-radius: 6px; text-decoration: none; font-weight: bold; margin: 8px 0 12px; }
      a.btn:hover { background: #1d4ed8; }
      code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
      input[type=text] { width: 100%; padding: 10px; border: 2px solid #e2e8f0; border-radius: 6px;
                          font-size: 14px; font-family: monospace; box-sizing: border-box; margin: 6px 0; }
      input[type=text]:focus { outline: none; border-color: #2563EB; }
      .submit-btn { background: #059669; color: white; border: none; padding: 10px 24px;
                     border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 4px; }
      .submit-btn:hover { background: #047857; }
      .submit-btn:disabled { background: #9ca3af; cursor: not-allowed; }
      .status { margin-top: 10px; padding: 10px; border-radius: 6px; font-size: 13px; display: none; }
      .status.error { background: #fef2f2; color: #be123c; display: block; }
      .status.success { background: #ecfdf5; color: #059669; display: block; }
      .status.loading { background: #eff6ff; color: #2563EB; display: block; }
      hr { border: none; border-top: 1px solid #e2e8f0; margin: 12px 0; }
    </style>

    <div class="step"><span class="step-num">1</span><span>Click the button below to open GitHub's token page:</span></div>
    <a class="btn" href="${patUrl}" target="_blank">🔐 Open GitHub Token Page →</a>
    <div class="step"><span class="step-num">2</span><span>The <code>gist</code> scope is pre-selected — just click <b>"Generate token"</b></span></div>
    <div class="step"><span class="step-num">3</span><span>Copy the token and paste it below:</span></div>
    <hr>
    <input type="text" id="tokenInput" placeholder="ghp_xxxxxxxxxxxxxxxxxxxx" />
    <button class="submit-btn" id="submitBtn" onclick="submitToken()">✅ Connect to GitHub</button>
    <div class="status" id="statusMsg"></div>

    <script>
      function submitToken() {
        var token = document.getElementById('tokenInput').value.trim();
        var btn = document.getElementById('submitBtn');
        var status = document.getElementById('statusMsg');
        
        if (!token) {
          status.className = 'status error';
          status.textContent = 'Please paste your token above.';
          return;
        }
        
        btn.disabled = true;
        btn.textContent = '⏳ Validating...';
        status.className = 'status loading';
        status.textContent = 'Validating token and creating your private Gist...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              status.className = 'status success';
              status.textContent = '✅ ' + result.message;
              btn.textContent = '✅ Done!';
              setTimeout(function() { google.script.host.close(); }, 2500);
            } else {
              status.className = 'status error';
              status.textContent = '❌ ' + result.message;
              btn.disabled = false;
              btn.textContent = '✅ Connect to GitHub';
            }
          })
          .withFailureHandler(function(err) {
            status.className = 'status error';
            status.textContent = '❌ Error: ' + err.message;
            btn.disabled = false;
            btn.textContent = '✅ Connect to GitHub';
          })
          ._processGistToken(token);
      }
    </script>`;

  const htmlOutput = HtmlService
    .createHtmlOutput(htmlContent)
    .setWidth(480)
    .setHeight(380);
  ui.showModalDialog(htmlOutput, "🔐 Setup GitHub Backup");
}

/**
 * Server-side handler called from the wizard HTML dialog.
 * Validates the token, creates a Gist, and populates Settings.
 * @param {string} token - The raw PAT string from the dialog input.
 * @returns {{success: boolean, message: string}}
 */
function _processGistToken(token) {
  const pat = (token || "").trim();

  if (!_validatePATFormat(pat)) {
    return { success: false, message: "Invalid format. Expected a token starting with 'ghp_' or 'github_pat_'." };
  }

  // Validate against GitHub API
  try {
    const testResponse = UrlFetchApp.fetch("https://api.github.com/user", {
      headers: { "Authorization": "Bearer " + pat, "Accept": "application/vnd.github.v3+json" },
      muteHttpExceptions: true
    });
    if (testResponse.getResponseCode() !== 200) {
      return { success: false, message: "GitHub rejected this token. Ensure it has 'gist' scope." };
    }
  } catch (e) {
    return { success: false, message: "Network error: " + e.message };
  }

  // Create Gist
  const gistId = autoCreateGist(pat);
  if (!gistId) {
    return { success: false, message: "Token is valid but Gist creation failed. Check Apps Script logs." };
  }

  // Populate Settings tab
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");
  if (configSheet) {
    // The PAT goes to secure storage, never into a cell.
    setSecret("githubPat", pat, ss);
    configSheet.getRange("B14").setValue(gistId);

    const gistUrl = _buildGistUrl(gistId);
    const richGistLink = SpreadsheetApp.newRichTextValue()
      .setText(gistUrl)
      .setLinkUrl(gistUrl)
      .build();
    configSheet.getRange("B15").setRichTextValue(richGistLink);
  }

  return { success: true, message: `Connected! Gist ID: ${gistId}. Every snapshot will now auto-sync to GitHub.` };
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
    const folder = folderIterator.hasNext()
      ? folderIterator.next()
      : DriveApp.createFolder(FOLDER_NAME);
    const folderUrl = folder.getUrl();

    const richDriveLink = SpreadsheetApp.newRichTextValue()
      .setText(folderUrl)
      .setLinkUrl(folderUrl)
      .build();
    configSheet.getRange("B16").setRichTextValue(richDriveLink);

    ui.alert(`✅ Google Drive Backup Connected!\n\nFolder: "${FOLDER_NAME}"\n\nA clickable link has been added to your Settings tab.\nEvery snapshot will now auto-sync a dated JSON file here.`);
  } catch (e) {
    ui.alert("❌ Drive Setup Failed:\n" + e.message);
  }
}
/**
 * Pure helper: decides which ledger rows should receive a fetched Zestimate.
 *
 * Guards against the three ways a bulk write can corrupt the ledger:
 *   1. Only rows whose Asset Class is exactly "Real Estate" are eligible.
 *   2. Every match is reported, not just the first (indexOf silently ignored duplicates).
 *   3. Non-positive / missing Zestimates are skipped rather than written as 0.
 *
 * @param {Array<Array<*>>} ledgerRows - values starting at firstRow; col 0 = Account, col 1 = Asset Class
 * @param {Array<{name: string, zestimate: *}>} fetched - resolved API results
 * @param {number} firstRow - the sheet row number corresponding to ledgerRows[0]
 * @returns {{updates: Array<{row: number, value: number}>, skipped: Array<string>}}
 */
function _planRealEstateUpdates(ledgerRows, fetched, firstRow) {
  const updates = [];
  const skipped = [];

  (fetched || []).forEach(f => {
    const name = f && f.name ? String(f.name).trim() : "";
    const value = Number(f && f.zestimate);

    if (!name) { skipped.push("(unnamed property): no account name"); return; }
    if (!(value > 0)) { skipped.push(`${name}: no usable Zestimate`); return; }

    let matched = 0;
    for (let i = 0; i < ledgerRows.length; i++) {
      if (String(ledgerRows[i][0]).trim() !== name) continue;
      matched++;
      if (String(ledgerRows[i][1]).trim() !== "Real Estate") {
        skipped.push(`${name}: row ${firstRow + i} is not a Real Estate row`);
        continue;
      }
      updates.push({ row: firstRow + i, value: value });
    }
    if (matched === 0) skipped.push(`${name}: no matching account on the ledger`);
  });

  return { updates: updates, skipped: skipped };
}

/**
 * Fetches Zestimates using config from the Settings tab.
 *
 * Writes ONLY the individual Real Estate cells that changed. It must never
 * round-trip a range through getValues()/setValues(): getValues() returns
 * computed results, so writing them back replaces every formula in the range
 * with a frozen literal.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet (for DI)
 * @returns {{updated: number, skipped: Array<string>}}
 */
function updateRealEstatePrices(ss_inject) {
  const ss = _resolveSpreadsheet(ss_inject);
  const interactive = _isInteractive();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet || !sheet) return { updated: 0, skipped: ["Required tabs not found"] };

  const apiKey = getSecret("rapidApiKey", ss);
  const apiHost = configSheet.getRange("B10").getValue();

  if (!apiKey) return { updated: 0, skipped: ["RapidAPI key not configured"] };

  const propData = configSheet.getRange("A29:B45").getValues();
  const properties = [];
  for (let i = 0; i < propData.length; i++) {
    if (propData[i][0] && propData[i][1]) {
      properties.push({ name: String(propData[i][0]), zpid: String(propData[i][1]) });
    }
  }

  const fetched = [];
  properties.forEach(prop => {
    try {
      const url = `https://${apiHost}/api/property-details/byzpid?zpid=${prop.zpid}`;
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: { 'X-RapidAPI-Key': apiKey, 'X-RapidAPI-Host': apiHost },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        fetched.push({ name: prop.name, zestimate: data && data.property ? data.property.zestimate : null });
      } else {
        Logger.log(`${prop.name}: HTTP ${response.getResponseCode()}`);
      }
    } catch (e) {
      Logger.log(`Fetch failed for ${prop.name}: ${e}`);
    }
  });

  const FIRST_DATA_ROW = 7;
  const lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW);
  const rowCount = lastRow - FIRST_DATA_ROW + 1;
  const ledgerRows = sheet.getRange(FIRST_DATA_ROW, 1, rowCount, 2).getValues();

  const plan = _planRealEstateUpdates(ledgerRows, fetched, FIRST_DATA_ROW);

  let written = 0;
  plan.updates.forEach(u => {
    const cell = sheet.getRange(u.row, 5);
    // Never overwrite a formula-driven Current Value (e.g. a Brokerage SUMPRODUCT row).
    if (cell.getFormula() !== "") {
      plan.skipped.push(`Row ${u.row}: cell holds a formula — left untouched`);
      return;
    }
    cell.setValue(u.value);
    written++;
  });

  plan.skipped.forEach(msg => Logger.log(`Zestimate skipped — ${msg}`));
  return { updated: written, skipped: plan.skipped };
}
function buildSettingsTab(ss_inject) {
  const ss = _resolveSpreadsheet(ss_inject);
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
  styleRow(sheet.getRange("A13:B13"), THEME.kpiCardBg).setValues([["GitHub PAT (gist scope)", SECRET_MOVED_NOTICE]]);
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
  const ss = _resolveSpreadsheet(ss_inject);
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
  // All formulas come from Formulas.gs so the builder and repairFormulas()
  // can never drift apart.
  const FIXED = _ledgerFixedFormulas();
  Object.keys(FIXED).forEach(a1 => sheet.getRange(a1).setFormula(FIXED[a1]));

  sheet.getRange("M1").setValue("⚙ INTERNAL — do not edit or delete")
    .setFontColor(THEME.mutedText).setFontSize(8);
  sheet.getRange("M2:M3").setNumberFormat("0.000000");
  sheet.getRange("N2:O3").setNumberFormat("#,##0.00");

  [1, 2].forEach(idx => {
    const sn = CARD_STYLES[idx]; const cn = CARD_LAYOUT[idx];
    sheet.getRange(cn.bg).setBackground(sn.bg);
    sheet.getRange(`${cn.lbl}2`).setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}2`).setNumberFormat("@").setFontColor(sn.valueFg).setFontSize(14).setFontWeight("bold");
    sheet.getRange(`${cn.lbl}3`).setFontColor(sn.labelFg).setFontWeight("bold").setFontSize(9);
    sheet.getRange(`${cn.val}3`).setNumberFormat("@").setFontColor(sn.subFg).setFontSize(11);
  });
  sheet.hideColumns(13, 3);   // M, N, O

  sheet.setRowHeight(2, 38); sheet.setRowHeight(3, 28);

  sheet.getRange("A4:C4").setBackground(THEME.quickStats.liquidBg);
  sheet.getRange("A4").setValue("🌊 Liquid Net Worth").setFontColor(THEME.quickStats.liquidFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("B4").setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.liquidFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("D4:G4").setBackground(THEME.quickStats.lockedBg);
  sheet.getRange("D4").setValue("🔒 Locked Net Worth").setFontColor(THEME.quickStats.lockedFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("E4").setNumberFormat(USD_ABBR_FMT).setFontColor(THEME.quickStats.lockedFg).setFontSize(11).setFontWeight("bold");

  sheet.getRange("H4:K4").setBackground(THEME.quickStats.fireBg);
  sheet.getRange("H4").setFontColor(THEME.quickStats.fireFg).setFontWeight("bold").setFontSize(9);
  sheet.getRange("I4").setNumberFormat("0.0%").setFontColor(THEME.quickStats.fireFg).setFontSize(11).setFontWeight("bold");

  sheet.setRowHeight(4, 28);
  sheet.getRange("A5:K5").setBackground(THEME.accentBar);
  sheet.setRowHeight(5, 3);

  const headers = LEDGER_HEADERS;
  sheet.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(6, 36);

  sheet.getRange(7, 1, DEFAULT_PORTFOLIO_DATA.length, headers.length).setValues(DEFAULT_PORTFOLIO_DATA);
  const NUM_ROWS = LEDGER_NUM_ROWS;
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
    const f = _ledgerRowFormulas(i + LEDGER_FIRST_ROW);
    exch.push([f.F]); gross.push([f.H]); net.push([f.I]);
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
  const ss = _resolveSpreadsheet(ss_inject);
  let sheet = ss.getSheetByName("Brokerage Holdings");
  if (!sheet) sheet = ss.insertSheet("Brokerage Holdings");
  else sheet.clear(); 

  sheet.setHiddenGridlines(true);
  
  const headers = HOLDINGS_HEADERS;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(THEME.headerBg).setFontColor(THEME.headerText).setFontWeight("bold");

  const sampleData = [
    ["Taxable Brokerage", "US Equity", "VTI", 150],
    ["Taxable Brokerage", "Individual Stocks", "AAPL", 50],
    ["401k / RRSP", "Technology", "QQQ", 200]
  ];
  sheet.getRange(2, 1, sampleData.length, 4).setValues(sampleData);

  const numRows = 99;
  const priceF = [], totalF = [];
  for (let i = 0; i < numRows; i++) {
    const f = _holdingsRowFormulas(i + HOLDINGS_FIRST_ROW);
    priceF.push([f.E]); totalF.push([f.F]);
  }
  sheet.getRange(2, 5, numRows, 1).setFormulas(priceF);
  sheet.getRange(2, 6, numRows, 1).setFormulas(totalF);

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
  const ss = _resolveSpreadsheet(ss_inject);
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
      .setOption('pieHole', 0.55)
      .setOption('colors', THEME.charts.donut)
      .setOption('pieSliceBorderColor', "transparent")
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '80%'})
      .setOption('legend', {position: 'right', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('pieSliceText', 'none')
      .setPosition(4, 2, 0, 0) // Row 4, Col B
      .build();

    uiSheet.insertChart(pieChart);
  }

  // B. Time-to-FIRE Forecast (Combo Chart: Actuals + CAGR Projection)
  const lastSnapRow = snapSheet.getLastRow();
  if (lastSnapRow > 1) {
    const snapData = snapSheet.getRange(2, 1, lastSnapRow - 1, 2).getValues();
    snapData.reverse(); // Snapshots are push-to-top, reverse for chronological order

    const forecastData = [["Date", "Actual Net Worth", "Projected Timeline"]];
    const cfgSheet = ss.getSheetByName("Settings & Config");
    const targetUSD = Number(cfgSheet && cfgSheet.getRange("B22").getValue())
      || DASHBOARD_CONFIG.fireTargetUSD || 3000000;

    if (snapData.length > 1) {
      const firstRow = snapData[0];
      const lastRow = snapData[snapData.length - 1];
      const firstDate = new Date(firstRow[0]);
      const lastDate = new Date(lastRow[0]);
      const firstNet = Number(firstRow[1]);
      const lastNet = Number(lastRow[1]);

      for (let i = 0; i < snapData.length; i++) {
        forecastData.push([snapData[i][0], snapData[i][1], null]);
      }

      const yearsElapsed = (lastDate - firstDate) / (1000 * 60 * 60 * 24 * 365.25);
      
      // Compute CAGR and project if there is actual positive growth and haven't hit target yet
      if (yearsElapsed >= 0 && firstNet > 0 && lastNet > firstNet && lastNet < targetUSD) {
        // Prevent Divide by Zero parsing if yearsElapsed is mathematically 0 (same day snapshot)
        const safeYears = Math.max(yearsElapsed, 0.25); 
        const cagr = Math.pow(lastNet / firstNet, 1 / safeYears) - 1;

        if (cagr > 0) {
          // Join the actual line to the projected line visually
          forecastData[forecastData.length - 1][2] = lastNet;

          let projectedNet = lastNet;
          let projectedDate = new Date(lastDate);
          let safetyStops = 0;

          // Steps by quarter up to 30 years (120 quarters)
          while (projectedNet < targetUSD && safetyStops < 120) {
            projectedDate.setMonth(projectedDate.getMonth() + 3);
            projectedNet = projectedNet * Math.pow(1 + cagr, 0.25);

            if (projectedNet > targetUSD) projectedNet = targetUSD;
            forecastData.push([new Date(projectedDate), null, projectedNet]);
            safetyStops++;
          }
        }
      }
    } else {
      forecastData.push([snapData[0][0], snapData[0][1], null]);
    }

    dataSheet.getRange(1, 1, forecastData.length, 3).setValues(forecastData);

    const comboChart = uiSheet.newChart()
      .asComboChart()
      .addRange(dataSheet.getRange(1, 1, forecastData.length, 1)) // X: Dates
      .addRange(dataSheet.getRange(1, 2, forecastData.length, 1)) // Y1: Actuals
      .addRange(dataSheet.getRange(1, 3, forecastData.length, 1)) // Y2: Forecast
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Time-to-FIRE Trajectory (USD)')
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('seriesType', 'area')
      .setOption('series', {
        0: {type: 'area', color: THEME.charts.area[0], pointSize: 6, lineWidth: 3, label: 'Actual Net Worth'},
        1: {type: 'line', color: THEME.charts.stacked[0], pointSize: 0, lineWidth: 3, lineDashStyle: [4, 4], label: 'Projected FIRE Timeline (CAGR)'}
      })
      .setOption('chartArea', {left: '15%', top: '15%', width: '80%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: '$#,###' })
      .setOption('hAxis', { textStyle: {color: THEME.charts.axisText}, format: 'MMM yyyy' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setPosition(4, 7, 0, 0) // Row 4, Col G
      .build();

    uiSheet.insertChart(comboChart);

    // C. Asset Class Evolution (Stacked Area Chart)
    const stackedArea = uiSheet.newChart()
      .asAreaChart()
      .addRange(snapSheet.getRange(1, 1, lastSnapRow, 1)) // X: Dates
      .addRange(snapSheet.getRange(1, 9, lastSnapRow, 1)) // Cash
      .addRange(snapSheet.getRange(1, 10, lastSnapRow, 1)) // Brokerage
      .addRange(snapSheet.getRange(1, 11, lastSnapRow, 1)) // Retirement
      .addRange(snapSheet.getRange(1, 8, lastSnapRow, 1))  // Real Estate
      .addRange(snapSheet.getRange(1, 12, lastSnapRow, 1)) // Liabilities
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Historical Asset Class Evolution')
      .setOption('isStacked', true)
      .setOption('colors', [THEME.assetRows['Cash'], THEME.assetRows['Brokerage'], THEME.assetRows['Retirement'], THEME.assetRows['Real Estate'], THEME.assetRows['Liability']]) 
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '10%', top: '15%', width: '85%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: '$#,###' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('series', {
        0: {label: 'Cash (USD)'},
        1: {label: 'Brokerage (USD)'},
        2: {label: 'Retirement (USD)'},
        3: {label: 'Real Estate (USD)'},
        4: {label: 'Liabilities (USD)'}
      })
      .setPosition(22, 2, 0, 0) // Row 22, Col B (Below the others)
      .build();

    uiSheet.insertChart(stackedArea);
  }

  // D. Portfolio X-Ray (Donut Chart)
  const holdingsSheet = ss.getSheetByName("Brokerage Holdings");
  if (holdingsSheet) {
    const lastXrayRow = holdingsSheet.getLastRow();
    if (lastXrayRow > 1) {
      const xrayChart = uiSheet.newChart()
        .asPieChart()
        .addRange(holdingsSheet.getRange(2, 2, lastXrayRow - 1, 1)) // Asset Category
        .addRange(holdingsSheet.getRange(2, 6, lastXrayRow - 1, 1)) // Total Value
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
        .setOption('title', 'Portfolio Exposure X-Ray')
        .setOption('pieHole', 0.55)
        .setOption('colors', THEME.charts.donut)
        .setOption('pieSliceBorderColor', "transparent")
        .setOption('backgroundColor', { fill: 'transparent' })
        .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '80%'})
        .setOption('legend', {position: 'right', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
        .setOption('pieSliceText', 'none')
        .setPosition(22, 7, 0, 0) // Row 22, Col G (Next to Liquid vs Locked)
        .build();

      uiSheet.insertChart(xrayChart);
    }
  }
}
/**
 * Execute a manual snapshot. Calculates deltas, generates insights,
 * then chains cloud backups with transparent status reporting.
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet
 * @param {boolean} [silent=false] - Suppress UI reports
 */
function captureSnapshot(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const logSheet = ss.getSheetByName("Snapshots");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!mainSheet || !logSheet) return;

  // Integrity gate: a snapshot taken over damaged formulas freezes a stale net
  // worth into permanent history. Refuse rather than record a bad reading.
  const audit = auditFormulaHealth(ss);
  if (!audit.healthy) {
    const report = _formatHealthReport(audit);
    Logger.log("Snapshot aborted — " + report);
    if (!silent) {
      const ui = SpreadsheetApp.getUi();
      const choice = ui.alert(
        "⚠️ Snapshot blocked",
        report + "\n\nRecording a snapshot now would freeze incorrect values into your history.\n\nSnapshot anyway?",
        ui.ButtonSet.YES_NO
      );
      if (choice !== ui.Button.YES) return { aborted: true, reason: report };
    } else {
      return { aborted: true, reason: report };
    }
  }

  const netUSD = mainSheet.getRange("B2").getValue();
  const grossUSD = mainSheet.getRange("B3").getValue();
  // BUGFIX: E2/H2 are the *display* cells (H2 is actually a label, not a value).
  // Read the hidden numeric backing cells so Snapshots always logs real numbers.
  const netCAD = Number(mainSheet.getRange("N2").getValue()) || 0;
  const netINR = Number(mainSheet.getRange("O2").getValue()) || 0;
  const dataRange = mainSheet.getRange("A7:J80").getValues(); 
  
  let liquidUSD = 0, lockedUSD = 0, totalReUSD = 0;
  let cashUSD = 0, brokerageUSD = 0, retirementUSD = 0, liabilityUSD = 0;

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
      
      if (assetClass === "Cash") cashUSD += netVal;
      else if (assetClass === "Brokerage") brokerageUSD += netVal;
      else if (assetClass === "Retirement") retirementUSD += netVal;
      else if (assetClass === "Real Estate") totalReUSD += netVal;
      else if (assetClass === "Liability") liabilityUSD += netVal;
    }
  }

  const prevNetUSD = logSheet.getRange(2, 2).getValue(); 
  const prevLiquid = logSheet.getRange(2, 3).getValue();
  const fireTarget = Number(configSheet && configSheet.getRange("B22").getValue())
    || DASHBOARD_CONFIG.fireTargetUSD || 3000000;
  let dollarDelta = "", pctGrowth = "", autoInsight = "Initial baseline snapshot established.", fireProgress = netUSD / fireTarget;

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
  const rowData = [
    new Date(), netUSD, liquidUSD, lockedUSD, grossUSD, 
    netCAD, netINR, totalReUSD, 
    cashUSD, brokerageUSD, retirementUSD, liabilityUSD,
    dollarDelta, pctGrowth, fireProgress, autoInsight, ""
  ];
  logSheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);

  logSheet.getRange(2, 2, 1, 11).setNumberFormat("$#,##0.00"); 
  logSheet.getRange(2, 13).setNumberFormat("[Color10]+$#,##0.00;[Color3]-$#,##0.00"); 
  logSheet.getRange(2, 14).setNumberFormat("[Color10]+0.00%;[Color3]-0.00%"); 
  logSheet.getRange(2, 11).setNumberFormat("0.00%"); 
  
  // --- Cloud Backup Chain with Transparent Status ---
  const gistConfigured = _isGistConfigured(configSheet);
  const gistOk    = gistConfigured ? backupToGitHub(ss, true) : false;
  const driveResult = backupToGoogleDrive(ss, true);
  const driveOk   = driveResult && driveResult.success;

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

  if (!silent) {
    SpreadsheetApp.getUi().alert(statusParts.join("\n"));
  }

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
    netWorthSecondary1:  Number(mainSheet.getRange("N2").getValue()) || 0,
    grossWorthSecondary1: Number(mainSheet.getRange("N3").getValue()) || 0,
    netWorthSecondary2:  Number(mainSheet.getRange("O2").getValue()) || 0,
    grossWorthSecondary2: Number(mainSheet.getRange("O3").getValue()) || 0,
    liquidNetWorthUSD: mainSheet.getRange("B4").getValue() || 0,
    lockedNetWorthUSD: mainSheet.getRange("E4").getValue() || 0,
    fireProgress:      mainSheet.getRange("I4").getValue() || 0,
  };

  // Latest snapshot row (if exists)
  let latestSnapshot = null;
  if (snapSheet && snapSheet.getLastRow() > 1) {
    const snapRow = snapSheet.getRange(2, 1, 1, 17).getValues()[0];
    latestSnapshot = {
      date:         snapRow[0],
      netUSD:       snapRow[1],
      liquidUSD:    snapRow[2],
      lockedUSD:    snapRow[3],
      grossUSD:     snapRow[4],
      valueDelta:   snapRow[12],
      pctGrowth:    snapRow[13],
      fireProgress: snapRow[14],
      autoInsight:  snapRow[15],
      manualNotes:  snapRow[16]
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
function forceBackup(ss_inject) {
  const ss = _resolveSpreadsheet(ss_inject);
  backupToGitHub(ss, false);
  backupToGoogleDrive(ss, false);
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
function _isGistConfigured(configSheet, ss_inject) {
  if (!configSheet) return false;
  const pat = getSecret("githubPat", ss_inject);
  const gistId = configSheet.getRange("B14").getValue();
  return !!pat && !!gistId && gistId !== "PASTE_GIST_ID_HERE";
}

/**
 * Disaster Recovery: Serializes live ledger into enriched JSON and pushes to a private GitHub Gist.
 * Silently skips if not configured (no errors thrown).
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet (for injection)
 * @param {boolean} [silent=false] - If true, suppresses UI alerts on success.
 * @returns {boolean} Whether the backup was attempted and succeeded.
 */
function backupToGitHub(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!_isGistConfigured(configSheet, ss)) {
    // Silently skip — not configured
    return false;
  }

  const githubToken = getSecret("githubPat", ss);
  const gistId = configSheet.getRange("B14").getValue();
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
 * Uses DriveApp — zero additional setup required beyond standard GAS authorization.
 * Creates a "WealthScript — Backups" folder automatically on first run.
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet (for DI/testing)
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 * @returns {{success: boolean, folder: GoogleAppsScript.Drive.Folder}} Result object
 */
function backupToGoogleDrive(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const FOLDER_NAME = "WealthScript \u2014 Backups";
  const ss = _resolveSpreadsheet(ss_inject);

  try {
    const backupData = _buildEnrichedBackup(ss);
    const jsonContent = JSON.stringify(backupData, null, 2);

    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext()
      ? folderIterator.next()
      : DriveApp.createFolder(FOLDER_NAME);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
    const fileName = `net_worth_${timestamp}.json`;
    folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `\u2705 Google Drive Backup Successful!\n\nFolder: "${FOLDER_NAME}"\nFile: ${fileName}`
      );
    }
    return { success: true, folder };
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
    return { success: false, folder: null };
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
/**
 * ==========================================
 * REPAIR, MIGRATION & FORMULA INTEGRITY
 * ==========================================
 * Non-destructive operations safe to run against a live sheet at any time.
 * Nothing in this file calls sheet.clear().
 */

const INTEGRITY_PROP_KEY = "wealthscript.formulaBaseline";

/* ------------------------------------------------------------------ *
 * FORMULA INTEGRITY
 * ------------------------------------------------------------------ */

/**
 * Pure helper: compares current formula counts against a stored baseline.
 * A DROP is the signal — growth is fine (the user added accounts).
 *
 * @param {Object<string,number>} current - label -> formula count
 * @param {Object<string,number>} baseline - label -> formula count
 * @returns {{ok: boolean, regressions: Array<{label: string, was: number, now: number}>}}
 */
function _diffFormulaCounts(current, baseline) {
  const regressions = [];
  Object.keys(baseline || {}).forEach(label => {
    const was = Number(baseline[label]) || 0;
    const now = Number(current && current[label]) || 0;
    if (now < was) regressions.push({ label: label, was: was, now: now });
  });
  return { ok: regressions.length === 0, regressions: regressions };
}

/**
 * Counts non-empty formulas in each managed range.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Object<string,number>}
 */
function _countManagedFormulas(ss) {
  const counts = {};
  _managedRanges().forEach(spec => {
    const sheet = ss.getSheetByName(spec.sheet);
    if (!sheet) { counts[spec.label] = 0; return; }
    let n = 0;
    sheet.getRange(spec.a1).getFormulas().forEach(row => {
      row.forEach(f => { if (f) n++; });
    });
    counts[spec.label] = n;
  });
  return counts;
}

/**
 * Pure helper: classifies Brokerage rows against the Holdings tab.
 *
 * A missing formula is only damage when there is something to link TO. Tracking
 * a brokerage as a single manually-typed number is a legitimate pattern, and
 * injecting a SUMPRODUCT there would evaluate to 0 and silently erase the
 * balance. So the rule is: broken only if holdings exist under that exact name.
 *
 * A hardcoded ZERO is never intentional tracking. It is the fingerprint of a
 * SUMPRODUCT that was already returning 0 (usually an account-name mismatch)
 * at the moment something froze the column into literals. Those rows are
 * reported as SUSPECT: repair cannot fix them, because there is nothing to link
 * to until a human reconciles the names.
 *
 * @param {Array<Array<*>>} ledgerRows - A..J values from LEDGER_FIRST_ROW
 * @param {Object<string,number>} holdingAccounts - account name -> holdings row count
 * @param {Array<string>} formulaCol - column E formulas, parallel to ledgerRows
 * @param {Array<string>} [linkedClasses] - asset classes sourced from Holdings
 * @returns {{broken: Array, suspect: Array, unlinked: Array, orphaned: Array}}
 */
function _classifyBrokerageRows(ledgerRows, holdingAccounts, formulaCol, linkedClasses) {
  const classes = linkedClasses || DASHBOARD_CONFIG.holdingsLinkedClasses || ["Brokerage"];
  const broken = [], suspect = [], unlinked = [], dormant = [];
  const seen = {};

  for (let i = 0; i < ledgerRows.length; i++) {
    const assetClass = String(ledgerRows[i][1]).trim();
    if (classes.indexOf(assetClass) === -1) continue;

    const account = String(ledgerRows[i][0] || "").trim();
    const status = String(ledgerRows[i][9] || "").trim();
    const value = ledgerRows[i][4];
    const row = LEDGER_FIRST_ROW + i;
    seen[account] = true;

    if (formulaCol[i]) continue;                       // linked — nothing to do

    const base = { row: row, account: account, status: status, assetClass: assetClass, value: value };

    if (holdingAccounts[account]) {
      broken.push(Object.assign({ holdings: holdingAccounts[account] }, base));
    } else if (value !== "" && value !== null && Number(value) !== 0) {
      unlinked.push(base);                       // a real manually-kept balance
    } else if (status === "Active") {
      suspect.push(base);                        // $0 on a live account: damage
    } else {
      // Inactive rows are excluded from every SUMIFS(...,"Active"), so a zero
      // there is harmless by definition. Noting it is fine; blocking snapshots
      // on it would train the user to click through the warning.
      dormant.push(base);
    }
  }

  // Holdings whose account name matches no Brokerage row: real money that the
  // dashboard is not counting anywhere.
  const orphaned = Object.keys(holdingAccounts)
    .filter(n => !seen[n])
    .map(n => ({ account: n, rows: holdingAccounts[n] }));

  return { broken: broken, suspect: suspect, unlinked: unlinked, dormant: dormant, orphaned: orphaned };
}

/**
 * Reads the sheet and classifies every Brokerage row.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {{broken: Array, unlinked: Array, orphaned: Array}}
 */
function _auditBrokerageLinks(ss) {
  const ledger = ss.getSheetByName("Dashboard & Ledger");
  const holdings = ss.getSheetByName("Brokerage Holdings");
  if (!ledger) return { broken: [], suspect: [], unlinked: [], dormant: [], orphaned: [] };

  const holdingAccounts = {};
  if (holdings) {
    holdings.getRange(HOLDINGS_FIRST_ROW, 1, HOLDINGS_NUM_ROWS, 1).getValues()
      .forEach(r => {
        const n = String(r[0] || "").trim();
        if (n) holdingAccounts[n] = (holdingAccounts[n] || 0) + 1;
      });
  }

  const rows = ledger.getRange(LEDGER_FIRST_ROW, 1, LEDGER_NUM_ROWS, 10).getValues();
  const formulaCol = ledger.getRange(LEDGER_FIRST_ROW, 5, LEDGER_NUM_ROWS, 1)
    .getFormulas().map(r => r[0]);

  return _classifyBrokerageRows(rows, holdingAccounts, formulaCol);
}

/**
 * Full health check. Called on demand from the menu and defensively by
 * captureSnapshot() before it writes a row.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @returns {{healthy: boolean, regressions: Array, brokenLinks: Array, counts: Object, hadBaseline: boolean}}
 */
function auditFormulaHealth(ss_inject) {
  const ss = _resolveSpreadsheet(ss_inject);
  const counts = _countManagedFormulas(ss);
  const links = _auditBrokerageLinks(ss);

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(INTEGRITY_PROP_KEY);
  let baseline = null;
  try { baseline = raw ? JSON.parse(raw) : null; } catch (e) { baseline = null; }

  const diff = _diffFormulaCounts(counts, baseline);

  return {
    // Only genuine damage blocks a snapshot. Unlinked and orphaned accounts are
    // reported for review but may well be intentional.
    healthy: diff.ok && links.broken.length === 0 && links.suspect.length === 0
      && _auditHeaders(ss).length === 0,
    regressions: diff.regressions,
    brokenLinks: links.broken,
    suspect: links.suspect,
    unlinked: links.unlinked,
    dormant: links.dormant,
    headerProblems: _auditHeaders(ss),
    orphaned: links.orphaned,
    counts: counts,
    hadBaseline: !!baseline
  };
}

/** Records the current formula counts as the accepted baseline. */
function _saveFormulaBaseline(ss) {
  PropertiesService.getDocumentProperties()
    .setProperty(INTEGRITY_PROP_KEY, JSON.stringify(_countManagedFormulas(ss)));
}

/** Pure helper: renders an audit result as a human-readable report. */
function _formatHealthReport(audit) {
  const notices = [];

  (audit.unlinked || []).forEach(u => notices.push(
    `  • Row ${u.row} — ${u.account}${u.status ? ` (${u.status})` : ""}: manual value, no holdings tracked. Fine if intentional.`));
  (audit.dormant || []).forEach(d => notices.push(
    `  • Row ${d.row} — ${d.account} [${d.assetClass}] is ${d.status || "not Active"} and holds 0. Excluded from net worth, so this is harmless.`));
  (audit.orphaned || []).forEach(o => notices.push(
    `  • "${o.account}" has ${o.rows} holdings row(s) but NO Brokerage row on the ledger — this money is not counted in your net worth.`));

  const lines = [];

  if ((audit.headerProblems || []).length) {
    lines.push("🛑 Header row does not match the expected layout.", "");
    audit.headerProblems.forEach(h => lines.push(`  • ${h}`));
    lines.push("", "A column has probably been inserted or deleted. Repair writes by",
                   "fixed cell reference, so it is BLOCKED until the headers match —",
                   "otherwise it would write every formula into the wrong column.", "");
    return lines.join("\n");
  }

  if (audit.healthy) {
    lines.push(audit.hadBaseline
      ? "✅ All managed formulas are intact."
      : "✅ No damage detected. Baseline recorded — future checks compare against it.");
  } else {
    lines.push("⚠️ Formula damage detected.", "");
    if (audit.regressions.length) {
      lines.push("Formula count dropped in:");
      audit.regressions.forEach(r => lines.push(`  • ${r.label} — was ${r.was}, now ${r.now}`));
      lines.push("");
    }
    if (audit.brokenLinks.length) {
      lines.push("Rows with holdings but no link (Repair Formulas fixes these):");
      audit.brokenLinks.forEach(b => lines.push(
        `  • Row ${b.row} — ${b.account} [${b.assetClass}] (${b.holdings} holdings row(s) waiting)`));
      lines.push("");
    }
    if ((audit.suspect || []).length) {
      lines.push("Rows frozen at 0 with no matching holdings — NEEDS YOUR ATTENTION:");
      audit.suspect.forEach(x => lines.push(
        `  • Row ${x.row} — ${x.account} [${x.assetClass}]. Counting as $0 in your net worth.`));
      lines.push("    Repair cannot fix these: the account name likely doesn't match the");
      lines.push("    Holdings tab. Reconcile the names, then run Repair Formulas.");
      lines.push("");
    }
    if (audit.brokenLinks.length) lines.push("Run WealthScript > 🛠 Repair Formulas to restore the links above.");
  }

  if (notices.length) {
    lines.push("", "ℹ️ For review (not errors, nothing will be changed):", ...notices);
  }

  return lines.join("\n");
}

/** Menu action: report formula health and record a baseline if none exists. */
function checkFormulaHealth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const audit = auditFormulaHealth(ss);
  if (!audit.hadBaseline) _saveFormulaBaseline(ss);
  SpreadsheetApp.getUi().alert(_formatHealthReport(audit));
}

/**
 * Checks both header rows against canonical. A mismatch means a column was
 * inserted or removed, and every hardcoded A1 write in this file would land in
 * the wrong cell.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Array<string>}
 */
function _auditHeaders(ss) {
  const out = [];
  const pairs = [
    { name: "Dashboard & Ledger", row: 6, expected: LEDGER_HEADERS },
    { name: "Brokerage Holdings", row: 1, expected: HOLDINGS_HEADERS }
  ];
  pairs.forEach(p => {
    const sheet = ss.getSheetByName(p.name);
    if (!sheet) return;
    const actual = sheet.getRange(p.row, 1, 1, p.expected.length).getValues()[0];
    _verifyHeaders(actual, p.expected).problems
      .forEach(msg => out.push(`${p.name} row ${p.row}: column ${msg}`));
  });
  return out;
}

/* ------------------------------------------------------------------ *
 * REPAIR
 * ------------------------------------------------------------------ */

/**
 * Re-injects every managed formula into a LIVE sheet without clearing anything.
 *
 * Safety rule: a cell holding a hardcoded literal in a manual-entry range is
 * never overwritten. That protects deliberately pinned values — option contract
 * prices on the Holdings tab, manually typed balances on the ledger — which a
 * naive fill-down would destroy.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{restored: number, preserved: number, details: Array<string>}}
 */
function repairFormulas(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const ledger = ss.getSheetByName("Dashboard & Ledger");
  const holdings = ss.getSheetByName("Brokerage Holdings");

  const details = [];
  let restored = 0, preserved = 0;

  if (!ledger || !holdings) {
    return { restored: 0, preserved: 0, details: ["Required tabs not found — run First Time Setup."] };
  }

  // Refuse rather than write into shifted columns.
  const headerProblems = _auditHeaders(ss);
  if (headerProblems.length) {
    const msg = ["🛑 Repair aborted — the header row doesn't match the expected layout.", ""]
      .concat(headerProblems.map(h => "  • " + h))
      .concat(["", "Repair writes by fixed cell reference. Fix the headers first,",
               "then run Repair Formulas again."]).join("\n");
    if (!silent) SpreadsheetApp.getUi().alert(msg);
    return { restored: 0, preserved: 0, details: headerProblems };
  }

  // --- 1. Fixed KPI cells. These are always formulas; a literal here is damage.
  const fixed = _ledgerFixedFormulas();
  Object.keys(fixed).forEach(a1 => {
    const cell = ledger.getRange(a1);
    if (cell.getFormula() !== fixed[a1]) {
      cell.setFormula(fixed[a1]);
      restored++;
      details.push(`Ledger ${a1} restored`);
    }
  });

  // --- 2. Ledger per-row columns F, H, I. Always formulas.
  for (let r = LEDGER_FIRST_ROW; r <= LEDGER_LAST_ROW; r++) {
    const f = _ledgerRowFormulas(r);
    ["F", "H", "I"].forEach(col => {
      const cell = ledger.getRange(`${col}${r}`);
      if (cell.getFormula() !== f[col]) { cell.setFormula(f[col]); restored++; }
    });
  }

  // --- 3. Ledger column E, Brokerage rows.
  //        CRITICAL: only link rows whose account name actually appears on the
  //        Holdings tab. Injecting a SUMPRODUCT that matches nothing evaluates
  //        to 0 and would silently erase a manually maintained balance.
  const links = _auditBrokerageLinks(ss);
  links.broken.forEach(b => {
    const want = _buildBrokerageFormula(b.row);
    const cell = ledger.getRange(b.row, 5);
    if (cell.getFormula() === want) return;
    if (cell.getFormula() === "" && cell.getValue() !== "") {
      details.push(`Row ${b.row} (${b.account}): replaced a frozen literal with the live Holdings link`);
    }
    cell.setFormula(want);
    restored++;
  });
  links.unlinked.forEach(u => {
    preserved++;
    details.push(`Row ${u.row} (${u.account}): kept manual value ${u.value} — no holdings under this name`);
  });
  links.dormant.forEach(d => {
    preserved++;
    details.push(`Row ${d.row} (${d.account}): ${d.status || "not Active"} and $0 — left alone, not counted in net worth`);
  });
  links.suspect.forEach(x => {
    details.push(`⚠️ Row ${x.row} (${x.account}) frozen at 0 with no matching holdings — counting as $0. Fix the account name, then re-run Repair.`);
  });
  links.orphaned.forEach(o => {
    details.push(`⚠️ "${o.account}" has ${o.rows} holdings row(s) but no ledger row — not counted in net worth`);
  });

  // --- 4. Holdings E (price) and F (total). E is skipped where the user has
  //        pinned a literal — options and unquotable tickers have no live price.
  for (let r = HOLDINGS_FIRST_ROW; r <= HOLDINGS_LAST_ROW; r++) {
    const f = _holdingsRowFormulas(r);

    // Column E is user-overridable: FX lookups on foreign-currency cash rows,
    // fallback prices for unquotable tickers, pinned option marks. Only fill
    // blanks or upgrade a known previous canonical formula. NEVER replace
    // something the user wrote — that is how an override becomes a wrong number
    // that still looks like a working formula.
    const priceCell = holdings.getRange(r, 5);
    const cur = priceCell.getFormula();
    const isBlank = cur === "" && priceCell.getValue() === "";

    if (cur === f.E) {
      // already canonical
    } else if (isBlank || _legacyHoldingsPriceFormulas(r).indexOf(cur) > -1) {
      priceCell.setFormula(f.E);
      restored++;
    } else {
      preserved++;
      details.push(cur
        ? `Holdings E${r}: kept custom formula ${cur}`
        : `Holdings E${r}: kept pinned price ${priceCell.getValue()}`);
    }

    const totalCell = holdings.getRange(r, 6);
    if (totalCell.getFormula() !== f.F) { totalCell.setFormula(f.F); restored++; }
  }

  _saveFormulaBaseline(ss);

  if (!silent) {
    const msg = [`✅ Repair complete.`, ``,
                 `Formulas restored: ${restored}`,
                 `Pinned values preserved: ${preserved}`, ``]
      .concat(details.length ? ["Notes:"].concat(details.slice(0, 25).map(d => "  • " + d)) : [])
      .join("\n");
    SpreadsheetApp.getUi().alert(msg);
  }

  return { restored: restored, preserved: preserved, details: details };
}

/* ------------------------------------------------------------------ *
 * MIGRATION
 * ------------------------------------------------------------------ */

/**
 * Brings a pre-existing sheet up to the current layout. Idempotent — running
 * it twice is a no-op.
 *
 * Migration 1: insert Settings!B22 (FIRE Target), shifting currencies to B24/B25.
 * Migration 2: create hidden FX helper cells M:O on the Dashboard.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{applied: Array<string>, skipped: Array<string>}}
 */
function migrateSheetLayout(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const cfg = ss.getSheetByName("Settings & Config");
  const ledger = ss.getSheetByName("Dashboard & Ledger");

  const applied = [], skipped = [];

  if (!cfg || !ledger) {
    return { applied: [], skipped: ["Required tabs not found — run First Time Setup."] };
  }

  const headerProblems = _auditHeaders(ss);
  if (headerProblems.length) {
    const msg = ["🛑 Migration aborted — the header row doesn't match the expected layout.", ""]
      .concat(headerProblems.map(h => "  • " + h)).join("\n");
    if (!silent) SpreadsheetApp.getUi().alert(msg);
    return { applied: [], skipped: headerProblems };
  }

  // --- Migration 1: FIRE target row ---
  if (String(cfg.getRange("A22").getValue()).indexOf("FIRE Target") === 0) {
    skipped.push("Settings!B22 FIRE target already present");
  } else {
    cfg.insertRowBefore(22);
    cfg.getRange("A22:B22")
      .setValues([["FIRE Target Net Worth (USD)", (DASHBOARD_CONFIG.fireTargetUSD || 3000000)]])
      .setBackground(THEME.kpiCardBg).setVerticalAlignment("middle");
    cfg.getRange("B22").setNumberFormat("$#,##0")
      .setNote("The FIRE Progress card and the Time-to-FIRE chart read this cell live. Change it any time — no rebuild needed.");
    applied.push("Inserted Settings!B22 (FIRE Target); currencies now B24/B25, ZPID map now A29:B45");
  }

  // --- Migration 2: hidden FX helper cells ---
  if (ledger.getRange("M2").getFormula() && ledger.getRange("O3").getFormula()) {
    skipped.push("Hidden helper cells M:O already present");
  } else {
    ledger.getRange("M1").setValue("⚙ INTERNAL — do not edit or delete")
      .setFontColor(THEME.mutedText).setFontSize(8);
    ledger.getRange("M2:M3").setNumberFormat("0.000000");
    ledger.getRange("N2:O3").setNumberFormat("#,##0.00");
    applied.push("Created hidden FX helper cells M1:O3");
  }

  // Formulas + column hiding are idempotent, so always reassert them.
  const fixed = _ledgerFixedFormulas();
  Object.keys(fixed).forEach(a1 => ledger.getRange(a1).setFormula(fixed[a1]));
  ["E2", "E3", "I2", "I3"].forEach(a1 => ledger.getRange(a1).setNumberFormat("@"));
  ledger.hideColumns(13, 3);   // M, N, O
  applied.push("Reasserted KPI card formulas and text formatting");

  _saveFormulaBaseline(ss);

  if (!silent) {
    const msg = ["🧭 Migration complete.", ""]
      .concat(applied.length ? ["Applied:"].concat(applied.map(a => "  • " + a)) : [])
      .concat(skipped.length ? ["", "Already up to date:"].concat(skipped.map(s => "  • " + s)) : [])
      .join("\n");
    SpreadsheetApp.getUi().alert(msg);
  }

  return { applied: applied, skipped: skipped };
}

/* ------------------------------------------------------------------ *
 * DESTRUCTIVE OPERATION GUARDS
 * ------------------------------------------------------------------ */

/**
 * Pure helper: does this ledger hold real user data?
 * Sample rows shipped by First Time Setup don't count as real data.
 *
 * @param {Array<Array<*>>} ledgerRows - values from A7 onward
 * @returns {boolean}
 */
function _hasUserLedgerData(ledgerRows) {
  const sampleNames = DEFAULT_PORTFOLIO_DATA.map(r => String(r[0]));
  for (let i = 0; i < (ledgerRows || []).length; i++) {
    const name = String(ledgerRows[i][0] || "").trim();
    if (!name) continue;
    if (sampleNames.indexOf(name) === -1) return true;
  }
  return false;
}

/**
 * Requires the user to type a confirmation phrase before a destructive rebuild.
 * @param {string} phrase
 * @param {string} warning
 * @returns {boolean} whether the user confirmed
 */
function _confirmDestructive(phrase, warning) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "⚠️ Destructive operation",
    `${warning}\n\nThis CANNOT be undone from the menu. Recover via File > Version history if needed.\n\nType ${phrase} to proceed:`,
    ui.ButtonSet.OK_CANCEL
  );
  return res.getSelectedButton() === ui.Button.OK &&
         String(res.getResponseText()).trim().toUpperCase() === phrase;
}

/** Menu action: full rebuild, behind a typed confirmation. */
function rebuildEverythingDestructive() {
  if (!_confirmDestructive("ERASE",
      "This ERASES every tab and replaces your ledger with the sample portfolio.")) {
    SpreadsheetApp.getUi().alert("Cancelled — nothing was changed.");
    return;
  }
  runFirstTimeSetup(null, false, true);
}

/** Menu action: rebuild only the Cash Flow tab, behind a typed confirmation. */
function rebuildCashFlowDestructive() {
  if (!_confirmDestructive("ERASE",
      "This ERASES your entire expense ledger on the Cash Flow & Burn tab.")) {
    SpreadsheetApp.getUi().alert("Cancelled — nothing was changed.");
    return;
  }
  buildCashFlowTab();
}


/* ------------------------------------------------------------------ *
 * SECRET STORAGE
 * ------------------------------------------------------------------ */

/** @returns {Object|null} the spec for a named secret */
function _secretSpec(name) {
  for (let i = 0; i < SECRET_SPECS.length; i++) {
    if (SECRET_SPECS[i].name === name) return SECRET_SPECS[i];
  }
  return null;
}

/**
 * Pure helper: is this cell value a real secret, or a placeholder / notice?
 * @param {*} value
 * @param {string} placeholder
 * @returns {boolean}
 */
function _isRealSecret(value, placeholder) {
  const v = String(value === undefined || value === null ? "" : value).trim();
  if (!v) return false;
  if (v === placeholder) return false;
  if (v === SECRET_MOVED_NOTICE) return false;
  return true;
}

/**
 * Reads a secret, preferring secure storage and falling back to the legacy
 * cell so a workbook that hasn't migrated yet keeps working.
 *
 * @param {string} name - key from SECRET_SPECS
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @returns {string} the secret, or "" when unset
 */
function getSecret(name, ss_inject) {
  const spec = _secretSpec(name);
  if (!spec) return "";

  const stored = PropertiesService.getDocumentProperties().getProperty(spec.prop);
  if (stored) return stored;

  const ss = _resolveSpreadsheet(ss_inject);
  const cfg = ss.getSheetByName("Settings & Config");
  if (!cfg) return "";

  const cellValue = cfg.getRange(spec.cell).getValue();
  return _isRealSecret(cellValue, spec.placeholder) ? String(cellValue).trim() : "";
}

/**
 * Writes a secret to secure storage and leaves a visible notice in its cell.
 * @param {string} name
 * @param {string} value
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 */
function setSecret(name, value, ss_inject) {
  const spec = _secretSpec(name);
  if (!spec) return;

  PropertiesService.getDocumentProperties().setProperty(spec.prop, String(value).trim());

  const ss = _resolveSpreadsheet(ss_inject);
  const cfg = ss.getSheetByName("Settings & Config");
  if (cfg) cfg.getRange(spec.cell).setValue(SECRET_MOVED_NOTICE);
}

/**
 * Pure helper: decides what a migration run would do.
 *
 * @param {Array<Object>} specs - SECRET_SPECS
 * @param {Object<string,*>} cellValues - name -> current cell value
 * @param {Object<string,*>} propValues - name -> current stored value
 * @returns {{move: Array, alreadySecure: Array, notSet: Array, conflict: Array}}
 */
function _planSecretMigration(specs, cellValues, propValues) {
  const move = [], alreadySecure = [], notSet = [], conflict = [];

  (specs || []).forEach(spec => {
    const cell = (cellValues || {})[spec.name];
    const prop = (propValues || {})[spec.name];
    const cellIsReal = _isRealSecret(cell, spec.placeholder);

    if (prop && cellIsReal && String(prop).trim() !== String(cell).trim()) {
      // Both set and different — never silently pick one.
      conflict.push(spec);
    } else if (prop) {
      alreadySecure.push(spec);
    } else if (cellIsReal) {
      move.push(spec);
    } else {
      notSet.push(spec);
    }
  });

  return { move: move, alreadySecure: alreadySecure, notSet: notSet, conflict: conflict };
}

/**
 * Menu action: moves credentials out of Settings cells into secure storage.
 * Idempotent — running it twice is a no-op.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{moved: Array<string>, alreadySecure: Array<string>, notSet: Array<string>, conflict: Array<string>}}
 */
function migrateSecretsToProperties(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const cfg = ss.getSheetByName("Settings & Config");
  const props = PropertiesService.getDocumentProperties();

  if (!cfg) {
    if (!silent) SpreadsheetApp.getUi().alert("Settings & Config tab not found.");
    return { moved: [], alreadySecure: [], notSet: [], conflict: [] };
  }

  const cellValues = {}, propValues = {};
  SECRET_SPECS.forEach(spec => {
    cellValues[spec.name] = cfg.getRange(spec.cell).getValue();
    propValues[spec.name] = props.getProperty(spec.prop);
  });

  const plan = _planSecretMigration(SECRET_SPECS, cellValues, propValues);

  plan.move.forEach(spec => {
    props.setProperty(spec.prop, String(cellValues[spec.name]).trim());
    cfg.getRange(spec.cell).setValue(SECRET_MOVED_NOTICE);
  });

  // A value already in secure storage means the cell copy is redundant.
  plan.alreadySecure.forEach(spec => {
    if (_isRealSecret(cellValues[spec.name], spec.placeholder)) {
      cfg.getRange(spec.cell).setValue(SECRET_MOVED_NOTICE);
    }
  });

  if (!silent) {
    const lines = ["🔒 Credential storage", ""];
    if (plan.move.length) {
      lines.push("Moved out of the sheet:");
      plan.move.forEach(s => lines.push(`  • ${s.label} (was ${s.cell})`));
      lines.push("");
    }
    if (plan.alreadySecure.length) {
      lines.push("Already secure:");
      plan.alreadySecure.forEach(s => lines.push(`  • ${s.label}`));
      lines.push("");
    }
    if (plan.notSet.length) {
      lines.push("Not configured (nothing to move):");
      plan.notSet.forEach(s => lines.push(`  • ${s.label}`));
      lines.push("");
    }
    if (plan.conflict.length) {
      lines.push("⚠️ Different values in the cell AND in secure storage —");
      lines.push("   left untouched so nothing is guessed at:");
      plan.conflict.forEach(s => lines.push(`  • ${s.label} (${s.cell})`));
      lines.push("   Clear whichever is stale, then re-run.");
      lines.push("");
    }
    if (plan.move.length) {
      lines.push("These values were visible to anyone with access to this");
      lines.push("workbook, including view-only collaborators. If the sheet has");
      lines.push("ever been shared, rotate them at the provider.");
    }
    SpreadsheetApp.getUi().alert(lines.join("\n"));
  }

  const names = arr => arr.map(s => s.label);
  return { moved: names(plan.move), alreadySecure: names(plan.alreadySecure),
           notSet: names(plan.notSet), conflict: names(plan.conflict) };
}
/**
 * ==========================================
 * BROKER SYNC — pluggable position feeds
 * ==========================================
 * Optional. Nothing here runs unless a user configures a provider, and the menu
 * item stays hidden until they do. WealthScript works exactly as before for
 * anyone who tracks holdings by hand.
 *
 * Provider 1: IBKR Flex Web Service.
 *
 * Design contract, inherited from repairFormulas(): sync UPDATES what it owns
 * and never deletes. Tickers that disappear are zeroed and flagged, not
 * removed. Rows belonging to other accounts are never touched. A user override
 * in a cell sync doesn't own survives untouched.
 */

const FLEX_BASE = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService";
const FLEX_VERSION = "3";
const FLEX_PROP_TOKEN = "wealthscript.flex.token";
const FLEX_PROP_QUERY = "wealthscript.flex.queryId";
const FLEX_PROP_ACCOUNT = "wealthscript.flex.accountName";

/** A sync that would shrink an account's block by more than this needs confirming. */
const SYNC_MAX_SHRINK = 0.30;

/* ------------------------------------------------------------------ *
 * PURE HELPERS
 * ------------------------------------------------------------------ */

/**
 * Builds the OCC-style option symbol used as the sheet's ticker for options.
 * GOOGLEFINANCE cannot price these, which is exactly why sync writes a literal
 * mark instead of a formula for option rows.
 *
 * @param {string} root - underlying symbol, e.g. "MSFT"
 * @param {string} expiry - yyyymmdd from Flex
 * @param {string} putCall - "C" or "P"
 * @param {number} strike
 * @returns {string} e.g. "MSFT260918C00545000"
 */
function _buildOccSymbol(root, expiry, putCall, strike) {
  const yymmdd = String(expiry || "").slice(2, 8);
  const cp = String(putCall || "").toUpperCase().charAt(0);
  const strikeInt = Math.round(Number(strike) * 1000);
  let padded = String(strikeInt);
  while (padded.length < 8) padded = "0" + padded;
  return `${String(root || "").toUpperCase()}${yymmdd}${cp}${padded}`;
}

/**
 * Normalises one Flex OpenPosition into a sheet row spec.
 *
 * Options are quoted per share by Flex; the sheet carries price per CONTRACT so
 * that quantity stays in contracts and matches your brokerage statement. Hence
 * price = markPrice * multiplier.
 *
 * @param {Object} attrs - attribute map from an <OpenPosition> element
 * @returns {{ticker: string, quantity: number, price: number, category: string, isOption: boolean}|null}
 */
function _normaliseFlexPosition(attrs) {
  if (!attrs) return null;

  const cat = String(attrs.assetCategory || "").toUpperCase();
  const mark = Number(attrs.markPrice);
  const isOption = (cat === "OPT" || cat === "FOP");
  const multiplier = Number(attrs.multiplier) || (isOption ? 100 : 1);

  // Quantity is only present when the Flex Query includes the Position field.
  // When it is absent, derive it: positionValue = qty * markPrice * multiplier.
  // Reporting a position with an unknowable quantity as "closed" would be far
  // worse than refusing, so return null and let the caller count the failure.
  let qty = Number(attrs.position);
  if (!isFinite(qty)) {
    const value = Number(attrs.positionValue);
    if (isFinite(value) && isFinite(mark) && mark !== 0 && multiplier !== 0) {
      qty = value / (mark * multiplier);
      // Share counts are whole or finely fractional; scrub float noise.
      qty = Math.abs(qty - Math.round(qty)) < 1e-6 ? Math.round(qty) : Number(qty.toFixed(6));
    }
  }
  if (!isFinite(qty) || qty === 0) return null;

  const ticker = isOption
    ? _buildOccSymbol(attrs.underlyingSymbol || attrs.symbol, attrs.expiry || attrs.expirationDate,
                      attrs.putCall, attrs.strike)
    : String(attrs.symbol || "").trim();

  if (!ticker) return null;

  return {
    ticker: ticker,
    quantity: qty,
    price: isFinite(mark) ? mark * (isOption ? multiplier : 1) : null,
    category: isOption ? "Option" : (cat === "STK" ? "Stock" : (cat || "Other")),
    isOption: isOption
  };
}

/**
 * Turns a Flex cash balance into a sheet row spec.
 * USD cash is priced at 1; foreign cash gets a live FX formula so the sheet
 * keeps converting after the sync run.
 *
 * @param {string} currency
 * @param {number} amount
 * @returns {{ticker: string, quantity: number, priceFormula: string, category: string}|null}
 */
function _normaliseFlexCash(currency, amount) {
  const cur = String(currency || "").toUpperCase().trim();
  const qty = Number(amount);
  if (!cur || cur === "BASE_SUMMARY" || !isFinite(qty) || qty === 0) return null;

  return {
    ticker: "Cash",
    quantity: qty,
    priceFormula: cur === "USD" ? "=1" : `=IFERROR(GOOGLEFINANCE("CURRENCY:${cur}USD"),1)`,
    category: "Cash",
    currency: cur
  };
}

/**
 * Plans a wholesale rebuild of one account's block on the Holdings grid.
 *
 * The feed is authoritative for the account it owns, so the block is rewritten
 * rather than reconciled row by row. Match-and-update produced a worse failure
 * mode: a feed that returned nothing looked identical to every position having
 * been closed, and quietly zeroed the lot.
 *
 * Guards, in order:
 *  - a feed carrying no positions is a failed feed, never an emptied account
 *  - a feed that would shrink the block by more than SYNC_MAX_SHRINK is
 *    reported as unsafe and applied only on explicit confirmation
 *  - the target block must be contiguous and consist only of rows this account
 *    already owns plus blank rows, so another account is never overwritten
 *
 * An account with no existing rows is bootstrapped into the first writable run
 * long enough to hold it — nothing has to be laid out by hand first. If the
 * account's current position can't fit the feed, the block relocates to a run
 * that can, and the old rows are cleared.
 *
 * @param {Array<Array<*>>} existingRows - A..G values starting at firstRow
 * @param {Array<Object>} positions - normalised specs
 * @param {string} accountName
 * @param {number} firstRow
 * @param {number} lastRow
 * @returns {{ok: boolean, reason: string, startRow: number, writes: Array, clears: Array, removed: Array, shrinkRatio: number}}
 */
function _planBlockSync(existingRows, positions, accountName, firstRow, lastRow) {
  const acct = String(accountName).trim();
  const fail = reason => ({ ok: false, reason: reason, startRow: 0, writes: [], clears: [], removed: [], shrinkRatio: 0 });

  if (!positions || !positions.length) {
    return fail("The feed returned no positions. Nothing was changed — a feed carrying " +
                "nothing is a failed request, not an emptied account.");
  }

  const ownedSet = {}, blankSet = {};
  const owned = [];
  for (let i = 0; i < existingRows.length; i++) {
    const row = firstRow + i;
    const a = String(existingRows[i][0] || "").trim();
    if (a === acct) {
      owned.push({ row: row, ticker: String(existingRows[i][2] || "").trim() });
      ownedSet[row] = true;
    } else if (!a) {
      blankSet[row] = true;
    }
  }

  const needed = positions.length;
  const writable = r => !!(ownedSet[r] || blankSet[r]);

  /** Length of the writable run starting at `start`. */
  const runFrom = start => {
    let n = 0;
    for (let r = start; r <= lastRow && writable(r); r++) n++;
    return n;
  };

  // Prefer to rewrite in place, anchored on the account's first existing row.
  // Otherwise look for any writable run long enough — which is also how an
  // account with NO existing rows gets bootstrapped.
  let startRow = null, relocated = false;

  if (owned.length && runFrom(owned[0].row) >= needed) {
    startRow = owned[0].row;
  } else {
    for (let r = firstRow; r <= lastRow - needed + 1; r++) {
      if (runFrom(r) >= needed) { startRow = r; relocated = owned.length > 0; break; }
    }
  }

  if (startRow === null) {
    return fail(`Need ${needed} consecutive rows for "${acct}" but the Holdings tab has ` +
                `no run that long. Free up space (or add rows) and re-run.`);
  }

  const bootstrap = owned.length === 0;
  const endRow = startRow + needed - 1;
  const writes = positions.map((spec, i) => ({ row: startRow + i, spec: spec }));

  // Owned rows outside the new block are stale and get cleared.
  const clears = [], removed = [];
  owned.forEach(o => {
    if (o.row >= startRow && o.row <= endRow) return;
    clears.push(o.row);
    if (o.ticker) removed.push(o.ticker);
  });

  const shrinkRatio = owned.length ? Math.max(0, (owned.length - needed) / owned.length) : 0;

  return { ok: true, reason: "", startRow: startRow, writes: writes, clears: clears,
           removed: removed, shrinkRatio: shrinkRatio, bootstrap: bootstrap, relocated: relocated };
}

/**
 * Plans a sync against the existing Holdings grid.
 *
 * Matching key is (account, ticker). For cash the key includes currency, since
 * one account can hold several currencies all tickered "Cash".
 *
 * @param {Array<Array<*>>} existingRows - A..G values starting at firstRow
 * @param {Array<Object>} positions - normalised specs
 * @param {string} accountName - the Holdings column A value this sync owns
 * @param {number} firstRow
 * @param {number} maxRow - last usable row
 * @returns {{updates: Array, additions: Array, zeroed: Array, notes: Array}}
 */
function _planHoldingsSync(existingRows, positions, accountName, firstRow, maxRow) {
  const updates = [], additions = [], zeroed = [], notes = [];
  const acct = String(accountName).trim();
  const keyOf = p => `${p.ticker}||${p.currency || ""}`;

  // Index existing rows owned by this account.
  const owned = {};
  const firstFree = [];
  for (let i = 0; i < existingRows.length; i++) {
    const row = firstRow + i;
    const rowAcct = String(existingRows[i][0] || "").trim();
    if (!rowAcct) { firstFree.push(row); continue; }
    if (rowAcct !== acct) continue;
    const ticker = String(existingRows[i][2] || "").trim();
    const cur = String(existingRows[i][6] || "").trim().toUpperCase();
    owned[`${ticker}||${ticker === "Cash" ? cur : ""}`] = { row: row, quantity: existingRows[i][3] };
  }

  const seen = {};
  positions.forEach(p => {
    const key = keyOf(p);
    seen[key] = true;
    const match = owned[key];

    if (match) {
      updates.push({ row: match.row, spec: p, wasQuantity: match.quantity });
    } else {
      const row = firstFree.shift();
      if (row === undefined || row > maxRow) {
        notes.push(`No free row for ${p.ticker} — add rows to the Holdings tab and re-run.`);
        return;
      }
      additions.push({ row: row, spec: p });
    }
  });

  // Positions that vanished: zero the quantity, never delete the row. A closed
  // position is information; a missing row is a gap you can't audit.
  Object.keys(owned).forEach(key => {
    if (seen[key]) return;
    if (Number(owned[key].quantity) === 0) return;
    zeroed.push({ row: owned[key].row, ticker: key.split("||")[0] });
  });

  return { updates: updates, additions: additions, zeroed: zeroed, notes: notes };
}

/**
 * Pure helper: extracts attribute maps for a given element name from Flex XML.
 * Written against a string rather than XmlService so it is unit testable.
 *
 * @param {string} xml
 * @param {string} tagName - e.g. "OpenPosition"
 * @returns {Array<Object<string,string>>}
 */
function _extractFlexElements(xml, tagName) {
  const out = [];
  const re = new RegExp(`<${tagName}\\b([^>]*?)/?>`, "g");
  let m;
  while ((m = re.exec(String(xml || ""))) !== null) {
    const attrs = {};
    const attrRe = /([A-Za-z_][\w:.-]*)\s*=\s*"([^"]*)"/g;
    let a;
    while ((a = attrRe.exec(m[1])) !== null) attrs[a[1]] = a[2];
    out.push(attrs);
  }
  return out;
}

/**
 * Pure helper: reads Status / ErrorCode / ErrorMessage from a Flex response.
 * @param {string} xml
 * @returns {{ok: boolean, referenceCode: string, error: string}}
 */
function _parseFlexEnvelope(xml) {
  const pick = tag => {
    const m = new RegExp(`<${tag}>([^<]*)</${tag}>`).exec(String(xml || ""));
    return m ? m[1].trim() : "";
  };
  const status = pick("Status");
  const code = pick("ErrorCode");
  const message = pick("ErrorMessage");

  return {
    ok: status === "Success",
    referenceCode: pick("ReferenceCode"),
    error: status === "Success" ? "" : (code ? `${code}: ${message}` : (message || "Unknown Flex error"))
  };
}

/* ------------------------------------------------------------------ *
 * WIZARD & FETCH
 * ------------------------------------------------------------------ */

/** True when a Flex token, query id and account name are all configured. */
function isIbkrFlexConfigured() {
  const p = PropertiesService.getDocumentProperties();
  return !!(p.getProperty(FLEX_PROP_TOKEN) && p.getProperty(FLEX_PROP_QUERY) && p.getProperty(FLEX_PROP_ACCOUNT));
}

/**
 * Menu action: collects Flex credentials.
 * Stored in DocumentProperties, never in a cell — a token in a sheet is visible
 * to every collaborator and would ride along into the Gist/Drive backups.
 */
function setupIbkrFlexWizard() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const tokenRes = ui.prompt("IBKR Flex — step 1 of 3",
    "Paste your Flex Web Service token.\n\n" +
    "Client Portal > Performance & Reports > Flex Queries >\n" +
    "Flex Web Service Configuration > Generate New Token.\n\n" +
    "IMPORTANT: set 'Should Expire After' to the longest option. The\n" +
    "default is 6 hours, which will break scheduled syncs.\n" +
    "Leave 'Valid For IP Address' BLANK — Google's servers change IP.",
    ui.ButtonSet.OK_CANCEL);
  if (tokenRes.getSelectedButton() !== ui.Button.OK) return;
  const token = String(tokenRes.getResponseText()).trim();
  if (!token) { ui.alert("No token entered — setup cancelled."); return; }

  const queryRes = ui.prompt("IBKR Flex — step 2 of 3",
    "Paste your Flex Query ID (the numeric id shown next to the query).",
    ui.ButtonSet.OK_CANCEL);
  if (queryRes.getSelectedButton() !== ui.Button.OK) return;
  const queryId = String(queryRes.getResponseText()).trim();

  const acctRes = ui.prompt("IBKR Flex — step 3 of 3",
    "Which Account Name on the Brokerage Holdings tab does this feed own?\n\n" +
    "Must match column A exactly (e.g. IBKR). Sync only ever writes to rows\n" +
    "carrying this name; everything else is left alone.",
    ui.ButtonSet.OK_CANCEL);
  if (acctRes.getSelectedButton() !== ui.Button.OK) return;
  const accountName = String(acctRes.getResponseText()).trim();

  if (!queryId || !accountName) { ui.alert("Incomplete — setup cancelled."); return; }

  props.setProperty(FLEX_PROP_TOKEN, token);
  props.setProperty(FLEX_PROP_QUERY, queryId);
  props.setProperty(FLEX_PROP_ACCOUNT, accountName);

  ui.alert("✅ IBKR Flex configured.\n\n" +
    "Reload the sheet to see '🔗 Sync IBKR Positions' in the menu.\n\n" +
    "Run it manually first and read the report before scheduling it.");
}

/**
 * Two-step Flex fetch: request a report, wait, then retrieve it.
 * @returns {{ok: boolean, xml: string, error: string}}
 */
function _fetchFlexStatement() {
  const props = PropertiesService.getDocumentProperties();
  const token = props.getProperty(FLEX_PROP_TOKEN);
  const queryId = props.getProperty(FLEX_PROP_QUERY);
  if (!token || !queryId) return { ok: false, xml: "", error: "IBKR Flex is not configured." };

  // Flex rejects requests without a User-Agent identifying the client.
  const opts = { method: "get", muteHttpExceptions: true, headers: { "User-Agent": "GoogleAppsScript/1.0" } };

  try {
    const sendUrl = `${FLEX_BASE}/SendRequest?t=${encodeURIComponent(token)}&q=${encodeURIComponent(queryId)}&v=${FLEX_VERSION}`;
    const sendXml = UrlFetchApp.fetch(sendUrl, opts).getContentText();
    const env = _parseFlexEnvelope(sendXml);
    if (!env.ok) return { ok: false, xml: "", error: env.error };

    // The report generates asynchronously; large queries need longer.
    for (let attempt = 0; attempt < 5; attempt++) {
      Utilities.sleep(5000);
      const getUrl = `${FLEX_BASE}/GetStatement?t=${encodeURIComponent(token)}&q=${encodeURIComponent(env.referenceCode)}&v=${FLEX_VERSION}`;
      const xml = UrlFetchApp.fetch(getUrl, opts).getContentText();
      if (xml.indexOf("<FlexQueryResponse") > -1) return { ok: true, xml: xml, error: "" };

      const err = _parseFlexEnvelope(xml);
      // 1019 = statement still generating; anything else is terminal.
      if (err.error.indexOf("1019") === -1) return { ok: false, xml: "", error: err.error };
    }
    return { ok: false, xml: "", error: "Flex report did not finish generating in time. Try again." };
  } catch (e) {
    return { ok: false, xml: "", error: String(e) };
  }
}

/* ------------------------------------------------------------------ *
 * SYNC
 * ------------------------------------------------------------------ */

/**
 * Menu action: rebuilds this account's block on the Brokerage Holdings tab.
 *
 * Column B is never written. In the stock layout it is a plain text category;
 * many users replace it with their own classification formula, and the sync has
 * no business overwriting either.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{written: number, cleared: number, notes: Array<string>}}
 */
function syncIbkrPositions(ss_inject, silent = false) {
  silent = _resolveSilent(silent, _isInteractive());
  const ss = _resolveSpreadsheet(ss_inject);
  const sheet = ss.getSheetByName("Brokerage Holdings");
  const props = PropertiesService.getDocumentProperties();
  const accountName = props.getProperty(FLEX_PROP_ACCOUNT);
  const ui = silent ? null : SpreadsheetApp.getUi();

  const bail = notes => ({ written: 0, cleared: 0, notes: notes });

  if (!sheet) return bail(["Brokerage Holdings tab not found."]);
  if (!accountName) return bail(["IBKR Flex is not configured."]);

  const headerProblems = _auditHeaders(ss);
  if (headerProblems.length) {
    const msg = ["🛑 Sync aborted — the Holdings header row doesn't match the expected layout.", ""]
      .concat(headerProblems.map(h => "  • " + h))
      .concat(["", "Sync writes by column position. Fix the headers first."]).join("\n");
    if (ui) ui.alert(msg);
    return bail(headerProblems);
  }

  const fetched = _fetchFlexStatement();
  if (!fetched.ok) {
    if (ui) ui.alert(`🛑 IBKR Flex request failed.\n\n${fetched.error}`);
    return bail([fetched.error]);
  }

  const positions = [];
  let unparseable = 0;
  _extractFlexElements(fetched.xml, "OpenPosition").forEach(a => {
    const p = _normaliseFlexPosition(a);
    if (p) positions.push(p); else unparseable++;
  });
  _extractFlexElements(fetched.xml, "CashReportCurrency").forEach(a => {
    const c = _normaliseFlexCash(a.currency, a.endingCash || a.endingSettledCash);
    if (c) positions.push(c);
  });

  // If the report carried positions we could not read, stop. Writing a partial
  // block would silently drop whatever failed to parse.
  if (unparseable > 0) {
    const msg = [`🛑 Sync aborted — ${unparseable} position(s) in the Flex report could not be read.`, "",
      "Usually the Flex Query is missing a field. In Client Portal, edit the query and",
      "make sure the Open Positions section includes Quantity (Position), MarkPrice,",
      "Multiplier, Symbol, UnderlyingSymbol, AssetCategory, Strike, Expiry and Put/Call.",
      "", "Nothing was changed."].join("\n");
    if (ui) ui.alert(msg);
    return bail([`${unparseable} unparseable positions`]);
  }

  const rows = sheet.getRange(HOLDINGS_FIRST_ROW, 1, HOLDINGS_NUM_ROWS, 7).getValues();
  const plan = _planBlockSync(rows, positions, accountName, HOLDINGS_FIRST_ROW, HOLDINGS_LAST_ROW);

  if (!plan.ok) {
    if (ui) ui.alert(`🛑 Sync aborted.\n\n${plan.reason}`);
    return bail([plan.reason]);
  }

  if (plan.shrinkRatio > SYNC_MAX_SHRINK) {
    const pct = Math.round(plan.shrinkRatio * 100);
    const question = [`⚠️ This sync would remove ${pct}% of the ${accountName} rows.`, "",
      `Positions in the feed: ${positions.length}`,
      `Rows to be cleared: ${plan.clears.length}`, "",
      plan.removed.length ? "Tickers disappearing: " + plan.removed.slice(0, 20).join(", ") : "",
      "", "A drop this large usually means an incomplete feed rather than closed",
      "positions. Proceed anyway?"].join("\n");
    if (!ui) return bail([`Refused: would remove ${pct}% of rows (silent mode)`]);
    if (ui.alert("Unusually large change", question, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      ui.alert("Cancelled — nothing was changed.");
      return bail(["Cancelled by user"]);
    }
  }

  plan.writes.forEach(w => {
    const spec = w.spec;
    sheet.getRange(w.row, 1).setValue(accountName);
    sheet.getRange(w.row, 3).setValue(spec.ticker);
    sheet.getRange(w.row, 4).setValue(spec.quantity);
    sheet.getRange(w.row, 7).setValue(spec.currency || "");

    if (spec.priceFormula) {
      sheet.getRange(w.row, 5).setFormula(spec.priceFormula);
    } else if (spec.isOption && spec.price !== null) {
      // GOOGLEFINANCE cannot price an OCC symbol, so the mark is a literal.
      // Refreshing it is the whole point of the sync.
      sheet.getRange(w.row, 5).setValue(spec.price);
    } else {
      // Equities keep a live formula so the sheet stays current between syncs
      // rather than freezing at the last statement's close.
      sheet.getRange(w.row, 5).setFormula(_holdingsRowFormulas(w.row).E);
    }
    sheet.getRange(w.row, 6).setFormula(_holdingsRowFormulas(w.row).F);
  });

  // Clear the DATA but restore the canonical formulas. Column F is a managed
  // range, so a bare clearContent() drops the formula count and makes the next
  // snapshot look like formula damage — the integrity check can't tell a
  // deliberate clear from a destructive one, and shouldn't have to.
  plan.clears.forEach(r => {
    sheet.getRange(r, 1, 1, 4).clearContent();   // A-D: account, category, ticker, quantity
    sheet.getRange(r, 7).clearContent();         // G: currency
    sheet.getRange(r, 5).setFormula(_holdingsRowFormulas(r).E);
    sheet.getRange(r, 6).setFormula(_holdingsRowFormulas(r).F);
  });

  // Sync legitimately reshapes the grid, so the accepted baseline moves with it.
  _saveFormulaBaseline(ss);

  if (ui) {
    const headline = plan.bootstrap
      ? `✅ Created the ${accountName} block from IBKR.`
      : (plan.relocated
          ? `✅ Rebuilt and MOVED the ${accountName} block from IBKR.`
          : `✅ Rebuilt the ${accountName} block from IBKR.`);
    const lines = [headline, "",
      `Rows written: ${plan.writes.length} (rows ${plan.startRow}–${plan.startRow + plan.writes.length - 1})`,
      `Stale rows cleared: ${plan.clears.length}`];
    if (plan.relocated) {
      lines.push("", "The block moved because it outgrew its previous position.",
                     "Check that the Dashboard row for this account still totals correctly.");
    }
    if (plan.removed.length) {
      lines.push("", "No longer held:", ...plan.removed.slice(0, 15).map(t => "  • " + t));
    }
    lines.push("", "Option marks are end-of-day from the Flex statement.",
                   "Equities keep live GOOGLEFINANCE prices.");
    ui.alert(lines.join("\n"));
  }

  return { written: plan.writes.length, cleared: plan.clears.length, notes: plan.notes || [] };
}
