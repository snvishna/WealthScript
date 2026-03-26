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
