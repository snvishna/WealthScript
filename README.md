# WealthScript 💰

A privacy-first, code-driven wealth tracker built entirely inside Google Sheets via Google Apps Script.

> **Why WealthScript?** Unlike Mint, Empower, or Monarch, WealthScript does **not** require handing your bank credentials to third-party aggregators like Plaid. Your financial data never leaves your private Google Drive. You manually input your balances — ensuring absolute privacy, security, and a hands-on relationship with your personal finances.

---

## ✨ Core Features

### 📊 Dashboard & Ledger
* **Professional KPI Banner** — Dark-themed title row with 3 live-computed cards:
  * **Card 1 (USD):** Net Worth and Gross Worth via `SUMIFS` over your active accounts.
  * **Cards 2 & 3 (Configurable):** Any two secondary currencies (default: CAD, INR). Change them in the **Settings & Config** tab — the dashboard updates instantly via `GOOGLEFINANCE`.
* **Abbreviated Number Formats** — Values display as `$1.55M`, `CA$320K`, `₹25M` for at-a-glance readability.
* **Quick-Stats Row** — 🌊 Liquid Net Worth, 🔒 Locked Net Worth, and 🔥 FIRE Progress (%) — all as live formulas.
* **Color-Coded Asset Classes** — 10 distinct background colors auto-applied to Cash, Brokerage, Crypto, Retirement, Health Savings, Real Estate, Commodity, Insurance, Receivable, and Liability rows.
* **Conditional Formatting** — Negative net worth cells are highlighted red.
* **Computed Columns** — Exchange Rate (via `GOOGLEFINANCE`), Gross Worth (USD), and Net Worth (USD) are all formula-driven for 70 rows.

### 📈 Brokerage Holdings
* Separate tab for stock/ETF/crypto positions with live `GOOGLEFINANCE` pricing.
* Columns: Account Name, Ticker, Quantity, Live Price, Total Value — all auto-computed.

### 📸 Snapshot Engine
A single click (`WealthScript > 📸 Log Snapshot & Cloud Sync`) performs **four actions**:
1. **Calculates** your net worth, gross worth, liquid/locked split, and real estate totals.
2. **Logs** a new row in the `Snapshots` tab with the date, all metrics, dollar delta, % growth, FIRE progress, and an auto-generated plain-English commentary (e.g., *"Net worth increased by $12,000. Liquid pool +$3,500."*).
3. **Backs up** your entire ledger to both GitHub Gist and Google Drive (see Disaster Recovery below).
4. **Refreshes** the visual charts on the dashboard.

### 💸 Cash Flow & Burn Rate
* Dedicated `💸 Cash Flow & Burn` tab with a manual expense ledger (Date, Category, Amount, Notes).
* KPI Summary: Average Monthly Burn, TTM (Trailing 12-Month) Expenses, Target FIRE Budget (from Settings), and Current Safe Withdrawal Rate (4% Rule).
* **Important:** Expenses are entered manually — WealthScript does not connect to bank accounts or credit cards. This is intentional for privacy.

### 🏠 Automated Real Estate Pricing
* Integrates with RapidAPI (Zillow Zestimate) to auto-update property values via a weekly cron trigger.
* Map your Account Name → ZPID in the Settings tab. The script fetches Zestimates and updates Current Value in bulk.

### 📉 Visual Dashboards
* Programmatic charting via `updateVisualDashboards()`:
  * **Donut Chart** — Asset class allocation breakdown.
  * **Area Chart** — Historical net worth trajectory (smooth curves).
  * **Stacked Column Chart** — Liquid vs. Locked assets over time.

---

## ☁️ Dual Disaster Recovery

Every snapshot silently triggers **two** independent backups:

### 1. GitHub Gist (Developer-Friendly)
* Your entire ledger is serialized to JSON and pushed to a **private GitHub Gist** via the GitHub API.
* Provides full version-controlled diff history — you can see exactly what changed each month.
* **Setup (Guided Wizard):**
  1. Click **WealthScript > 🔐 Setup GitHub Backup** from the menu bar.
  2. A new browser tab opens to GitHub's token creation page with `gist` scope pre-selected.
  3. Click "Generate token", copy it, and paste it into the dialog that appears back in your Sheet.
  4. The wizard validates the token, auto-creates a private Gist, and populates your Settings tab with the credentials and a **clickable hyperlink** to your live Gist.
* If no PAT is configured, Gist backup silently skips — no errors.

### 2. Google Drive (Zero-Config)
* A dated JSON file (e.g., `net_worth_2026-03-24T20-00.json`) is saved to a `WealthScript — Backups` folder in your Google Drive.
* **No API keys or tokens needed** — uses your existing Google auth context.
* **Setup:** Click **WealthScript > 📁 Setup Google Drive Backup** to create the folder and add a **clickable hyperlink** to your Settings tab.
* Files are ~2KB each; even 5 years of monthly snapshots is under 2MB.

### Manual Backup
From the **WealthScript** menu, click **☁️ Force Cloud Backup** to trigger both methods on-demand with UI confirmation alerts.

---

## 🚀 Quick Start Guide

### Phase 1: Deploy the Engine
1. Create a blank [Google Sheet](https://sheets.new).
2. Go to **Extensions > Apps Script**.
3. Delete the default `Code.gs` file. Paste the **entire contents** of [`deploy/code.gs`](./deploy/code.gs) into the editor.
   *(This distribution file is auto-compiled from `src/` modules for easy installation.)*
4. Click **Save** (💾). Close the Apps Script tab. **Refresh the sheet.**

### Phase 2: First Time Setup
1. A **WealthScript** menu will appear in your menu bar.
2. Click **WealthScript > 🚀 Run First Time Setup**.
3. Authorize the script when prompted. You will see a standard Google OAuth consent screen.

   > **Note on the OAuth prompt:** Google may show an "unverified app" warning because WealthScript is an open-source project, not a published Google Workspace Add-on. This is expected — click **"Advanced" → "Go to [project] (unsafe)"** to proceed. You are authorizing your *own* code running in your *own* Google account.
   >
   > The consent screen will list these permissions — all are required:
   > - *"See, edit, create and delete your Sheets"* — to build and update your tracker.
   > - *"See, edit, create and delete files in Google Drive"* — to create the backup folder and files.
   > - *"Connect to an external service"* — for GOOGLEFINANCE, GitHub, and RapidAPI.
   > - *"Allow this app to run when you are not present"* — for the weekly Real Estate cron.
   > - *"Display third-party web content"* — for the setup wizard dialog.

4. The script builds all tabs and sets up a weekly cron for real estate updates.


### Phase 3: Configure & Customize
1. **Settings & Config** tab — paste API credentials, adjust FIRE targets, and set your dashboard currencies.
2. **Brokerage Holdings** tab — add your stock/crypto tickers and quantities.
3. **Dashboard & Ledger** tab — replace sample accounts with your real assets and balances.
4. **Cash Flow & Burn** tab — replace sample expenses with your actual monthly outlays.

### Settings & Config — Full Reference

| Section | Cell | Field | Default |
|---------|------|-------|---------|
| Real Estate API | B9 | RapidAPI Key | `PASTE_KEY_HERE` |
| Real Estate API | B10 | RapidAPI Host | `real-estate101.p.rapidapi.com` |
| Cloud Backup | B13 | GitHub PAT (gist scope) | — |
| Cloud Backup | B14 | GitHub Gist ID | *(auto-created by wizard)* |
| Cloud Backup | B15 | GitHub Gist URL | *(clickable hyperlink, set by wizard)* |
| Cloud Backup | B16 | Google Drive Backup Folder | *(clickable hyperlink, set by wizard)* |
| FIRE & Cash Flow | B19 | Target Monthly FIRE Budget (USD) | `$20,000` |
| FIRE & Cash Flow | B20 | Estimated Monthly Rental Income (USD) | `$0` |
| FIRE & Cash Flow | B21 | Annual Portfolio Return Rate | `7.00%` |
| Dashboard Currency | B23 | Secondary Currency (Card 2) | `CAD` |
| Dashboard Currency | B24 | Secondary Currency (Card 3) | `INR` |
| ZPID Mapping | A28:B44 | Property Name → ZPID pairs | sample data |

### Phase 4: Enable Native Tables & Grouping
1. On `Dashboard & Ledger`, select all data rows (Row 7+).
2. Click **Format > Convert to Table**.
3. In the Table dropdown: enable *gridlines*, *alternating colors*, *condensed view*.
4. **Create group by view** → choose **Asset Class**.
5. Right-click column headers for Gross/Net Worth → set calculation to **Sum**.

---

## 🏠 Setting up Automated Real Estate (Zestimates)

1. Create a free [RapidAPI](https://rapidapi.com) account.
2. Subscribe to the `Zillow.com API` (or `real-estate101`) — free tier gives 20-50 calls/month.
3. Paste your `X-RapidAPI-Key` into the Settings tab.
4. **Find your ZPID:** On Zillow, look at your property URL: `zillow.com/homedetails/…/87654321_zpid/`. The number is your ZPID.
5. Add your Account Name (must match the ledger exactly!) and ZPID to the mapping table.

---

## 🛠 Developer Notes

### Customizing Default Accounts
Modify the `DEFAULT_PORTFOLIO_DATA` array at the top of `code.gs` before running First Time Setup. Each row follows the format:
```
["Account Name", "Asset Class", "Currency", Initial Capital, Current Value, "", Tax Rate, "", "", "Status", "Remarks"]
```

### Customizing Dashboard Currencies
Edit the `DASHBOARD_CONFIG` constant in `code.gs`:
```javascript
const DASHBOARD_CONFIG = {
  secondaryCurrencies: ["CAD", "INR"], // Any GOOGLEFINANCE code: EUR, GBP, AUD…
  fireTargetUSD: 3000000,
};
```
Or change them live in the **Settings & Config** tab (rows B14/B15) — the dashboard references those cells directly.

### WealthScript Menu Actions

| Menu Item | Function | Description |
|-----------|----------|-------------|
| 🚀 Run First Time Setup | `runFirstTimeSetup()` | Builds all tabs, sets up cron triggers |
| 📸 Log Snapshot & Cloud Sync | `captureSnapshot()` | Logs snapshot, backs up to Gist + Drive, refreshes charts |
| 🔄 Refresh Real Estate Prices | `updateRealEstatePrices()` | Fetches Zestimates via RapidAPI |
| 📊 Update Visual Dashboards | `updateVisualDashboards()` | Repaints donut, area, and stacked bar charts |
| 💸 Rebuild Cash Flow Tab | `buildCashFlowTab()` | Rebuilds the Cash Flow & Burn tab from scratch |
| ☁️ Force Cloud Backup | `forceBackup()` | Runs both Gist and Drive backups with UI alerts |

### Test Suite (`tests.gs`)
The repository includes an isolated unit test suite with **52 assertions across 10 test suites**. All tests are pure — no live Sheets or network calls.

| Suite | What It Tests |
|-------|---------------|
| `test_calcGrowthDelta` | Net worth delta and % growth calculation |
| `test_calcFireProgress` | FIRE progress ratio (net worth / target) |
| `test_classifyAsset` | Liquid vs. locked vs. skip classification |
| `test_classifyAsset_extended` | Edge cases: HSA, Insurance, Commodity, Liability, negative values |
| `test_cashFlowKpis` | Safe Withdrawal Rate + TTM expense aggregation |
| `test_buildLedgerSnapshot` | JSON serialization: blank-row skipping, field fidelity, edge inputs |
| `test_driveBackupPruning` | Pruning threshold logic for backup file management |
| `test_currencySymbol` | Currency code → symbol mapping (10 codes + unknown fallback) |
| `test_abbrFmt` | Abbreviated number format string generation |
| `test_generateInsight` | Snapshot auto-commentary (positive, negative, zero delta, first snapshot) |

**To run:** Open `tests.gs` in the Apps Script editor → select `runAllTests` → click ▶ Run → check View → Logs.

### Project File Structure
| File | Purpose |
|------|---------|
| `code.gs` | Main application — tab builders, snapshot engine, backup sync, API integrations |
| `tests.gs` | 52-assertion isolated unit test suite |
| `README.md` | This file |
| `docs/specs/` | Feature specs and architecture documents |

---

## 🔐 OAuth Scopes & Privacy

WealthScript uses a minimal, explicitly declared permission manifest (`deploy/appsscript.json`). Applying it in Step 4 of the Quick Start Guide **replaces auto-detected broad scopes with these 4 specific ones:**

| Permission shown on consent screen | Scope declared | Why it's needed |
|---|---|---|
| *"See, edit, create and delete your Sheets"* | `spreadsheets` | Builds tabs, writes balances, updates formulas |
| *"Connect to an external service"* | `script.external_request` | GOOGLEFINANCE, GitHub Gist API, RapidAPI (Zillow) |
| *"Allow this app to run when you are not present"* | `script.scriptapp` | Weekly cron trigger for Real Estate price updates |
| *"Display third-party web content in sidebars"* | `script.container.ui` | Setup wizard dialog (GitHub + Drive onboarding) |

> **Critically: WealthScript does NOT request full Google Drive access.**
> The `drive.file` scope is used for backups \u2014 which means the app can only see and modify files **it created itself**. It cannot read, list, or modify any other file in your Google Drive.

> **On the "unverified app" warning:** WealthScript is an open-source personal tool, not a submitted Google Workspace Add-on. Google shows this warning for any self-deployed script that isn't registered through their OAuth verification program. You are running your own code in your own account \u2014 clicking "Advanced \u2192 Go to project (unsafe)" is safe and expected.
