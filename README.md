# WealthScript 💰

A privacy-first, code-driven wealth tracker built entirely inside Google Sheets via Google Apps Script.

> **Why WealthScript?** Unlike Mint, Empower, or Monarch, WealthScript does **not** require handing your bank credentials to third-party aggregators like Plaid. Your financial data never leaves your private Google Drive. You manually input your balances — ensuring absolute privacy, security, and a hands-on relationship with your personal finances.

---

## ✨ Core Features

### 📊 Dashboard & Ledger
* **Professional KPI Banner** — Dark-themed title row with 3 live-computed cards:
  * **Card 1 (USD):** Net Worth and Gross Worth via `SUMIFS` over your active accounts.
  * **Cards 2 & 3 (Configurable):** Any two secondary currencies (default: CAD, INR). Change them in the **Settings & Config** tab — the dashboard updates instantly via `GOOGLEFINANCE`.
* **Locale-Aware Abbreviated Formats** — Values display as `$1.55M`, `CA$320K`, and — for Indian-numbering currencies (INR, PKR, LKR, NPR, BDT) — `₹39Cr` / `₹4.5L` rather than millions. Cards 2 & 3 render as *text*; their numeric values live in hidden helper columns `M:O`, which the snapshot and backup engines read. Do not delete those columns.
* **Quick-Stats Row** — 🌊 Liquid Net Worth, 🔒 Locked Net Worth, and 🔥 FIRE Progress (%) — all as live formulas.
* **Color-Coded Asset Classes** — 10 distinct background colors auto-applied to Cash, Brokerage, Crypto, Retirement, Health Savings, Real Estate, Commodity, Insurance, Receivable, and Liability rows.
* **Conditional Formatting** — Negative net worth cells are highlighted red.
* **Computed Columns** — Exchange Rate (via `GOOGLEFINANCE`), Gross Worth (USD), and Net Worth (USD) are all formula-driven for 70 rows.

### 📈 Brokerage Holdings
* Separate tab for stock/ETF/crypto positions with live `GOOGLEFINANCE` pricing.
* Columns: Account Name, Ticker, Quantity, Live Price, Total Value — all auto-computed.

### 🔗 Broker Sync (optional)

WealthScript is broker-agnostic by default — most users maintain the Brokerage
Holdings tab by hand, and nothing here changes that. If you happen to use
Interactive Brokers, you can connect the **Flex Web Service** to refresh one
account's positions automatically.

- **Opt-in.** The sync action only appears in the menu once you've configured a
  provider. Users who don't connect one see a single unobtrusive
  `🔗 Connect IBKR (optional)` entry.
- **Scoped.** A sync owns exactly one Account Name on the Holdings tab. Rows
  belonging to any other account are never read or written.
- **Block rebuild, with guards.** The feed is authoritative for the account it
  owns, so its block is rewritten wholesale. A feed carrying no positions is
  treated as a failed request, never an emptied account. A sync that would shrink
  the block by more than 30% asks first. Any position the parser can't read
  aborts the whole run rather than writing a partial block.
- **Hybrid pricing.** Options get a literal mark from the statement, because
  GOOGLEFINANCE cannot price an OCC symbol. Everything else keeps a live
  GOOGLEFINANCE formula so the sheet stays current between syncs instead of
  freezing at the last close.
- **Column B is never written**, so your own asset-classification formula
  survives.
- **Credentials never touch a cell.** The Flex token lives in
  DocumentProperties, so it isn't visible to collaborators and doesn't ride
  along into the Gist/Drive backups.

#### Setting up IBKR Flex — step by step

**A. Create the Flex Query**

1. Log in to [Client Portal](https://www.interactivebrokers.com/sso/Login).
2. **Performance & Reports → Flex Queries → Create Activity Flex Query** (the `+` button).
3. Name it something recognisable, e.g. `WealthScript Refresh Query`.
4. Open the **Open Positions** section and enable these fields:
   `Symbol`, `UnderlyingSymbol`, `AssetCategory`, `Position`, `MarkPrice`,
   `Multiplier`, `Strike`, `Expiry`, `PutCall`, `Currency`, `PositionValue`.
   Set *Options* to **Summary** (not Lot-level) so each holding is one row.
5. Open the **Cash Report** section and enable `Currency` and `EndingCash`.
6. Delivery configuration: **Format = XML**, **Period = Last Business Day**.
7. Save. The query list now shows a numeric **Query ID** — note it down.

> If `Position` is missing, quantity is derived from
> `positionValue / (markPrice × multiplier)`. That works, but enabling the field
> is more robust — a position with a zero mark can't be derived.

**B. Enable the Flex Web Service and generate a token**

8. **Performance & Reports → Flex Queries → Flex Web Service Configuration.**
9. Tick the **Flex Web Service Status** checkbox and **Save**. The status becomes ACTIVE.
10. In **Should Expire After**, pick the **longest** available option. The default
    is 6 hours, which will break any scheduled sync.
11. Leave **Valid For IP Address** blank — Google's servers don't have a fixed IP.
12. Click **Generate New Token** and copy it immediately.

**C. Connect the sheet**

13. **WealthScript → 🔗 Connect IBKR (optional)** and paste, in order: the token,
    the Query ID, and the Account Name exactly as it appears in column A of your
    Brokerage Holdings tab (e.g. `IBKR`).
14. Reload the sheet. The menu now shows **🔗 Sync IBKR Positions**.
15. Run it once manually and read the report before scheduling anything.

#### Regenerating the token

Tokens expire, and generating a new one **invalidates the current one** — so do
these two steps together or the sync will start failing with error `1012`.

1. **Performance & Reports → Flex Queries → Flex Web Service Configuration.**
2. Set **Should Expire After** to the longest option, leave the IP field blank,
   and click **Generate New Token**. Copy it.
3. In the sheet: **WealthScript → ⚙️ Reconfigure IBKR Flex** and paste the new
   token. The Query ID and Account Name are unchanged, but re-enter them.

Regenerate immediately if a token has ever appeared in a chat log, screenshot,
commit, or support ticket. A Flex token grants read access to your full
statement.

#### Flex error codes you may hit

| Code | Meaning | Fix |
|------|---------|-----|
| `1012` | Token has expired | Generate a new token, re-run the wizard |
| `1018` | Too many requests | Rate limited — wait and retry |
| `1019` | Statement still generating | Handled automatically by retry |
| `1020` | Invalid request | Stale reference code, wrong Query ID, or a token that was superseded |

Note that `q=` means different things in the two endpoints: the **Query ID** for
`SendRequest`, the **Reference Code** it returns for `GetStatement`. Reference
codes are single-use and short-lived.

Adding another broker means implementing two pure functions
 (a response parser
and a position normaliser) and reusing `_planHoldingsSync()`.

### 🔒 Where credentials are stored

| Secret | Location | Visible to |
|--------|----------|-----------|
| IBKR Flex token / Query ID | `DocumentProperties` | Editors who open Apps Script |
| GitHub PAT | `DocumentProperties` | Editors who open Apps Script |
| RapidAPI key | `DocumentProperties` | Editors who open Apps Script |

`DocumentProperties` is key/value storage attached to the bound Apps Script
project rather than to sheet content. It does not appear in the grid, in a
CSV/XLSX export, or in the Gist and Drive backups.

It is **not** a secret vault: anyone with edit access can open the script editor
and print it. It protects against view-only collaborators and accidental
leakage through exports, not against your own editors.

**Migrating an existing workbook.** Earlier versions kept the GitHub PAT in
`Settings & Config!B13` and the RapidAPI key in `B9`, where any collaborator
could read them. Run **WealthScript → 🔒 Secure Stored Credentials** once: it
copies each value into secure storage and replaces the cell with a
`🔒 Stored securely` notice. Idempotent, so a second run does nothing.

If a cell and secure storage hold *different* values, neither is touched and
both are reported — the tool will not guess which is current, since restoring a
revoked credential is worse than doing nothing. Clear the stale one and re-run.

Reads fall back to the legacy cell when secure storage is empty, so an
un-migrated workbook keeps working. Nothing breaks if you never run it; you just
stay exposed.

**Rotate after migrating.** Moving a secret does not un-share it. If the
workbook has ever been shared, or a value has appeared in a screenshot or an
export, regenerate it at the provider — GitHub for the PAT, RapidAPI for the
key, Client Portal for the Flex token.

### 🩺 Formula Integrity & Recovery

- **Damage detection** — `captureSnapshot()` refuses to run when managed formulas have been
  destroyed, so a stale net worth is never frozen into your permanent history. This is the
  safety net: the real cost of a broken formula isn't the broken cell, it's the months of
  corrupted snapshots taken over it.
- **Non-destructive repair** — `repairFormulas()` restores every managed formula in a live
  sheet without clearing anything, and never overwrites a deliberately pinned literal
  (option contract prices, manually typed balances).
- **Idempotent migration** — `migrateSheetLayout()` applies layout changes from a new
  release to an existing workbook. Safe to run repeatedly.
- **Guarded destructive ops** — builders that call `sheet.clear()` are hidden from the menu
  once the workbook holds real data, and live behind a Danger Zone submenu requiring a typed
  `ERASE` confirmation.

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
| FIRE & Cash Flow | B22 | **FIRE Target Net Worth (USD)** | `$10,000,000` |
| Dashboard Currency | B24 | Secondary Currency (Card 2) | `CAD` |
| Dashboard Currency | B25 | Secondary Currency (Card 3) | `INR` |
| ZPID Mapping | A29:B45 | Property Name → ZPID pairs | sample data |

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
  fireTargetUSD: 10000000,
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
| `test_buildAbbrDisplayFormula` | Locale-aware KPI display formula (crore/lakh vs. M/K) |
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
| *"See, edit, create and delete files in Google Drive"* | `drive` | Creates the backup folder and dated JSON files |
| *"Connect to an external service"* | `script.external_request` | GOOGLEFINANCE, GitHub Gist API, RapidAPI (Zillow) |
| *"Allow this app to run when you are not present"* | `script.scriptapp` | Weekly cron trigger for Real Estate price updates |
| *"Display third-party web content in sidebars"* | `script.container.ui` | Setup wizard dialog (GitHub + Drive onboarding) |

> **On the "unverified app" warning:** WealthScript is an open-source personal tool, not a submitted Google Workspace Add-on. Google shows this warning for any self-deployed script that isn't registered through their OAuth verification program. You are running your own code in your own account — clicking "Advanced → Go to project (unsafe)" is safe and expected.
