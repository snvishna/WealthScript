# WealthScript 💰

A highly durable, automated, and professional-grade wealth tracker built directly into Google Sheets via Google Apps Script.

**Privacy-First & Offline-Ready:** Unlike Mint, Empower, or Monarch, WealthScript does not require handing over your bank credentials to third-party aggregators (like Plaid). Your data lives entirely in your private Google Drive. You manually input your balances—ensuring absolute privacy, security, and a hands-on relationship with your personal finances.

It features native multi-currency support (USD, CAD, INR), automated stock/crypto pricing, live real estate API ingestion, programmatic charting, a reverse-chronological snapshot engine, and **automated Cloud Disaster Recovery** via GitHub Gists.

---

## ✨ Core Features

* **Multi-Currency Engine:** Seamlessly unifies USD, CAD, and INR assets into a single master net worth, updating exchange rates live.
* **Smart Tax Calculation:** Automatically separates Capital Gains (taxed only on profit) from Pre-Tax Retirement accounts (taxed on the full balance) based on your provided cost basis.
* **Live Equity Sync:** Bypasses third-party APIs by using Google's native `GOOGLEFINANCE` engine for bulletproof stock, ETF, and crypto pricing.
* **Automated Real Estate Pricing:** Integrates with RapidAPI to pull live Zestimates for your properties via a weekly cron job.
* **Professional Dashboard:** Dark KPI banner with live Net / Gross Worth in USD and two user-configurable secondary currencies (e.g., CAD, INR) driven directly from the Settings tab. Liquid / Locked / FIRE Progress quick-stats generated from code with color-coded asset classes and conditional formatting.
* **The Snapshot Engine:** A single click instantly logs your current metrics, parses Liquid vs. Locked assets, and auto-generates a human-readable commentary on what moved your portfolio.
* **Cash Flow & Burn Rate Pipeline:** A dedicated `💸 Cash Flow & Burn` tab tracks monthly expenses with live KPIs: Average Burn, TTM Expenses, Target FIRE Budget, and Safe Withdrawal Rate.
* **Dual Disaster Recovery:** Every snapshot silently fires two backups in parallel — a JSON patch to a private **GitHub Gist** (developer-friendly diff history) AND a dated JSON file to **Google Drive** (zero-config, no API key needed).
* **Native Google Tables Support:** Designed explicitly to utilize Google Sheets' new "Convert to Table" feature for native asset class grouping.


---

## 🚀 Quick Start Guide

### Phase 1: Deploy the Engine
1. Create a brand new, blank [Google Sheet](https://sheets.new).
2. From the top menu, select **Extensions > Apps Script**.
3. Delete any default code in the editor, and paste the entirety of the `code.gs` file from this repository.
4. *(Optional but Recommended)*: In `code.gs`, paste your GitHub Personal Access Token into the `CLOUD_SYNC_CONFIG` block at the top so the script can auto-create your backup Gist. 
5. Click the **Save** icon (floppy disk).
6. Close the Apps Script tab and return to your Google Sheet. **Refresh the page.**

### Phase 2: First Time Setup
1. Look at your top menu bar (File, Edit, View, etc.). You will see a new menu item called **WealthScript**.
2. Click it, and select **"🚀 Run First Time Setup"**.
3. *Note: Google will present an "Authorization Required" popup. Click Continue > Choose your account > Advanced > "Go to Untitled Project (unsafe)" > Allow. (This is standard for custom scripts).*
4. The script will automatically generate your foundational tabs, establish the weekly cron job for real estate updates, and auto-provision your GitHub Gist.

### Phase 3: Configuration & Live Data
1. **The Settings Tab:** Go to the `Settings & Config` tab. Paste your API credentials and map your real estate ZPIDs. Review the new **FIRE & Cash Flow Config** section and adjust your target budget and return rate.
2. **The Holdings Tab:** Go to the `Brokerage Holdings` tab. List your stock/crypto tickers and quantities. The script will fetch their live prices.
3. **The Master Ledger:** Go to the `Dashboard & Ledger` tab. Enter your actual bank accounts, balances, and mortgage debt.
    * *Tip: For credit cards or loans, enter the Current Value as a **negative number**.*
4. **The Cash Flow Tab:** Go to the `💸 Cash Flow & Burn` tab. Replace the sample expense rows with your actual monthly outlays. The KPI summary at the top will update automatically.

#### Settings & Config — Full Reference

| Section | Row | Field | Default |
|---------|-----|-------|---------|
| Real Estate API | B2 | RapidAPI Key | `PASTE_KEY_HERE` |
| Real Estate API | B3 | RapidAPI Host | `real-estate101.p.rapidapi.com` |
| Cloud Backup | B6 | GitHub PAT (gist scope) | — |
| Cloud Backup | B7 | GitHub Gist ID | — |
| FIRE & Cash Flow | B10 | Target Monthly FIRE Budget (USD) | `$20,000` |
| FIRE & Cash Flow | B11 | Estimated Monthly Rental Income (USD) | `$0` |
| FIRE & Cash Flow | B12 | Annual Portfolio Return Rate | `7.00%` |
| Dashboard Currency | B14 | Secondary Currency (Card 2) | `CAD` |
| Dashboard Currency | B15 | Secondary Currency (Card 3) | `INR` |
| ZPID Mapping | A19:B35 | Property Name → ZPID pairs | sample data |

### Phase 4: Enable Native Tables & Grouping (Crucial Step)
Google Sheets' "Tables" feature is brand new and cannot be fully automated via script yet. You must enable it manually to get the visual asset grouping:
1. On your `Dashboard & Ledger`, highlight all your data (from Row 7 down to the bottom of your accounts).
2. Click **Format > Convert to Table**.
3. A "Table" dropdown will appear in the top-left of your new table. Click it, select **Table formatting**, and enable *Show table gridlines*, *Show alternating colours*, and *Show condensed view*.
4. **Group by Asset Class:** Click that same Table dropdown, select **Create group by view**, and choose **Asset Class**. 
5. In your new Grouped View, right-click the column headers for "Gross Worth", "Net Worth", and "Current Value" and set their calculation to **Sum**. 

---

## ☁️ Disaster Recovery Setup (GitHub Gist)

To ensure you never lose your financial history if a spreadsheet gets corrupted, this tracker pushes version-controlled JSON to GitHub.

**The Automated Way (Recommended):**
1. Go to GitHub > Settings > Developer Settings > Personal Access Tokens (Classic). 
2. Generate a new token and check **ONLY the `gist` scope**.
3. Open your `code.gs` file and paste the token into `CLOUD_SYNC_CONFIG.githubPAT` *before* running First Time Setup. The script will automatically create a secret Gist for you.

**The Manual Way:**
1. Generate the same PAT token as above.
2. Go to [gist.github.com](https://gist.github.com/) and create a new **Secret Gist** named `net_worth_backup.json` with empty brackets `{}`. 
3. Look at the URL of your new gist: `gist.github.com/yourusername/87654321abcd`. The long string at the end (`87654321abcd`) is your **Gist ID**.
4. Paste both your Token and your Gist ID into the `Settings & Config` tab manually.

---

## 🏠 Setting up Automated Real Estate (Zestimates)

1. Create a free account at [RapidAPI](https://rapidapi.com).
2. Search for `Zillow.com API` (or similar, like `real-estate101`). Subscribe to their free "Basic" tier (usually gives 20-50 free calls/month).
3. Copy your `X-RapidAPI-Key`. Paste it into the `Settings & Config` tab.
4. **Find your ZPID:** Go to Zillow, search for your property, and look at the URL. It will look like: `zillow.com/homedetails/123-Main-St/87654321_zpid/`. Your ZPID is the number (`87654321`). 
5. Add your Account Name (must match the ledger exactly!) and the ZPID to the mapping table in Settings.

---

## 📸 How to use the Snapshot Engine

Whenever you finish manually reconciling your bank accounts at the end of the month, simply go to the **WealthScript** menu at the top of your sheet and click **"📸 Log Snapshot & Cloud Sync"**.

**This single click does four things:**
1. Calculates your new totals, figures out your Liquid/Locked ratios, and records the exact currency exchange rates.
2. Injects a new row at the top of your `Snapshots` tab with an AI-generated summary of what moved your portfolio.
3. Automatically repaints the visual charts on your master dashboard to reflect the new data.
4. Silently pushes your new, current ledger state directly to your private GitHub Gist for disaster recovery. 

---

## 🛠 Developer Notes

### Customizing Default Accounts
If you want to customize the default dummy accounts that populate when you click "First Time Setup", open `code.gs` and modify the `DEFAULT_PORTFOLIO_DATA` array at the very top of the file before running the script.

### Test Suite (`tests.gs`)
The repository includes an isolated unit test suite in `tests.gs`. Tests cover pure business logic (delta calculation, FIRE progress ratio, asset liquidity classification) with **no live Sheets or network calls**.

**To run the tests:**
1. Open your Apps Script project (Extensions → Apps Script).
2. In the file selector on the left, open `tests.gs`.
3. Select the `runAllTests` function from the function dropdown.
4. Click ▶ **Run**.
5. Open **View → Logs** to see the pass/fail results.

### Project File Structure
| File | Purpose |
|------|---------|
| `code.gs` | Main application logic — tab builders, snapshot engine, API integrations |
| `tests.gs` | Isolated unit tests for pure business logic functions |
| `README.md` | This file |
| `docs/specs/` | Feature specs and architecture documents |

