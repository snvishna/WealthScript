# Context: WealthScript 2.0 Architecture

You are an expert Google Apps Script developer and financial modeling architect. I am building a professional-grade WealthScript in Google Sheets. 

I already have a robust, deployed `code.gs` backend that handles:
1.  **Multi-Currency Ledger:** Calculates Live USD net worth from CAD and INR assets using `GOOGLEFINANCE`.
2.  **Smart Taxes:** Deducts capital gains tax on taxable brokerages and income tax on pre-tax retirement accounts dynamically.
3.  **Automated Real Estate:** Uses a weekly Time-Driven trigger to fetch Zestimates from RapidAPI.
4.  **Disaster Recovery:** A manual and automated trigger that serializes the live ledger into JSON and patches it to a private GitHub Gist using a Personal Access Token.
5.  **Visual Dashboards:** Uses `SpreadsheetApp.newChart()` to generate a Donut Chart and Area Charts on a dedicated `📊 Insights & Analytics` tab, using a hidden `ChartData` tab as the backend.

### The Mission
I need to build out the final three "Epics" of this application. We must maintain the existing architectural pattern: all configuration (API keys, target numbers) lives in the `Settings & Config` tab, and all data processing must be highly efficient to prevent Google Apps Script execution timeouts.

Please provide the exact Apps Script functions and UI generation code to fulfill the following three Epics. We will integrate these into the existing `code.gs` file.

---

### Epic 1: The Expense & Burn Rate Pipeline
**Objective:** Net worth growth is only half the equation; we need to track monthly cash outflows to calculate the true FIRE (Financial Independence, Retire Early) timeline.

**Requirements:**
1.  Create a new function `buildCashFlowTab()` that generates a tab called `💸 Cash Flow & Burn`.
2.  This tab should have two sections: 
    * **Top:** A summary showing "Average Monthly Burn", "TTM (Trailing Twelve Months) Expenses", and "Current Safe Withdrawal Rate".
    * **Bottom:** A simple ledger with columns: `[Date, Category, Amount (USD), Notes]`.
3.  Update the `Settings & Config` tab generator to include a field for "Target Monthly FIRE Budget" (Default: $20,000).

---

### Epic 2: The FIRE Projections Engine (The $3M / $20k Goal)
**Objective:** I want to retire by age 50 with a $3,000,000 liquid net worth generating $20,000/month in cash flow. I need a predictive modeling tab.

**Requirements:**
1.  Create a function `buildFIREProjectionsTab()` that generates a tab called `🔥 FIRE Projections`.
2.  **Inputs:** It must dynamically pull my current "Liquid USD" from the latest snapshot, my "Current Monthly Burn" from the Cash Flow tab, and my real estate rental income (add a config for this in Settings).
3.  **The Math:** Project my portfolio forward by month.
    * Assume a default 7% annualized real return on liquid assets (configurable in Settings).
    * Assume my monthly surplus (Income - Burn) is added to the liquid principal.
    * Calculate the exact month/year I will cross the $3,000,000 liquid threshold.
4.  **Visualization:** Generate a line chart on this tab showing the "Projected Growth" line intersecting with the "Target $3M" horizontal line.

---

### Epic 3: True AI Agent Analysis
**Objective:** Currently, my `captureSnapshot()` function concatenates a hardcoded string (e.g., "Net worth increased by X. Liquid pool decreased by Y."). I want to replace this with an actual LLM analysis.

**Requirements:**
1.  Update the `Settings & Config` tab to accept an OpenAI API Key or Anthropic API Key.
2.  Write a function `generateAIInsight(currentJSON, previousJSON)` that takes the current month's ledger state and the previous month's ledger state.
3.  Construct a system prompt that tells the LLM to act as a strict financial analyst. It should analyze the deltas, identify which specific assets over/underperformed, flag liquidity risks, and output a concise, 3-sentence executive summary.
4.  Call the LLM API via `UrlFetchApp` and inject the resulting summary directly into the "Auto-Insights" column of the new row in the `Snapshots` tab. Ensure there is proper error handling (fallback to the hardcoded string if the API fails or times out).

**Instructions for your response:**
Please provide the code in logical blocks. Start with the updates required for the `Settings & Config` tab, then provide the code for Epic 1, Epic 2, and Epic 3. Ensure all code is strictly ES6 compliant and uses `SpreadsheetApp` best practices.
