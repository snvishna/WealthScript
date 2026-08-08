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
