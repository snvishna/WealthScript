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
  const qty = Number(attrs.position);
  const mark = Number(attrs.markPrice);
  if (!isFinite(qty) || qty === 0) return null;

  const isOption = (cat === "OPT" || cat === "FOP");
  const multiplier = Number(attrs.multiplier) || (isOption ? 100 : 1);

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
 * Menu action: refreshes this account's rows on the Brokerage Holdings tab.
 *
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject]
 * @param {boolean} [silent=false]
 * @returns {{updated: number, added: number, zeroed: number, notes: Array<string>}}
 */
function syncIbkrPositions(ss_inject, silent = false) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Brokerage Holdings");
  const props = PropertiesService.getDocumentProperties();
  const accountName = props.getProperty(FLEX_PROP_ACCOUNT);

  if (!sheet) return { updated: 0, added: 0, zeroed: 0, notes: ["Brokerage Holdings tab not found."] };
  if (!accountName) return { updated: 0, added: 0, zeroed: 0, notes: ["IBKR Flex is not configured."] };

  const headerProblems = _auditHeaders(ss);
  if (headerProblems.length) {
    if (!silent) SpreadsheetApp.getUi().alert("🛑 Sync aborted — Holdings headers don't match the expected layout.");
    return { updated: 0, added: 0, zeroed: 0, notes: headerProblems };
  }

  const fetched = _fetchFlexStatement();
  if (!fetched.ok) {
    if (!silent) SpreadsheetApp.getUi().alert(`🛑 IBKR Flex request failed.\n\n${fetched.error}`);
    return { updated: 0, added: 0, zeroed: 0, notes: [fetched.error] };
  }

  const positions = [];
  _extractFlexElements(fetched.xml, "OpenPosition").forEach(a => {
    const p = _normaliseFlexPosition(a);
    if (p) positions.push(p);
  });
  _extractFlexElements(fetched.xml, "CashReportCurrency").forEach(a => {
    const c = _normaliseFlexCash(a.currency, a.endingCash || a.endingSettledCash);
    if (c) positions.push(c);
  });

  if (!positions.length) {
    if (!silent) SpreadsheetApp.getUi().alert(
      "⚠️ The Flex report contained no positions.\n\n" +
      "Check that your Flex Query includes the Open Positions and Cash Report\n" +
      "sections. Nothing was changed.");
    return { updated: 0, added: 0, zeroed: 0, notes: ["No positions in Flex report"] };
  }

  const rows = sheet.getRange(HOLDINGS_FIRST_ROW, 1, HOLDINGS_NUM_ROWS, 7).getValues();
  const plan = _planHoldingsSync(rows, positions, accountName, HOLDINGS_FIRST_ROW, HOLDINGS_LAST_ROW);

  const writeSpec = (row, spec, isNew) => {
    if (isNew) {
      sheet.getRange(row, 1).setValue(accountName);
      sheet.getRange(row, 3).setValue(spec.ticker);
      if (spec.currency) sheet.getRange(row, 7).setValue(spec.currency);
    }
    sheet.getRange(row, 4).setValue(spec.quantity);

    if (spec.priceFormula) {
      sheet.getRange(row, 5).setFormula(spec.priceFormula);
    } else if (spec.isOption && spec.price !== null) {
      // GOOGLEFINANCE can't price an OCC symbol, so the mark is written as a
      // literal. This is precisely the hand-typing the sync exists to remove.
      sheet.getRange(row, 5).setValue(spec.price);
    }
    // Non-option securities keep the canonical GOOGLEFINANCE formula in E, so
    // the sheet stays live between syncs rather than freezing at the mark.
  };

  plan.updates.forEach(u => writeSpec(u.row, u.spec, false));
  plan.additions.forEach(a => writeSpec(a.row, a.spec, true));
  plan.zeroed.forEach(z => sheet.getRange(z.row, 4).setValue(0));

  if (!silent) {
    const lines = [`✅ Synced ${accountName} from IBKR.`, "",
      `Positions updated: ${plan.updates.length}`,
      `New positions added: ${plan.additions.length}`,
      `Closed positions zeroed: ${plan.zeroed.length}`];
    if (plan.zeroed.length) {
      lines.push("", "Zeroed (rows kept so you can audit them):");
      plan.zeroed.slice(0, 15).forEach(z => lines.push(`  • ${z.ticker}`));
    }
    if (plan.notes.length) lines.push("", "Notes:", ...plan.notes.map(n => "  • " + n));
    SpreadsheetApp.getUi().alert(lines.join("\n"));
  }

  return { updated: plan.updates.length, added: plan.additions.length,
           zeroed: plan.zeroed.length, notes: plan.notes };
}
