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
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
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
