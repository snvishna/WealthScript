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
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet || !sheet) return { updated: 0, skipped: ["Required tabs not found"] };

  const apiKey = configSheet.getRange("B9").getValue();
  const apiHost = configSheet.getRange("B10").getValue();

  if (!apiKey || apiKey === "PASTE_KEY_HERE") return { updated: 0, skipped: ["RapidAPI key not configured"] };

  const propData = configSheet.getRange("A28:B44").getValues();
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
