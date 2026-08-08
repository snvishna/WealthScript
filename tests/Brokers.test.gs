/**
 * Tests for the broker sync layer. All pure — no UrlFetchApp, no Sheets.
 */

function test_buildOccSymbol() {
  Assert.equal(_buildOccSymbol("MSFT", "20260918", "C", 545), "MSFT260918C00545000",
    'occ: matches the symbol already in the sheet');
  Assert.equal(_buildOccSymbol("SPCX", "20260904", "P", 122), "SPCX260904P00122000",
    'occ: puts encode correctly');
  Assert.equal(_buildOccSymbol("aapl", "20261218", "call", 292.5), "AAPL261218C00292500",
    'occ: fractional strikes and lowercase roots normalise');
}

function test_normaliseFlexPosition() {
  const opt = _normaliseFlexPosition({
    assetCategory: "OPT", underlyingSymbol: "MSFT", expiry: "20260918",
    putCall: "C", strike: "545", position: "-2", markPrice: "5.15944145", multiplier: "100"
  });
  Assert.equal(opt.ticker, "MSFT260918C00545000", 'flexPos: option ticker built from components');
  Assert.equal(opt.quantity, -2, 'flexPos: short position stays negative');
  Assert.isTrue(Math.abs(opt.price - 515.944145) < 0.0001, 'flexPos: option price scaled to per-contract');
  Assert.isTrue(opt.isOption, 'flexPos: flagged as option so a literal mark is written');

  const stk = _normaliseFlexPosition({
    assetCategory: "STK", symbol: "AMZN", position: "1100", markPrice: "274.48"
  });
  Assert.equal(stk.ticker, "AMZN", 'flexPos: equity uses the plain symbol');
  Assert.equal(stk.price, 274.48, 'flexPos: equity price not multiplied');
  Assert.isTrue(!stk.isOption, 'flexPos: equity keeps its live GOOGLEFINANCE formula');

  Assert.equal(_normaliseFlexPosition({ assetCategory: "STK", symbol: "X", position: "0" }), null,
    'flexPos: zero-quantity rows are dropped');
  Assert.equal(_normaliseFlexPosition(null), null, 'flexPos: null input is safe');
}

function test_normaliseFlexCash() {
  const usd = _normaliseFlexCash("USD", "11128.41");
  Assert.equal(usd.priceFormula, "=1", 'flexCash: USD priced at 1, never sent to GOOGLEFINANCE');
  Assert.equal(usd.ticker, "Cash", 'flexCash: uses the reserved Cash ticker');

  const cad = _normaliseFlexCash("CAD", "110000");
  Assert.isTrue(cad.priceFormula.indexOf("CURRENCY:CADUSD") > -1,
    'flexCash: foreign cash gets a live FX formula, not a frozen rate');
  Assert.equal(cad.currency, "CAD", 'flexCash: currency retained for row matching');

  Assert.equal(_normaliseFlexCash("USD", "0"), null, 'flexCash: zero balances dropped');
  Assert.equal(_normaliseFlexCash("BASE_SUMMARY", "500"), null, 'flexCash: summary pseudo-currency ignored');
}

function test_parseFlexEnvelope() {
  const ok = _parseFlexEnvelope('<FlexStatementResponse><Status>Success</Status><ReferenceCode>1234567890</ReferenceCode></FlexStatementResponse>');
  Assert.isTrue(ok.ok, 'envelope: success detected');
  Assert.equal(ok.referenceCode, "1234567890", 'envelope: reference code extracted');

  const bad = _parseFlexEnvelope('<FlexStatementResponse><Status>Fail</Status><ErrorCode>1012</ErrorCode><ErrorMessage>Token has expired.</ErrorMessage></FlexStatementResponse>');
  Assert.isTrue(!bad.ok, 'envelope: failure detected');
  Assert.isTrue(bad.error.indexOf("1012") > -1 && bad.error.indexOf("expired") > -1,
    'envelope: code and message surfaced to the user');
}

function test_extractFlexElements() {
  const xml = '<FlexQueryResponse><OpenPositions>' +
    '<OpenPosition accountId="U123" symbol="AMZN" position="1100" markPrice="274.48" assetCategory="STK" />' +
    '<OpenPosition accountId="U123" symbol="MSFT" position="290" markPrice="499.99" assetCategory="STK" />' +
    '</OpenPositions><CashReport>' +
    '<CashReportCurrency currency="CAD" endingCash="110000" />' +
    '</CashReport></FlexQueryResponse>';

  const pos = _extractFlexElements(xml, "OpenPosition");
  Assert.equal(pos.length, 2, 'extract: finds every OpenPosition');
  Assert.equal(pos[0].symbol, "AMZN", 'extract: attributes parsed');

  const cash = _extractFlexElements(xml, "CashReportCurrency");
  Assert.equal(cash.length, 1, 'extract: cash section parsed separately');
  Assert.equal(cash[0].currency, "CAD", 'extract: currency attribute read');

  Assert.equal(_extractFlexElements("", "OpenPosition").length, 0, 'extract: empty input is safe');
}

function test_planHoldingsSync() {
  //             A account   B cat    C ticker   D qty  E  F  G currency
  const mk = (a, t, q, cur) => [a, "", t, q, "", "", cur || ""];
  const rows = [
    mk("IBKR", "AMZN", 1608),                 // row 2  - held, quantity changed
    mk("IBKR", "Cash", 11128.41, "USD"),      // row 3  - USD cash
    mk("IBKR", "Cash", 110000, "CAD"),        // row 4  - CAD cash, same ticker
    mk("Robinhood", "GME", 10),               // row 5  - another account
    mk("IBKR", "DOCU", 40),                   // row 6  - position closed
    ["", "", "", "", "", "", ""]              // row 7  - free
  ];
  const positions = [
    { ticker: "AMZN", quantity: 1100, price: 274.48, isOption: false },
    { ticker: "Cash", quantity: 11128.41, priceFormula: "=1", currency: "USD" },
    { ticker: "Cash", quantity: 110000, priceFormula: "=X", currency: "CAD" },
    { ticker: "MSFT260918C00545000", quantity: -2, price: 515.94, isOption: true }
  ];

  const plan = _planHoldingsSync(rows, positions, "IBKR", 2, 7);

  Assert.equal(plan.updates.length, 3, 'plan: matched rows are updated in place');
  Assert.equal(plan.updates[0].row, 2, 'plan: AMZN matched to its existing row');
  Assert.isTrue(plan.updates.some(u => u.row === 3) && plan.updates.some(u => u.row === 4),
    'plan: two Cash rows disambiguated by currency, not merged');

  Assert.equal(plan.additions.length, 1, 'plan: the new option is added');
  Assert.equal(plan.additions[0].row, 7, 'plan: added into the free row');

  Assert.equal(plan.zeroed.length, 1, 'plan: the closed position is zeroed');
  Assert.equal(plan.zeroed[0].ticker, "DOCU", 'plan: names the closed position');

  Assert.isTrue(!plan.updates.concat(plan.additions).some(x => x.row === 5),
    'plan: another account\'s row is never touched');
}

function test_planHoldingsSyncRunsOutOfRoom() {
  const rows = [["IBKR", "", "AMZN", 100, "", "", ""]];
  const plan = _planHoldingsSync(rows, [
    { ticker: "AMZN", quantity: 100, isOption: false },
    { ticker: "TSLA", quantity: 50, isOption: false }
  ], "IBKR", 2, 2);

  Assert.equal(plan.additions.length, 0, 'room: nothing written when there is no free row');
  Assert.isTrue(plan.notes.join(" ").indexOf("TSLA") > -1, 'room: the skipped position is reported, not silently dropped');
}

function test_deriveQuantityFromPositionValue() {
  // Real Flex output omits the Position attribute unless the query includes it.
  const aapl = _normaliseFlexPosition({
    assetCategory: "STK", symbol: "AAPL", underlyingSymbol: "AAPL",
    multiplier: "1", markPrice: "313.33", positionValue: "12533.2"
  });
  Assert.equal(aapl.quantity, 40, 'derive: equity quantity recovered from positionValue');

  const amzn = _normaliseFlexPosition({
    assetCategory: "STK", symbol: "AMZN", multiplier: "1",
    markPrice: "274.48", positionValue: "301928"
  });
  Assert.equal(amzn.quantity, 1100, 'derive: large share count exact, no float noise');

  const opt = _normaliseFlexPosition({
    assetCategory: "OPT", symbol: "MSFT 260918C00545000", underlyingSymbol: "MSFT",
    multiplier: "100", strike: "545", expiry: "20260918", putCall: "C",
    markPrice: "5.1331", positionValue: "-1026.62"
  });
  Assert.equal(opt.quantity, -2, 'derive: short option contracts recovered and stay negative');
  Assert.equal(opt.ticker, "MSFT260918C00545000", 'derive: OCC symbol rebuilt without the Flex space');
  Assert.isTrue(Math.abs(opt.price - 513.31) < 0.01, 'derive: option price per contract');

  Assert.equal(_normaliseFlexPosition({
    assetCategory: "STK", symbol: "X", markPrice: "0", positionValue: "100"
  }), null, 'derive: zero mark cannot yield a quantity, so the row is refused not guessed');
}

function test_planBlockSyncRefusesEmptyFeed() {
  const rows = [["IBKR", "", "AAPL", 40, "", "", ""], ["IBKR", "", "AMZN", 1100, "", "", ""]];
  const plan = _planBlockSync(rows, [], "IBKR", 18, 100);
  Assert.isTrue(!plan.ok, 'block: an empty feed is refused');
  Assert.isTrue(plan.reason.indexOf("failed request") > -1, 'block: explains it is a failed feed, not an emptied account');
}

function test_planBlockSyncRewritesInPlace() {
  const rows = [
    ["Robinhood", "", "GME", 10, "", "", ""],   // 18 - other account, above
    ["IBKR", "", "AAPL", 40, "", "", ""],       // 19
    ["IBKR", "", "AMZN", 1100, "", "", ""],     // 20
    ["IBKR", "", "DOCU", 40, "", "", ""],       // 21 - no longer held
    ["", "", "", "", "", "", ""]                // 22 - blank
  ];
  const positions = [
    { ticker: "AAPL", quantity: 40, isOption: false },
    { ticker: "AMZN", quantity: 1100, isOption: false }
  ];
  const plan = _planBlockSync(rows, positions, "IBKR", 18, 100);

  Assert.isTrue(plan.ok, 'block: a valid rewrite is planned');
  Assert.equal(plan.startRow, 19, 'block: starts at the account\'s first row, not row 1');
  Assert.equal(plan.writes.length, 2, 'block: one write per position');
  Assert.equal(plan.clears.length, 1, 'block: the stale row is cleared');
  Assert.equal(plan.clears[0], 21, 'block: clears the right row');
  Assert.equal(plan.removed[0], "DOCU", 'block: names what is no longer held');
  Assert.isTrue(!plan.writes.some(w => w.row === 18), 'block: never writes over another account');
}

function test_planBlockSyncRefusesToStraddleAnotherAccount() {
  const rows = [
    ["IBKR", "", "AAPL", 40, "", "", ""],       // 18
    ["Robinhood", "", "GME", 10, "", "", ""],   // 19 - in the way
    ["IBKR", "", "AMZN", 1100, "", "", ""],     // 20
    ["", "", "", "", "", "", ""],               // 21
    ["", "", "", "", "", "", ""]                // 22
  ];
  const plan = _planBlockSync(rows, [
    { ticker: "AAPL", quantity: 40 }, { ticker: "AMZN", quantity: 1100 }, { ticker: "TSLA", quantity: 5 }
  ], "IBKR", 18, 22);

  // The block can't grow in place, so it relocates below rather than writing
  // over row 19. Either way, the foreign row is never touched.
  Assert.isTrue(plan.ok, 'straddle: relocates instead of refusing outright');
  Assert.isTrue(plan.relocated, 'straddle: flagged as a move');
  Assert.isTrue(!plan.writes.some(w => w.row === 19), 'straddle: never writes over a foreign row');
  Assert.equal(plan.startRow, 20, 'straddle: reuses the second owned row as the new anchor');
  Assert.equal(plan.clears.length, 1, 'straddle: only the genuinely vacated row is cleared');
  Assert.equal(plan.clears[0], 18, 'straddle: clears the row left behind, not the one reused');
}

function test_planBlockSyncFlagsLargeShrink() {
  const rows = [];
  for (let i = 0; i < 10; i++) rows.push(["IBKR", "", "T" + i, 1, "", "", ""]);
  const plan = _planBlockSync(rows, [{ ticker: "T0", quantity: 1 }], "IBKR", 18, 100);

  Assert.isTrue(plan.ok, 'shrink: planned, but flagged for confirmation');
  Assert.isTrue(plan.shrinkRatio > SYNC_MAX_SHRINK, 'shrink: a 90% drop exceeds the threshold');
  Assert.equal(plan.clears.length, 9, 'shrink: reports every row that would go');
}

function test_headerAnnotationTolerated() {
  const withSuffix = HOLDINGS_HEADERS.slice();
  withSuffix[5] = "Total Value";
  Assert.isTrue(_verifyHeaders(withSuffix, HOLDINGS_HEADERS).ok,
    'annotation: an unannotated header still matches the annotated canonical');

  const annotated = HOLDINGS_HEADERS.slice();
  annotated[3] = "Quantity (shares)";
  Assert.isTrue(_verifyHeaders(annotated, HOLDINGS_HEADERS).ok,
    'annotation: a user-added parenthetical is accepted');

  const wrong = LEDGER_HEADERS.slice();
  wrong[9] = "Net Worth (USD)";
  Assert.isTrue(!_verifyHeaders(wrong, LEDGER_HEADERS).ok,
    'annotation: a genuinely wrong header still fails');
}


function test_planBlockSyncBootstrapsEmptyAccount() {
  // No row anywhere carries "IBKR" — the first sync must still work.
  const rows = [
    ["Robinhood", "", "GME", 10, "", "", ""],   // 18
    ["", "", "", "", "", "", ""],               // 19
    ["", "", "", "", "", "", ""],               // 20
    ["", "", "", "", "", "", ""]                // 21
  ];
  const plan = _planBlockSync(rows, [
    { ticker: "AAPL", quantity: 40 }, { ticker: "AMZN", quantity: 1100 }
  ], "IBKR", 18, 21);

  Assert.isTrue(plan.ok, 'bootstrap: an account with no existing rows still syncs');
  Assert.isTrue(plan.bootstrap, 'bootstrap: flagged so the report says "Created" not "Rebuilt"');
  Assert.equal(plan.startRow, 19, 'bootstrap: lands in the first free run, after the other account');
  Assert.equal(plan.writes.length, 2, 'bootstrap: every position written');
  Assert.equal(plan.clears.length, 0, 'bootstrap: nothing to clear');
}

function test_planBlockSyncRelocatesWhenOutgrown() {
  // IBKR sits at row 18 with one free row after it, then another account.
  const rows = [
    ["IBKR", "", "AAPL", 40, "", "", ""],       // 18
    ["Robinhood", "", "GME", 10, "", "", ""],   // 19 - blocks growth
    ["", "", "", "", "", "", ""],               // 20
    ["", "", "", "", "", "", ""],               // 21
    ["", "", "", "", "", "", ""]                // 22
  ];
  const plan = _planBlockSync(rows, [
    { ticker: "AAPL", quantity: 40 }, { ticker: "AMZN", quantity: 1100 }, { ticker: "TSLA", quantity: 5 }
  ], "IBKR", 18, 22);

  Assert.isTrue(plan.ok, 'relocate: grows into free space instead of refusing');
  Assert.isTrue(plan.relocated, 'relocate: flagged so the user is told the block moved');
  Assert.equal(plan.startRow, 20, 'relocate: moved past the blocking account');
  Assert.equal(plan.clears.length, 1, 'relocate: the old anchor row is cleared');
  Assert.equal(plan.clears[0], 18, 'relocate: clears the vacated row');
  Assert.isTrue(!plan.writes.some(w => w.row === 19), 'relocate: never writes over the other account');
}

function test_planBlockSyncRefusesWhenNoRunFits() {
  const rows = [
    ["Robinhood", "", "GME", 10, "", "", ""],   // 18
    ["", "", "", "", "", "", ""],               // 19
    ["Coinbase", "", "BTC", 1, "", "", ""]      // 20
  ];
  const plan = _planBlockSync(rows, [
    { ticker: "AAPL", quantity: 40 }, { ticker: "AMZN", quantity: 1100 }
  ], "IBKR", 18, 20);

  Assert.isTrue(!plan.ok, 'noroom: refuses when no run is long enough');
  Assert.isTrue(plan.reason.indexOf("consecutive rows") > -1, 'noroom: explains what is needed');
}

function test_planBlockSyncShrinkNeverNegative() {
  const rows = [["IBKR", "", "AAPL", 40, "", "", ""], ["", "", "", "", "", "", ""], ["", "", "", "", "", "", ""]];
  const plan = _planBlockSync(rows, [
    { ticker: "AAPL", quantity: 40 }, { ticker: "AMZN", quantity: 1100 }, { ticker: "TSLA", quantity: 5 }
  ], "IBKR", 18, 20);
  Assert.equal(plan.shrinkRatio, 0, 'shrink: growing the block is never reported as shrinkage');
}
