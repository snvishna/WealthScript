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
