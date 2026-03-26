const _currencySymbol = (code) => {
  const SYM = { USD:'$', EUR:'€', GBP:'£', INR:'₹', JPY:'¥',
                CAD:'CA$', AUD:'A$', SGD:'S$', CHF:'Fr', MXN:'MX$' };
  return SYM[code.toUpperCase()] || code;
};

const _abbrFmt = (code) => {
  const s = _currencySymbol(code);
  return `[>999999]"${s}"0.00,,"M";[>999]"${s}"0,"K";"${s}"0`;
};

const _generateInsight = (currentNet, prevNet, liquidUSD, prevLiquid) => {
  if (!prevNet || isNaN(prevNet)) return "Initial baseline snapshot established.";
  const dollarDelta = currentNet - prevNet;
  const liquidDelta = liquidUSD - prevLiquid;
  const formatVal = (val) => "$" + Math.abs(val).toLocaleString('en-US', { maximumFractionDigits: 0 });
  const trend = dollarDelta >= 0 ? "Increased" : "Decreased";
  const sign = (val) => val >= 0 ? "+" : "-";
  return `Net worth ${trend.toLowerCase()} by ${formatVal(dollarDelta)}. Liquid pool ${sign(liquidDelta)}${formatVal(liquidDelta)}.`;
};

function test_currencySymbol() {
  Assert.equal(_currencySymbol('USD'), '$', 'currSym: USD maps to $');
  Assert.equal(_currencySymbol('EUR'), '€', 'currSym: EUR maps to €');
  Assert.equal(_currencySymbol('GBP'), '£', 'currSym: GBP maps to £');
  Assert.equal(_currencySymbol('INR'), '₹', 'currSym: INR maps to ₹');
  Assert.equal(_currencySymbol('JPY'), '¥', 'currSym: JPY maps to ¥');
  Assert.equal(_currencySymbol('CAD'), 'CA$', 'currSym: CAD maps to CA$');
  Assert.equal(_currencySymbol('cad'), 'CA$', 'currSym: case-insensitive');
  Assert.equal(_currencySymbol('XYZ'), 'XYZ', 'currSym: unknown code returns the code itself');
}

function test_abbrFmt() {
  const usdFmt = _abbrFmt('USD');
  Assert.isTrue(usdFmt.includes('"$"'), 'abbrFmt: USD format includes $ symbol');
  Assert.isTrue(usdFmt.includes('"M"'), 'abbrFmt: USD format includes M suffix');
  Assert.isTrue(usdFmt.includes('"K"'), 'abbrFmt: USD format includes K suffix');

  const inrFmt = _abbrFmt('INR');
  Assert.isTrue(inrFmt.includes('"₹"'), 'abbrFmt: INR format includes ₹ symbol');

  const unknownFmt = _abbrFmt('BRL');
  Assert.isTrue(unknownFmt.includes('"BRL"'), 'abbrFmt: unknown code embeds the code string');
}

function test_generateInsight() {
  Assert.equal(
    _generateInsight(1000000, null, 200000, 0),
    'Initial baseline snapshot established.',
    'insight: first snapshot returns baseline message'
  );

  const up = _generateInsight(1100000, 1000000, 250000, 200000);
  Assert.isTrue(up.includes('increased'), 'insight: positive delta says "increased"');
  Assert.isTrue(up.includes('+'), 'insight: positive liquid delta shows + sign');

  const down = _generateInsight(900000, 1000000, 180000, 200000);
  Assert.isTrue(down.includes('decreased'), 'insight: negative delta says "decreased"');
  Assert.isTrue(down.includes('-'), 'insight: negative liquid delta shows - sign');

  const flat = _generateInsight(1000000, 1000000, 200000, 200000);
  Assert.isTrue(flat.includes('increased'), 'insight: zero delta treated as increase');
  Assert.isTrue(flat.includes('$0'), 'insight: zero delta formatted as $0');
}
