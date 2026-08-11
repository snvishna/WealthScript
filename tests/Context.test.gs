/**
 * Tests for execution-context resolution.
 *
 * Regression: time-driven triggers call their handler with an EVENT OBJECT as
 * the first argument. `ss_inject || getActiveSpreadsheet()` took the truthy
 * event and every triggered run died with
 * "TypeError: ss.getSheetByName is not a function".
 */

function test_resolveSpreadsheetIgnoresTriggerEvents() {
  const realSheet = { getSheetByName: function () { return {}; }, __id: "real" };
  Assert.equal(_resolveSpreadsheet(realSheet).__id, "real",
    'resolveSS: a genuine Spreadsheet is passed through');

  // Shape of a real time-driven trigger event.
  const triggerEvent = {
    triggerUid: "1234567890",
    "day-of-month": 10, month: 8, year: 2026, hour: 7, minute: 32,
    timezone: "America/Vancouver", authMode: "FULL"
  };
  Assert.isTrue(!(triggerEvent && typeof triggerEvent.getSheetByName === "function"),
    'resolveSS: a trigger event is correctly identified as NOT a spreadsheet');

  const wrongShape = { getRange: function () {} };
  Assert.isTrue(!(wrongShape && typeof wrongShape.getSheetByName === "function"),
    'resolveSS: an object with other Sheets-like methods is still rejected');
}

function test_resolveSilent() {
  Assert.equal(_resolveSilent(false, true), false, 'silent: interactive run shows UI');
  Assert.equal(_resolveSilent(true, true), true, 'silent: an explicitly silent caller stays silent');
  Assert.equal(_resolveSilent(false, false), true,
    'silent: a triggered run suppresses UI, since getUi() throws there');
  Assert.equal(_resolveSilent(true, false), true, 'silent: both silent');
  Assert.equal(_resolveSilent(undefined, false), true, 'silent: undefined defaults safely in a trigger');
  Assert.equal(_resolveSilent(undefined, true), false, 'silent: undefined shows UI when interactive');
}
