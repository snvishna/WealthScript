/**
 * Fetches Zestimates using config from the Settings tab.
 */
function updateRealEstatePrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");
  
  if(!configSheet) return SpreadsheetApp.getUi().alert("Settings tab missing. Run First Time Setup.");

  const apiKey = configSheet.getRange("B2").getValue();
  const apiHost = configSheet.getRange("B3").getValue();
  
  if (!apiKey || apiKey === "PASTE_KEY_HERE") return; 

  const propData = configSheet.getRange("A21:B37").getValues();
  const properties = [];
  for (let i = 0; i < propData.length; i++) {
    if (propData[i][0] && propData[i][1]) {
      properties.push({ name: String(propData[i][0]), zpid: String(propData[i][1]) });
    }
  }

  const accountNames = sheet.getRange("A1:A100").getValues().flat();
  const currentValues = sheet.getRange("E1:E100").getValues();
  let updated = false;

  properties.forEach(prop => {
    try {
      const url = `https://${apiHost}/api/property-details/byzpid?zpid=${prop.zpid}`;
      const options = {
        method: 'GET',
        headers: {
          'X-RapidAPI-Key': apiKey,
          'X-RapidAPI-Host': apiHost
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const zestimate = data.property.zestimate; 
        const targetRowIdx = accountNames.indexOf(prop.name); 
        
        if (targetRowIdx >= 0 && zestimate) {
          currentValues[targetRowIdx][0] = zestimate;
          updated = true;
        }
      }
    } catch (e) {
      Logger.log(`Script crashed on ${prop.name}: ${e}`);
    }
  });

  if (updated) {
    sheet.getRange("E1:E100").setValues(currentValues);
  }
}
