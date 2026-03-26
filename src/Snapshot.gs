/**
 * Execute a manual snapshot. Calculates deltas, generates insights,
 * then chains cloud backups with transparent status reporting.
 */
function captureSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const logSheet = ss.getSheetByName("Snapshots");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!mainSheet || !logSheet) return;

  const netUSD = mainSheet.getRange("B2").getValue();
  const grossUSD = mainSheet.getRange("B3").getValue();
  const netCAD = mainSheet.getRange("E2").getValue();
  const netINR = mainSheet.getRange("H2").getValue();
  const dataRange = mainSheet.getRange("A7:J80").getValues(); 
  
  let liquidUSD = 0, lockedUSD = 0, totalReUSD = 0;

  for (let i = 0; i < dataRange.length; i++) {
    let assetClass = String(dataRange[i][1]);
    let netVal = Number(dataRange[i][8]); 
    let status = String(dataRange[i][9]); 

    if (status === "Active" && !isNaN(netVal) && netVal !== 0) {
      if (["Cash", "Brokerage", "Crypto", "Receivable"].includes(assetClass)) {
        liquidUSD += netVal;
      } else {
        lockedUSD += netVal; 
      }
      if (assetClass === "Real Estate") totalReUSD += netVal;
    }
  }

  const prevNetUSD = logSheet.getRange(2, 2).getValue(); 
  const prevLiquid = logSheet.getRange(2, 3).getValue();
  let dollarDelta = "", pctGrowth = "", autoInsight = "Initial baseline snapshot established.", fireProgress = netUSD / 3000000; 

  if (prevNetUSD && !isNaN(prevNetUSD)) {
    dollarDelta = netUSD - prevNetUSD;
    pctGrowth = dollarDelta / prevNetUSD;
    let liquidDelta = liquidUSD - prevLiquid;
    
    let formatVal = (val) => "$" + Math.abs(val).toLocaleString('en-US', {maximumFractionDigits:0});
    let trend = dollarDelta >= 0 ? "Increased" : "Decreased";
    let sign = (val) => val >= 0 ? "+" : "-";

    autoInsight = `Net worth ${trend.toLowerCase()} by ${formatVal(dollarDelta)}. Liquid pool ${sign(liquidDelta)}${formatVal(liquidDelta)}.`;
  }

  logSheet.insertRowBefore(2);
  const rowData = [new Date(), netUSD, liquidUSD, lockedUSD, grossUSD, netCAD, netINR, totalReUSD, dollarDelta, pctGrowth, fireProgress, autoInsight, ""];
  logSheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);

  logSheet.getRange(2, 2, 1, 7).setNumberFormat("$#,##0.00"); 
  logSheet.getRange(2, 9).setNumberFormat("[Color10]+$#,##0.00;[Color3]-$#,##0.00"); 
  logSheet.getRange(2, 10).setNumberFormat("[Color10]+0.00%;[Color3]-0.00%"); 
  logSheet.getRange(2, 11).setNumberFormat("0.00%"); 
  
  // --- Cloud Backup Chain with Transparent Status ---
  const gistConfigured = _isGistConfigured(configSheet);
  const gistOk    = gistConfigured ? backupToGitHub(true) : false;
  const driveOk   = backupToGoogleDrive(true);

  // Build transparent status message
  const statusParts = ["✅ Snapshot captured successfully!"];
  
  if (gistConfigured && gistOk) {
    statusParts.push("☁️ GitHub Gist — Synced");
  } else if (gistConfigured && !gistOk) {
    statusParts.push("⚠️ GitHub Gist — Sync failed (check logs)");
  } else {
    statusParts.push("💤 GitHub Gist — Not configured");
  }

  if (driveOk) {
    statusParts.push("📁 Google Drive — Synced");
  } else {
    statusParts.push("💤 Google Drive — Not synced");
  }

  if (!gistConfigured && !driveOk) {
    statusParts.push("\n💡 Tip: Set up cloud backups from the WealthScript menu:\n• 🔐 Setup GitHub Backup\n• 📁 Setup Google Drive Backup");
  }

  SpreadsheetApp.getUi().alert(statusParts.join("\n"));

  updateVisualDashboards(); 
}
