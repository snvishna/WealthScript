/**
 * Generates professional, full-screen visual dashboards on a dedicated tab.
 */
function updateVisualDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Dashboard & Ledger");
  const snapSheet = ss.getSheetByName("Snapshots");

  if (!ledgerSheet || !snapSheet) return;

  // 1. Setup the Dedicated UI Tab
  let dashboardName = "📊 Insights & Analytics";
  let uiSheet = ss.getSheetByName(dashboardName);
  
  if (!uiSheet) {
    uiSheet = ss.insertSheet(dashboardName, 1); 
  } else {
    const existingCharts = uiSheet.getCharts();
    for (let i = 0; i < existingCharts.length; i++) {
      uiSheet.removeChart(existingCharts[i]);
    }
    uiSheet.clear(); 
  }

  // Format the Dashboard Canvas
  uiSheet.getRange("A1:Z100").setBackground(THEME.canvas); 
  uiSheet.getRange("B2").setValue("PORTFOLIO ANALYTICS & TRENDS").setFontWeight("bold").setFontSize(16).setFontColor(THEME.assetText);
  
  uiSheet.setHiddenGridlines(true);

  // 2. Setup the Hidden Data Backend
  let dataSheet = ss.getSheetByName("ChartData");
  if (!dataSheet) {
    dataSheet = ss.insertSheet("ChartData");
    dataSheet.hideSheet(); 
  } else {
    dataSheet.clear();
  }

  // --- CHART BUILDERS ---

  // A. Asset Allocation (Modern Minimalist Donut)
  const lastAssetRow = ledgerSheet.getLastRow();
  if (lastAssetRow > 6) {
    const pieChart = uiSheet.newChart()
      .asPieChart()
      .addRange(ledgerSheet.getRange(7, 2, lastAssetRow - 6, 1)) // Asset Class
      .addRange(ledgerSheet.getRange(7, 9, lastAssetRow - 6, 1)) // Net Worth (USD)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
      .setOption('title', 'Asset Allocation Framework')
      .setOption('pieHole', 0.55)
      .setOption('colors', THEME.charts.donut)
      .setOption('pieSliceBorderColor', "transparent")
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '80%'})
      .setOption('legend', {position: 'right', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('pieSliceText', 'none')
      .setPosition(4, 2, 0, 0) // Row 4, Col B
      .build();

    uiSheet.insertChart(pieChart);
  }

  // B. Historical Net Worth (Smooth Area Chart)
  const lastSnapRow = snapSheet.getLastRow();
  if (lastSnapRow > 1) {
    const areaChart = uiSheet.newChart()
      .asAreaChart()
      .addRange(snapSheet.getRange(1, 1, lastSnapRow, 1)) // X: Dates
      .addRange(snapSheet.getRange(1, 2, lastSnapRow, 1)) // Y1: Net USD
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Net Worth Trajectory (USD)')
      .setOption('colors', THEME.charts.area) 
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('curveType', 'function')
      .setOption('pointSize', 6)
      .setOption('lineWidth', 3)
      .setOption('chartArea', {left: '15%', top: '15%', width: '80%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: '$#,###' })
      .setOption('hAxis', { textStyle: {color: THEME.charts.axisText}, format: 'MMM yyyy' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('trendlines', {0: {type: 'polynomial', degree: 2, color: THEME.charts.legendText, opacity: 0.3, visibleInLegend: true, labelInLegend: 'Trend'}})
      .setOption('series', {
        0: {label: 'Total Net Worth'}
      })
      .setPosition(4, 7, 0, 0) // Row 4, Col G (Next to Donut)
      .build();

    uiSheet.insertChart(areaChart);

    // C. Liquid vs Locked (Stacked Column Chart)
    const stackedBar = uiSheet.newChart()
      .asColumnChart()
      .addRange(snapSheet.getRange(1, 1, lastSnapRow, 1)) // X: Dates
      .addRange(snapSheet.getRange(1, 3, lastSnapRow, 1)) // Y1: Liquid
      .addRange(snapSheet.getRange(1, 4, lastSnapRow, 1)) // Y2: Locked
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Liquidity Profile: Liquid vs. Locked Assets')
      .setOption('isStacked', true)
      .setOption('colors', THEME.charts.stacked) 
      .setOption('dataOpacity', 0.85)
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('chartArea', {left: '10%', top: '15%', width: '85%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: '$#,###' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setOption('series', {
        0: {label: 'Liquid Assets'},
        1: {label: 'Locked Assets'}
      })
      .setPosition(22, 2, 0, 0) // Row 22, Col B (Below the others)
      .build();

    uiSheet.insertChart(stackedBar);
  }

  // D. Portfolio X-Ray (Donut Chart)
  const holdingsSheet = ss.getSheetByName("Brokerage Holdings");
  if (holdingsSheet) {
    const lastXrayRow = holdingsSheet.getLastRow();
    if (lastXrayRow > 1) {
      const xrayChart = uiSheet.newChart()
        .asPieChart()
        .addRange(holdingsSheet.getRange(2, 2, lastXrayRow - 1, 1)) // Asset Category
        .addRange(holdingsSheet.getRange(2, 6, lastXrayRow - 1, 1)) // Total Value
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
        .setOption('title', 'Portfolio Exposure X-Ray')
        .setOption('pieHole', 0.55)
        .setOption('colors', THEME.charts.donut)
        .setOption('pieSliceBorderColor', "transparent")
        .setOption('backgroundColor', { fill: 'transparent' })
        .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '80%'})
        .setOption('legend', {position: 'right', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
        .setOption('pieSliceText', 'none')
        .setPosition(22, 7, 0, 0) // Row 22, Col G (Next to Liquid vs Locked)
        .build();

      uiSheet.insertChart(xrayChart);
    }
  }
}
