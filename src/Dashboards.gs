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

  // B. Time-to-FIRE Forecast (Combo Chart: Actuals + CAGR Projection)
  const lastSnapRow = snapSheet.getLastRow();
  if (lastSnapRow > 1) {
    const snapData = snapSheet.getRange(2, 1, lastSnapRow - 1, 2).getValues();
    snapData.reverse(); // Snapshots are push-to-top, reverse for chronological order

    const forecastData = [["Date", "Actual Net Worth", "Projected Timeline"]];
    const targetUSD = DASHBOARD_CONFIG.fireTargetUSD || 3000000;

    if (snapData.length > 1) {
      const firstRow = snapData[0];
      const lastRow = snapData[snapData.length - 1];
      const firstDate = new Date(firstRow[0]);
      const lastDate = new Date(lastRow[0]);
      const firstNet = Number(firstRow[1]);
      const lastNet = Number(lastRow[1]);

      for (let i = 0; i < snapData.length; i++) {
        forecastData.push([snapData[i][0], snapData[i][1], null]);
      }

      const yearsElapsed = (lastDate - firstDate) / (1000 * 60 * 60 * 24 * 365.25);
      
      // Compute CAGR and project if there is actual positive growth and haven't hit target yet
      if (yearsElapsed >= 0 && firstNet > 0 && lastNet > firstNet && lastNet < targetUSD) {
        // Prevent Divide by Zero parsing if yearsElapsed is mathematically 0 (same day snapshot)
        const safeYears = Math.max(yearsElapsed, 0.25); 
        const cagr = Math.pow(lastNet / firstNet, 1 / safeYears) - 1;

        if (cagr > 0) {
          // Join the actual line to the projected line visually
          forecastData[forecastData.length - 1][2] = lastNet;

          let projectedNet = lastNet;
          let projectedDate = new Date(lastDate);
          let safetyStops = 0;

          // Steps by quarter up to 30 years (120 quarters)
          while (projectedNet < targetUSD && safetyStops < 120) {
            projectedDate.setMonth(projectedDate.getMonth() + 3);
            projectedNet = projectedNet * Math.pow(1 + cagr, 0.25);

            if (projectedNet > targetUSD) projectedNet = targetUSD;
            forecastData.push([new Date(projectedDate), null, projectedNet]);
            safetyStops++;
          }
        }
      }
    } else {
      forecastData.push([snapData[0][0], snapData[0][1], null]);
    }

    dataSheet.getRange(1, 1, forecastData.length, 3).setValues(forecastData);

    const comboChart = uiSheet.newChart()
      .asComboChart()
      .addRange(dataSheet.getRange(1, 1, forecastData.length, 1)) // X: Dates
      .addRange(dataSheet.getRange(1, 2, forecastData.length, 1)) // Y1: Actuals
      .addRange(dataSheet.getRange(1, 3, forecastData.length, 1)) // Y2: Forecast
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Time-to-FIRE Trajectory (USD)')
      .setOption('backgroundColor', { fill: 'transparent' })
      .setOption('seriesType', 'area')
      .setOption('series', {
        0: {type: 'area', color: THEME.charts.area[0], pointSize: 6, lineWidth: 3, label: 'Actual Net Worth'},
        1: {type: 'line', color: THEME.charts.stacked[0], pointSize: 0, lineWidth: 3, lineDashStyle: [4, 4], label: 'Projected FIRE Timeline (CAGR)'}
      })
      .setOption('chartArea', {left: '15%', top: '15%', width: '80%', height: '70%'})
      .setOption('vAxis', { gridlines: {color: THEME.charts.gridlines}, textStyle: {color: THEME.charts.axisText}, format: '$#,###' })
      .setOption('hAxis', { textStyle: {color: THEME.charts.axisText}, format: 'MMM yyyy' })
      .setOption('legend', {position: 'top', alignment: 'end', textStyle: {fontSize: 12, color: THEME.charts.legendText}})
      .setPosition(4, 7, 0, 0) // Row 4, Col G
      .build();

    uiSheet.insertChart(comboChart);

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
