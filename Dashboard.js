/**
 * @file Manages the Dashboard and Job Data sheets,
 * including chart creation and formula setup for helper data.
 */

/**
 * Gets or creates the Dashboard sheet and sets its tab color.
 * Final positioning is handled by `runFullProjectInitialSetup` in `Main.js`.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} The dashboard sheet or null if an error occurs.
 */
function getOrCreateDashboardSheet(spreadsheet) {
  const FUNC_NAME = "getOrCreateDashboardSheet";
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') { 
    Logger.log(`[${FUNC_NAME} ERROR] Invalid spreadsheet object provided.`); 
    return null; 
  }
  let dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME); // From Config.gs
  if (!dashboardSheet) {
    try {
      dashboardSheet = spreadsheet.insertSheet(DASHBOARD_TAB_NAME); 
      Logger.log(`[${FUNC_NAME} INFO] Created new dashboard sheet: "${DASHBOARD_TAB_NAME}".`);
    } catch (eCreate) { 
      Logger.log(`[${FUNC_NAME} ERROR] Failed to create dashboard sheet "${DASHBOARD_TAB_NAME}": ${eCreate.message}`); 
      return null; 
    }
  } else { 
    Logger.log(`[${FUNC_NAME} INFO] Found existing dashboard sheet: "${DASHBOARD_TAB_NAME}".`); 
  }
  
  if (dashboardSheet) {
    try { dashboardSheet.setTabColor(BRAND_COLORS.CAROLINA_BLUE); } // From Config.gs
    catch (eTabColor) { Logger.log(`[${FUNC_NAME} WARN] Failed to set tab color for dashboard: ${eTabColor.message}`); }
  }
  return dashboardSheet;
}

/**
 * Formats the dashboard sheet layout, including titles, scorecards, and charts.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet object.
 * @returns {boolean} True if formatting was successful, false otherwise.
 */
function formatDashboardSheet(dashboardSheet) {
  const FUNC_NAME = "formatDashboardSheet";
  if (!dashboardSheet || typeof dashboardSheet.getName !== 'function') { 
    Logger.log(`[${FUNC_NAME} ERROR] Invalid dashboardSheet object.`); 
    return false; 
  }
  Logger.log(`[${FUNC_NAME} INFO] Starting formatting for dashboard: "${dashboardSheet.getName()}".`);

  try {
    dashboardSheet.clear({contentsOnly: false, formatOnly: false, skipConditionalFormatRules: false}); 
    try { dashboardSheet.setConditionalFormatRules([]); } 
    catch (e) { Logger.log(`[${FUNC_NAME} WARN] Could not clear conditional format rules: ${e.message}`);}
    try { dashboardSheet.setHiddenGridlines(true); } 
    catch (e) { Logger.log(`[${FUNC_NAME} WARN] Error hiding gridlines: ${e.toString()}`); }

    const MAIN_TITLE_BG = BRAND_COLORS.LAPIS_LAZULI; const HEADER_TEXT_COLOR = BRAND_COLORS.WHITE;
    const CARD_BG = BRAND_COLORS.PALE_GREY; const CARD_TEXT_COLOR = BRAND_COLORS.CHARCOAL;
    const CARD_BORDER_COLOR = BRAND_COLORS.MEDIUM_GREY_BORDER;
    const PRIMARY_VALUE_COLOR = BRAND_COLORS.PALE_ORANGE; const SECONDARY_VALUE_COLOR = BRAND_COLORS.HUNYADI_YELLOW;
    const METRIC_FONT_SIZE = 15; const METRIC_FONT_WEIGHT = "bold"; const LABEL_FONT_WEIGHT = "bold";
    const spacerColAWidth = 20; const labelWidth = 150; const valueWidth = 75; const spacerS = 15;

    dashboardSheet.getRange("A1:M1").merge().setValue("CareerSuite.AI Job Application Dashboard")
                  .setBackground(MAIN_TITLE_BG).setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight("bold")
                  .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dashboardSheet.setRowHeight(1, 45); dashboardSheet.setRowHeight(2, 10); 

    dashboardSheet.getRange("B3").setValue("Key Metrics Overview:").setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(3, 30); dashboardSheet.setRowHeight(4, 10);

    const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'`;
    const companyColLetter = columnToLetter(COMPANY_COL);
    const jobTitleColLetter = columnToLetter(JOB_TITLE_COL);
    const statusColLetter = columnToLetter(STATUS_COL);
    const peakStatusColLetter = columnToLetter(PEAK_STATUS_COL);

    // Scorecard Setup (Formulas direct to Applications sheet)
    // Row 1
    dashboardSheet.getRange("B5").setValue("Total Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("C5").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E5").setValue("Peak Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F5").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H5").setValue("Interview Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("I5").setFormula(`=IFERROR(F5/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.getRange("K5").setValue("Offer Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("L5").setFormula(`=IFERROR(F7/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(5, 40); dashboardSheet.setRowHeight(6, 10);
    // Row 2
    dashboardSheet.getRange("B7").setValue("Active Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    let activeAppsFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>"&"", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${REJECTED_STATUS}", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${ACCEPTED_STATUS}"), 0)`;
    dashboardSheet.getRange("C7").setFormula(activeAppsFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E7").setValue("Peak Offers").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${OFFER_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H7").setValue("Current Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("I7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("K7").setValue("Current Assessments").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("L7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${ASSESSMENT_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(7, 40); dashboardSheet.setRowHeight(8, 10);
    // Row 3
    dashboardSheet.getRange("B9").setValue("Total Rejections").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("C9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E9").setValue("Apps Viewed (Peak)").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${APPLICATION_VIEWED_STATUS}"),0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H9").setValue("Manual Review").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    const compColManualFormula = `${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const titleColManualFormula = `${appSheetNameForFormula}!${jobTitleColLetter}2:${jobTitleColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const statusColManualFormula = `${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const finalManualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(SIGN((${compColManualFormula})+(${titleColManualFormula})+(${statusColManualFormula})))),0)`;
    dashboardSheet.getRange("I9").setFormula(finalManualReviewFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.getRange("K9").setValue("Direct Reject Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    const directRejectFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${DEFAULT_STATUS}",${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}")/C5, 0)`;
    dashboardSheet.getRange("L9").setFormula(directRejectFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(9, 40); dashboardSheet.setRowHeight(10, 15);

    const scorecardRangesToStyle = ["B5:C5", "E5:F5", "H5:I5", "K5:L5", "B7:C7", "E7:F7", "H7:I7", "K7:L7", "B9:C9", "E9:F9", "H9:I9", "K9:L9"];
    scorecardRangesToStyle.forEach(rangeString => {
      const range = dashboardSheet.getRange(rangeString);
      range.setBackground(CARD_BG).setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
    });
    
    const chartSectionTitleRow1 = 11; dashboardSheet.getRange("B"+chartSectionTitleRow1).setValue("Platform & Weekly Trends").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(chartSectionTitleRow1, 25); dashboardSheet.setRowHeight(chartSectionTitleRow1+1, 5);
    const chartSectionTitleRow2 = 28; dashboardSheet.getRange("B"+chartSectionTitleRow2).setValue("Application Funnel Analysis").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(chartSectionTitleRow2, 25); dashboardSheet.setRowHeight(chartSectionTitleRow2+1, 5);

    dashboardSheet.setColumnWidth(1, spacerColAWidth); dashboardSheet.setColumnWidth(2, labelWidth); dashboardSheet.setColumnWidth(3, valueWidth); dashboardSheet.setColumnWidth(4, spacerS);
    dashboardSheet.setColumnWidth(5, labelWidth); dashboardSheet.setColumnWidth(6, valueWidth); dashboardSheet.setColumnWidth(7, spacerS);
    dashboardSheet.setColumnWidth(8, labelWidth); dashboardSheet.setColumnWidth(9, valueWidth); dashboardSheet.setColumnWidth(10, spacerS);
    dashboardSheet.setColumnWidth(11, labelWidth); dashboardSheet.setColumnWidth(12, valueWidth); dashboardSheet.setColumnWidth(13, spacerColAWidth);
    Logger.log(`[${FUNC_NAME} INFO] Dashboard column widths set.`);

    const lastUsedDataColumnOnDashboard = 13; const maxColsDashboard = dashboardSheet.getMaxColumns();
    if (maxColsDashboard > lastUsedDataColumnOnDashboard) dashboardSheet.deleteColumns(lastUsedDataColumnOnDashboard + 1, maxColsDashboard - lastUsedDataColumnOnDashboard);
    const lastUsedDataRowOnDashboard = 45; const maxRowsDashboard = dashboardSheet.getMaxRows();
    if (maxRowsDashboard > lastUsedDataRowOnDashboard) dashboardSheet.deleteRows(lastUsedDataRowOnDashboard + 1, maxRowsDashboard - lastUsedDataRowOnDashboard);
    
    Logger.log(`[${FUNC_NAME} INFO] Formatting concluded for Dashboard visuals and scorecards.`);
    return true;
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Major dashboard formatting error: ${e.toString()}\nStack: ${e.stack}`);
    return false;
  }
}

/**
 * Gets or creates the job data sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} The job data sheet or null if an error occurs.
 */
function getOrCreateJobDataSheet(spreadsheet) {
  const FUNC_NAME = "getOrCreateJobDataSheet";
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') { Logger.log(`[${FUNC_NAME} ERROR] Invalid spreadsheet.`); return null;}
  let jobDataSheet = spreadsheet.getSheetByName(JOB_DATA_SHEET_NAME);
  if (!jobDataSheet) {
    try { jobDataSheet = spreadsheet.insertSheet(JOB_DATA_SHEET_NAME); Logger.log(`[${FUNC_NAME} INFO] Created: "${JOB_DATA_SHEET_NAME}".`);}
    catch (eCreate) { Logger.log(`[${FUNC_NAME} ERROR] Create Fail for "${JOB_DATA_SHEET_NAME}": ${eCreate.message}`); return null;}
  } else { Logger.log(`[${FUNC_NAME} INFO] Found: "${JOB_DATA_SHEET_NAME}".`); }
  return jobDataSheet;
}

/**
 * Sets up the formulas in the `Job Data` sheet. This is called once during initial setup.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} jobDataSheet The `Job Data` sheet object.
 * @returns {boolean} True if formulas were set successfully, false otherwise.
 */
function setupJobDataSheetFormulas(jobDataSheet) {
  const FUNC_NAME = "setupJobDataSheetFormulas";
  if (!jobDataSheet || typeof jobDataSheet.getName !== 'function') {
    Logger.log(`[${FUNC_NAME} ERROR] Invalid jobDataSheet object passed.`);
    return false;
  }
  Logger.log(`[${FUNC_NAME} INFO] Setting up formulas in "${jobDataSheet.getName()}" based on NEW LOGIC structure.`);

  try {
    // Clear existing content to ensure clean state for new formulas
    const maxRows = jobDataSheet.getMaxRows();
    const maxCols = jobDataSheet.getMaxColumns();
    if (maxRows > 0 && maxCols > 0) {
        jobDataSheet.getRange(1, 1, maxRows, maxCols).clearContent().clearNote();
        Logger.log(`[${FUNC_NAME} INFO] Cleared content from job data sheet "${jobDataSheet.getName()}".`);
    }

    // --- Define sheet references and column letters needed for formulas ---
    const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'!`;
    const platformColLetter = columnToLetter(PLATFORM_COL);
    const emailDateColLetter = columnToLetter(EMAIL_DATE_COL);
    const peakStatusColLetter = columnToLetter(PEAK_STATUS_COL);
    const companyColLetter = columnToLetter(COMPANY_COL);

    // --- 1. Platform Distribution Data (Formulas for Columns A:B) ---
    jobDataSheet.getRange("A1").setValue("Platform");
    jobDataSheet.getRange("B1").setValue("Count");
    const platformQueryFormula = `=IFERROR(QUERY(${appSheetNameForFormula}${platformColLetter}2:${platformColLetter}, "SELECT Col1, COUNT(Col1) WHERE Col1 IS NOT NULL AND Col1 <> '' GROUP BY Col1 ORDER BY COUNT(Col1) DESC LABEL Col1 '', COUNT(Col1) ''", 0), {"No Platform Data",0})`;
    jobDataSheet.getRange("A2").setFormula(platformQueryFormula);
    Logger.log(`[${FUNC_NAME} INFO] Platform distribution formula set in A2: ${platformQueryFormula}`);

    // --- 2. Data for Applications Over Time (Weekly) Chart (Columns D:E) ---
    jobDataSheet.getRange("D1").setValue("Week Starting");
    jobDataSheet.getRange("E1").setValue("Applications");
    const weeklyQueryFormula = `=IFERROR(QUERY(${appSheetNameForFormula}${emailDateColLetter}2:${emailDateColLetter}, "SELECT DATE(YEAR(Col1), MONTH(Col1), DAY(Col1) - WEEKDAY(Col1, 2) + 1), COUNT(Col1) WHERE Col1 IS NOT NULL GROUP BY DATE(YEAR(Col1), MONTH(Col1), DAY(Col1) - WEEKDAY(Col1, 2) + 1) ORDER BY DATE(YEAR(Col1), MONTH(Col1), DAY(Col1) - WEEKDAY(Col1, 2) + 1) ASC LABEL DATE(YEAR(Col1), MONTH(Col1), DAY(Col1) - WEEKDAY(Col1, 2) + 1) '', COUNT(Col1) ''", 0), {"No Date Data",0})`;
    jobDataSheet.getRange("D2").setFormula(weeklyQueryFormula);
    jobDataSheet.getRange("D2:D").setNumberFormat("M/d/yyyy");
    Logger.log(`[${FUNC_NAME} INFO] Weekly applications formula set in D2: ${weeklyQueryFormula}`);

    // --- 3. Data for Application Funnel (Peak Stages) Chart (Columns G:H) ---
    jobDataSheet.getRange("G1").setValue("Stage");
    jobDataSheet.getRange("H1").setValue("Count");
    const funnelStagesValues = [DEFAULT_STATUS, APPLICATION_VIEWED_STATUS, ASSESSMENT_STATUS, INTERVIEW_STATUS, OFFER_STATUS];
    jobDataSheet.getRange(2, 7, funnelStagesValues.length, 1).setValues(funnelStagesValues.map(stage => [stage]));
    
    jobDataSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}${companyColLetter}2:${companyColLetter}),0)`);
    for (let i = 1; i < funnelStagesValues.length; i++) {
      jobDataSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}${peakStatusColLetter}2:${peakStatusColLetter}, G${i + 2}),0)`);
    }
    Logger.log(`[${FUNC_NAME} INFO] Funnel stage formulas set in G:H.`);
    
    SpreadsheetApp.flush();
    return true;
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Error setting formulas in job data sheet: ${e.toString()}\nStack: ${e.stack}`);
    return false;
  }
}


/**
 * Ensures charts on the Dashboard sheet are created or updated.
 * This function relies on the `Job Data` sheet being populated by formulas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The "Dashboard" sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} jobDataSheet The "Job Data" sheet object.
 */
function updateDashboardMetrics(dashboardSheet, jobDataSheet) {
  const FUNC_NAME = "updateDashboardMetrics";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - Verifying/Updating Charts (Job Data is Formula-Driven) ====`);

  if (!jobDataSheet || typeof jobDataSheet.getName !== 'function') { Logger.log(`[${FUNC_NAME} ERROR] Job Data sheet invalid.`); return; }
  if (!dashboardSheet && DEBUG_MODE) { Logger.log(`[${FUNC_NAME} WARN] Dashboard sheet invalid. Charts cannot be updated.`); }

  Logger.log(`[${FUNC_NAME} INFO] Job data is formula-driven. Refreshing charts on "${dashboardSheet ? dashboardSheet.getName() : 'N/A'}"...`);
  
  SpreadsheetApp.flush(); 
  Utilities.sleep(2000);

  if (dashboardSheet && jobDataSheet) {
     Logger.log(`[${FUNC_NAME} INFO] Calling chart creation/update functions...`);
     try {
        updatePlatformDistributionChart(dashboardSheet, jobDataSheet);
        updateApplicationsOverTimeChart(dashboardSheet, jobDataSheet);
        updateApplicationFunnelChart(dashboardSheet, jobDataSheet);
        Logger.log(`[${FUNC_NAME} INFO] Chart update/creation process successfully called.`);
     } catch (e) { 
        Logger.log(`[${FUNC_NAME} ERROR] Chart update calls threw an error: ${e.toString()}\nStack: ${e.stack}`); 
     }
  } else {
      Logger.log(`[${FUNC_NAME} WARN] Skipping chart object updates - dashboardSheet or jobDataSheet is missing.`);
  }
  Logger.log(`\n==== ${FUNC_NAME} FINISHED ====`);
}

/**
 * Updates the platform distribution pie chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} jobDataSheet The job data sheet.
 */
function updatePlatformDistributionChart(dashboardSheet, jobDataSheet) {
  const FUNC_NAME = "updatePlatformDistributionChart";
  const CHART_TITLE = "Platform Distribution";
  const ANCHOR_ROW = 13, ANCHOR_COL = 2;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { 
    if (chart.getOptions().get('title') === CHART_TITLE && 
        chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && 
        chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) {
      existingChart = chart;
    }
  });

  if (jobDataSheet.getRange("A1").getValue().toString().trim() !== "Platform") {
      return;
  }
  const dataRange = jobDataSheet.getRange("A:B");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${jobDataSheet.getName()}!A:B for chart "${CHART_TITLE}".`);

  const optionsObject = { 
    title: CHART_TITLE, 
    pieHole: 0.4, 
    width: 480, 
    height: 300, 
    sliceVisibilityThreshold: 0,
    is3D: true, 
    colors: BRAND_COLORS_CHART_ARRAY() 
  };
  
  try {
    let chartBuilder;
    if (existingChart) { 
      chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.PIE); 
    } else { 
      chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(dataRange); 
    }

    for (const key in optionsObject) { 
      chartBuilder = chartBuilder.setOption(key, optionsObject[key]); 
    } 
    
    chartBuilder = chartBuilder.setOption('legend', 'right');

    if (existingChart) {
      dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    } else {
      dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    }
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { 
    Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); 
  }
}

/**
 * Updates the applications over time line chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} jobDataSheet The job data sheet.
 */
function updateApplicationsOverTimeChart(dashboardSheet, jobDataSheet) {
  const FUNC_NAME = "updateApplicationsOverTimeChart";
  const CHART_TITLE = "Applications Over Time (Weekly)";
  const ANCHOR_ROW = 13, ANCHOR_COL = 8;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE && chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) existingChart = chart; });

  if (jobDataSheet.getRange("D1").getValue().toString().trim() !== "Week Starting") {
      Logger.log(`[${FUNC_NAME} WARN] Header "Week Starting" missing in Job Data D1. Removing chart: ${CHART_TITLE}`);
      if (existingChart) try { dashboardSheet.removeChart(existingChart); } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Remove chart: ${e.message}`); }
      return;
  }
  const dataRange = jobDataSheet.getRange("D:E");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${jobDataSheet.getName()}!D:E for chart "${CHART_TITLE}".`);
  
  const optionsObject = { title:CHART_TITLE, hAxis:{title:'Week Starting',textStyle:{fontSize:10},format:'M/d', gridlines: {color: '#EEE'}}, vAxis:{title:'Applications',textStyle:{fontSize:10},viewWindow:{min:0},gridlines:{count:-1, color: '#CCC'}}, legend:{position:'none'}, colors:[BRAND_COLORS.LAPIS_LAZULI], width:480, height:300, pointSize: 5, lineWidth: 2 };
  try {
    let chartBuilder;
    if (existingChart) { chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.LINE); }
    else { chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.LINE).addRange(dataRange); }
    for (const key in optionsObject) { chartBuilder = chartBuilder.setOption(key, optionsObject[key]); }

    if (existingChart) dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    else dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); }
}

/**
 * Updates the application funnel column chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} jobDataSheet The job data sheet.
 */
function updateApplicationFunnelChart(dashboardSheet, jobDataSheet) {
  const FUNC_NAME = "updateApplicationFunnelChart";
  const CHART_TITLE = "Application Funnel (Peak Stages)";
  const ANCHOR_ROW = 30, ANCHOR_COL = 2;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE && chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) existingChart = chart; });
  
  if (jobDataSheet.getRange("G1").getValue().toString().trim() !== "Stage") {
      Logger.log(`[${FUNC_NAME} WARN] Header "Stage" missing in Job Data G1. Removing chart: ${CHART_TITLE}`);
      if (existingChart) try { dashboardSheet.removeChart(existingChart); } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Remove chart: ${e.message}`); }
      return;
  }
  const dataRange = jobDataSheet.getRange("G:H");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${jobDataSheet.getName()}!G:H for chart "${CHART_TITLE}".`);

  const optionsObject = { title:CHART_TITLE, hAxis:{title:'Application Stage',textStyle:{fontSize:10},slantedText:true,slantedTextAngle:30}, vAxis:{title:'Applications',textStyle:{fontSize:10},viewWindow:{min:0},gridlines:{count:-1, color: '#CCC'}}, legend:{position:'none'}, colors:[BRAND_COLORS.CAROLINA_BLUE], bar:{groupWidth:'60%'}, width:480, height:300 };
  try {
    let chartBuilder;
    if (existingChart) { chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.COLUMN); }
    else { chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(dataRange); }
    for (const key in optionsObject) { chartBuilder = chartBuilder.setOption(key, optionsObject[key]); }

    if (existingChart) dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    else dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); }
}

/**
 * Returns an array of brand colors for charts.
 * @returns {string[]} An array of hex color codes.
 */
function BRAND_COLORS_CHART_ARRAY() {
    return [
        BRAND_COLORS.LAPIS_LAZULI,
        BRAND_COLORS.CAROLINA_BLUE,
        BRAND_COLORS.HUNYADI_YELLOW,
        BRAND_COLORS.PALE_ORANGE,
        BRAND_COLORS.CHARCOAL,
        // Adding a few more standard, visually distinct colors as fallbacks
        "#27AE60", // Green
        "#8E44AD", // Purple
        "#E67E22", // Orange
        "#16A085", // Teal
        "#C0392B"  // Red
    ];
}
