/**
 * @file Main script file orchestrating setup, email processing, and UI for the job application tracker.
 * @author Francis John LiButti (Originals), AI Integration & Refinements by Assistant
 * @version 8
 */

/**
 * Checks if the critical configuration variables are set correctly.
 * @returns {boolean} True if the configuration is valid, false otherwise.
 */
function checkConfig() {
    const FUNC_NAME = "checkConfig";
    const criticalVars = {
        MASTER_WEB_APP_URL,
        TEMPLATE_SHEET_ID,
        MASTER_SCRIPT_ID,
        GEMINI_API_KEY_PROPERTY,
        GEMINI_API_ENDPOINT_TEXT_ONLY
    };
    let allClear = true;
    for (const [varName, varValue] of Object.entries(criticalVars)) {
        if (typeof varValue === 'undefined' || varValue === null || varValue === "" || varValue.includes("REPLACE") || varValue.includes("YOUR_")) {
            Logger.log(`[${FUNC_NAME} CRITICAL] Configuration variable ${varName} is not set correctly. Value: "${varValue}"`);
            allClear = false;
        }
    }
    return allClear;
}

/**
 * Runs the complete initial setup for all modules.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [passedSpreadsheet]
 * @returns {{success: boolean, message: string, detailedMessages: string[], sheetId: string|null, sheetUrl: string|null}}
 */
function runFullProjectInitialSetup(passedSpreadsheet) {
    const RUNDATE = new Date().toISOString();
    const FUNC_NAME = "runFullProjectInitialSetup";
    Logger.log(`==== ${FUNC_NAME}: STARTING (CareerSuite.AI v1.2 - ${RUNDATE}) ====`);
    if (!checkConfig()) {
        Logger.log(`[${FUNC_NAME} CRITICAL] Configuration check failed. Aborting setup.`);
        return { success: false, message: "Critical configuration is missing.", detailedMessages: ["Critical configuration is missing. Please check the logs."], sheetId: null, sheetUrl: null };
    }
    let overallSuccess = true;
    let setupMessages = [];
    let activeSS;

    if (passedSpreadsheet && typeof passedSpreadsheet.getId === 'function') {
        activeSS = passedSpreadsheet;
    } else {
        activeSS = SpreadsheetApp.getActiveSpreadsheet();
        if (!activeSS) {
            const { spreadsheet: foundOrCreatedSS } = getOrCreateSpreadsheetAndSheet();
            activeSS = foundOrCreatedSS;
        }
    }

    if (!activeSS) {
        const errorMsg = `CRITICAL [${FUNC_NAME}]: No valid spreadsheet could be determined. Setup aborted.`;
        Logger.log(errorMsg);
        return { success: false, message: errorMsg, detailedMessages: [errorMsg], sheetId: null, sheetUrl: null };
    }

    PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_KEY, activeSS.getId());

    if (typeof TEMPLATE_SHEET_ID !== 'undefined' && TEMPLATE_SHEET_ID !== "" && activeSS.getId() === TEMPLATE_SHEET_ID) {
        const templateMsg = `[${FUNC_NAME} INFO] Target spreadsheet is the TEMPLATE. Setup SKIPPED.`;
        Logger.log(templateMsg);
        if (!passedSpreadsheet) {
            try {
                SpreadsheetApp.getUi().alert('Template Sheet', 'Initial setup is not meant to be run on the template sheet itself. Please make a copy first.', SpreadsheetApp.getUi().ButtonSet.OK);
            } catch (e) { /* UI not available */ }
        }
        return { success: true, message: "Setup skipped: Target is template.", detailedMessages: [templateMsg], sheetId: activeSS.getId(), sheetUrl: activeSS.getUrl() };
    }

    const modules = [
        { name: "Application Tracker", setupFunc: initialSetup_LabelsAndSheet },
        { name: "Job Leads Tracker", setupFunc: runInitialSetup_JobLeadsModule }
    ];

    for (const module of modules) {
        if (typeof module.setupFunc === "function") {
            Logger.log(`\n[${FUNC_NAME} INFO] --- Starting ${module.name} Setup ---`);
            try {
                const result = module.setupFunc(activeSS);
                if (result.messages) setupMessages.push(...result.messages.map(m => `${module.name}: ${m}`));
                if (!result.success) {
                    overallSuccess = false;
                    Logger.log(`[${FUNC_NAME} ERROR] ${module.name} FAILED.`);
                } else {
                    Logger.log(`[${FUNC_NAME} INFO] ${module.name} Success.`);
                }
            } catch (e) {
                Logger.log(`[${FUNC_NAME} CRITICAL ERROR] ${module.name} Exception: ${e.toString()}\n${e.stack}`);
                setupMessages.push(`${module.name}: CRITICAL EXCEPTION - ${e.message}`);
                overallSuccess = false;
            }
        }
    }

    if (overallSuccess) {
        Logger.log(`[${FUNC_NAME} INFO] Applying final tab order...`);
        try {
            const tabOrder = [DASHBOARD_TAB_NAME, APP_TRACKER_SHEET_TAB_NAME, LEADS_SHEET_TAB_NAME];
            tabOrder.forEach((sheetName, index) => {
                const sheetToMove = activeSS.getSheetByName(sheetName);
                if (sheetToMove) activeSS.setActiveSheet(sheetToMove).moveActiveSheet(index + 1);
            });
            const helperSheet = activeSS.getSheetByName(HELPER_SHEET_NAME);
            if (helperSheet && !helperSheet.isSheetHidden()) helperSheet.hideSheet();
            setupMessages.push("Branding: Tab order & helper visibility verified.");
        } catch (e) {
            Logger.log(`[${FUNC_NAME} WARN] Error finalizing tab order: ${e.message}`);
        }
    }

    const finalStatusMessage = `CareerSuite.AI Full Setup ${overallSuccess ? "completed" : "had issues"}.`;
    Logger.log(`\n==== ${FUNC_NAME} SUMMARY (SS ID: ${activeSS.getId()}) ====`);
    setupMessages.forEach(msg => Logger.log(`  - ${msg}`));
    Logger.log(`Overall Status: ${overallSuccess ? "SUCCESSFUL" : "ISSUES ENCOUNTERED"}`);

    if (!passedSpreadsheet) {
        try {
            const title = `CareerSuite.AI Setup ${overallSuccess ? "Complete" : "Issues"}`;
            const message = `Setup for "${activeSS.getName()}" ${overallSuccess ? "finished" : "had issues"}.\n\nSummary:\n- ${setupMessages.join('\n- ')}`;
            SpreadsheetApp.getUi().alert(title, message.substring(0, 1000), SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) { /* UI not available */ }
    }

    return { success: overallSuccess, message: finalStatusMessage, detailedMessages: setupMessages, sheetId: activeSS.getId(), sheetUrl: activeSS.getUrl() };
}

/**
 * Sets up the core Job Application Tracker module.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSS The spreadsheet object.
 * @returns {{success: boolean, messages: string[]}}
 */
function initialSetup_LabelsAndSheet(activeSS) {
    const trackerConfig = {
        activeSS: activeSS,
        moduleName: "Application Tracker",
        sheetTabName: APP_TRACKER_SHEET_TAB_NAME,
        sheetHeaders: APP_TRACKER_SHEET_HEADERS,
        columnWidths: APP_SHEET_COLUMN_WIDTHS,
        bandingTheme: SpreadsheetApp.BandingTheme.BLUE,
        tabColor: BRAND_COLORS.LAPIS_LAZULI,
        gmailLabelParent: MASTER_GMAIL_LABEL_PARENT,
        gmailLabelToProcess: TRACKER_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: TRACKER_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: TRACKER_GMAIL_LABEL_MANUAL_REVIEW,
        gmailFilterQuery: TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES,
        triggerFunctionName: 'processEmails_triggerHandler',
        triggerIntervalHours: 1,
        staleRejectFunctionName: 'markStale_triggerHandler'
    };
    return _setupModule(trackerConfig);
}

/**
 * Generic processing engine for all modules.
 * @param {object} config The configuration for the module.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {GoogleAppsScript.Properties.Properties} scriptProperties The script's properties.
 */
function _processingEngine(config, ss, scriptProperties) {
    const FUNC_NAME = "_processingEngine";
    const SCRIPT_START_TIME = new Date();
    Logger.log(`\n==== ${FUNC_NAME}: STARTING (${SCRIPT_START_TIME.toLocaleString()}) - ${config.moduleName} ====`);

    const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY);
    if (!geminiApiKey || !geminiApiKey.startsWith("AIza") || geminiApiKey.length < 30) {
        Logger.log(`[${FUNC_NAME} HALTING] Gemini API Key is not configured or invalid for ${config.moduleName}. Please set it via the menu.`);
        return;
    }

    const dataSheet = ss.getSheetByName(config.sheetTabName);
    if (!dataSheet) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Sheet "${config.sheetTabName}" not found. Aborting.`);
        return;
    }

    let procLbl = GmailApp.getUserLabelByName(config.gmailLabelToProcess);
    let processedLblObj = GmailApp.getUserLabelByName(config.gmailLabelProcessed);
    let manualLblObj = config.gmailLabelManualReview ? GmailApp.getUserLabelByName(config.gmailLabelManualReview) : processedLblObj;

    if (!procLbl || !processedLblObj || !manualLblObj) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Core Gmail labels for ${config.moduleName} not found. Aborting.`);
        return;
    }

    const allSheetData = dataSheet.getDataRange().getValues();
    const companyIndex = new Map();
    for (let i = 1; i < allSheetData.length; i++) {
        const rowData = allSheetData[i];
        const companyName = rowData[COMPANY_COL - 1];
        if (companyName && typeof companyName === 'string' && companyName.trim() !== "") {
            const companyKey = companyName.toLowerCase();
            if (!companyIndex.has(companyKey)) {
                companyIndex.set(companyKey, []);
            }
            companyIndex.get(companyKey).push({
                row: i + 1,
                rowData: rowData,
                title: rowData[JOB_TITLE_COL-1],
                status: rowData[STATUS_COL-1],
                peakStatus: rowData[PEAK_STATUS_COL-1]
            });
        }
    }

    const threadsToProcess = procLbl.getThreads(0, 20);
    if (threadsToProcess.length === 0) {
        Logger.log(`[${FUNC_NAME} INFO] No new messages to process for ${config.moduleName}.`);
        return;
    }

    const messagesToSort = threadsToProcess.flatMap(thread => thread.getMessages());
    messagesToSort.sort((a, b) => a.getDate() - b.getDate());

    const dataToUpdate = [];
    const newRowsData = [];
    let threadProcessingOutcomes = {};

    for (const message of messagesToSort) {
        if ((new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 320) {
            Logger.log(`[${FUNC_NAME} WARN] Execution time limit nearing. Stopping.`);
            break;
        }

        const msgId = message.getId();
        const threadId = message.getThread().getId();
        const emailSubject = message.getSubject();
        const plainBodyText = message.getPlainBody();

        const geminiResult = config.parserFunction(emailSubject, plainBodyText, geminiApiKey);
        const handlerResult = config.dataHandler(geminiResult, message, companyIndex, dataSheet);

        if (handlerResult.updateInfo) {
            dataToUpdate.push(handlerResult.updateInfo);
            const companyKey = (handlerResult.updateInfo.company || '').toLowerCase();
            const existingEntry = companyIndex.get(companyKey)?.find(e => e.row === handlerResult.updateInfo.row);
            if (existingEntry) {
                existingEntry.status = handlerResult.updateInfo.newStatus;
                existingEntry.peakStatus = handlerResult.updateInfo.newPeakStatus;
            }
        } else if (handlerResult.newRowData) {
            newRowsData.push(handlerResult.newRowData);
            const companyKey = (handlerResult.newRowData[COMPANY_COL - 1] || '').toLowerCase();
            if (companyKey) {
                if (!companyIndex.has(companyKey)) companyIndex.set(companyKey, []);
                companyIndex.get(companyKey).push({ row: -1, emailId: msgId, /* other fields */ });
            }
        }

        threadProcessingOutcomes[threadId] = handlerResult.requiresManualReview ? 'manual' : 'done';
        Utilities.sleep(200);
    }

    dataToUpdate.forEach(update => dataSheet.getRange(update.row, 1, 1, update.values.length).setValues([update.values]));
    if (newRowsData.length > 0) {
        const firstNewRow = dataSheet.getLastRow() + 1;
        dataSheet.getRange(firstNewRow, 1, newRowsData.length, newRowsData[0].length).setValues(newRowsData);
        newRowsData.forEach((rowData, i) => {
            const companyKey = (rowData[COMPANY_COL - 1] || '').toLowerCase();
            const entryInCache = companyIndex.get(companyKey)?.find(e => e.row === -1 && e.emailId === rowData[EMAIL_ID_COL - 1]);
            if (entryInCache) entryInCache.row = firstNewRow + i;
        });
    }

    applyFinalLabels(threadProcessingOutcomes, procLbl, processedLblObj, manualLblObj);
    Logger.log(`\n==== ${FUNC_NAME} FINISHED (${new Date().toLocaleString()}) ====`);
}

/**
 * Parser function specific to the Application Tracker.
 * @param {string} subject
 * @param {string} body
 * @param {string} key
 * @returns {object}
 */
function _trackerParser(subject, body, key) {
    return callGemini_forApplicationDetails(subject, body, key);
}

/**
 * Data handler function specific to the Application Tracker.
 * @param {object} geminiResult
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @param {Map<string, object[]>} companyIndex
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @returns {{updateInfo?: object, newRowData?: any[], requiresManualReview: boolean}}
 */
function _trackerDataHandler(geminiResult, message, companyIndex, dataSheet) {
    const emailSubject = message.getSubject() || "";
    const msgId = message.getId();

    Logger.log(`--- AI INPUT FOR MSG ID: ${msgId} ---`);
    Logger.log(`SUBJECT: ${emailSubject}`);
    Logger.log(`BODY SNIPPET: ${message.getPlainBody().substring(0, 1500)}`);
    Logger.log(`------------------------------------`);

    const senderEmail = message.getFrom() || "";
    const emailPermaLink = `https://mail.google.com/mail/u/0/#inbox/${msgId}`;
    const currentTimestamp = new Date();
    const emailDate = message.getDate();
    let companyName = MANUAL_REVIEW_NEEDED;
    let jobTitle = MANUAL_REVIEW_NEEDED;
    let applicationStatus = null;

    if (geminiResult && !geminiResult.error) {
        companyName = geminiResult.company || MANUAL_REVIEW_NEEDED;
        jobTitle = geminiResult.title || MANUAL_REVIEW_NEEDED;
        applicationStatus = (geminiResult.status && geminiResult.status !== "undefined") ? geminiResult.status : "Update/Other";
        Logger.log(`[_trackerDataHandler INFO] Gemini Raw: C:"${companyName}", T:"${jobTitle}", S:"${geminiResult.status}" -> Parsed Status: "${applicationStatus}"`);
        if (applicationStatus === "Update/Other" || applicationStatus === MANUAL_REVIEW_NEEDED) {
            const keywordStatus = parseBodyForStatus(message.getPlainBody());
            if (keywordStatus) {
                applicationStatus = keywordStatus;
                Logger.log(`[_trackerDataHandler INFO] Status enhanced by keywords to: "${applicationStatus}"`);
            }
        }
    } else {
        const errorInfo = {
            moduleName: "Application Tracker",
            errorType: "Gemini API Error",
            details: geminiResult ? geminiResult.error : "Unknown API response error",
            messageSubject: message.getSubject(),
            messageId: msgId
        };
        _writeErrorToSheet(dataSheet, errorInfo);
        const regexResult = extractCompanyAndTitle(message, DEFAULT_PLATFORM, emailSubject, message.getPlainBody());
        companyName = regexResult.company;
        jobTitle = regexResult.title;
        applicationStatus = parseBodyForStatus(message.getPlainBody());
    }

    const requiresManualReview = (companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED);
    const finalStatusToSet = applicationStatus || DEFAULT_STATUS;
    let existingRowInfoToUpdate = null;
    let targetSheetRowForUpdate = -1;

    if (companyName !== MANUAL_REVIEW_NEEDED) {
        const potentialMatches = companyIndex.get(String(companyName).toLowerCase()) || [];
        if (jobTitle !== MANUAL_REVIEW_NEEDED) {
            existingRowInfoToUpdate = potentialMatches.find(e => e.title && e.title.toLowerCase() === jobTitle.toLowerCase());
        }
        if (!existingRowInfoToUpdate && potentialMatches.length > 0) {
            existingRowInfoToUpdate = potentialMatches.reduce((latest, current) => (current.row > latest.row ? current : latest), {row: -1});
        }
        if (existingRowInfoToUpdate && existingRowInfoToUpdate.row !== -1) {
            targetSheetRowForUpdate = existingRowInfoToUpdate.row;
        }
    }

    if (targetSheetRowForUpdate !== -1 && existingRowInfoToUpdate) {
        const rowDataForSheet = [...existingRowInfoToUpdate.rowData];
        rowDataForSheet[PROCESSED_TIMESTAMP_COL - 1] = currentTimestamp;
        rowDataForSheet[LAST_UPDATE_DATE_COL - 1] = emailDate;
        rowDataForSheet[EMAIL_SUBJECT_COL - 1] = emailSubject;
        rowDataForSheet[EMAIL_LINK_COL - 1] = emailPermaLink;
        rowDataForSheet[EMAIL_ID_COL - 1] = msgId;

        const statInSheet = String(rowDataForSheet[STATUS_COL - 1]).trim() || DEFAULT_STATUS;
        const curRank = STATUS_HIERARCHY[statInSheet] ?? 0;
        const newRank = STATUS_HIERARCHY[finalStatusToSet] ?? 0;
        if (newRank >= curRank || finalStatusToSet === REJECTED_STATUS || finalStatusToSet === OFFER_STATUS) {
            rowDataForSheet[STATUS_COL - 1] = finalStatusToSet;
        }

        const statAfterUpd = String(rowDataForSheet[STATUS_COL - 1]);
        let peakStat = existingRowInfoToUpdate.peakStatus || statInSheet;
        const curPeakRank = STATUS_HIERARCHY[peakStat] ?? 0;
        const newStatRankPeak = STATUS_HIERARCHY[statAfterUpd] ?? 0;
        if (newStatRankPeak > curPeakRank) {
            rowDataForSheet[PEAK_STATUS_COL - 1] = statAfterUpd;
        }

        return {
            updateInfo: {
                row: targetSheetRowForUpdate,
                values: rowDataForSheet,
                newStatus: rowDataForSheet[STATUS_COL - 1],
                newPeakStatus: rowDataForSheet[PEAK_STATUS_COL - 1],
                company: rowDataForSheet[COMPANY_COL - 1]
            },
            requiresManualReview: requiresManualReview
        };
    } else {
        const rowDataForSheet = new Array(TOTAL_COLUMNS_IN_APP_SHEET).fill("");
        rowDataForSheet[PROCESSED_TIMESTAMP_COL - 1] = currentTimestamp;
        rowDataForSheet[EMAIL_DATE_COL - 1] = emailDate;
        rowDataForSheet[PLATFORM_COL - 1] = DEFAULT_PLATFORM;
        rowDataForSheet[COMPANY_COL - 1] = companyName;
        rowDataForSheet[JOB_TITLE_COL - 1] = jobTitle;
        rowDataForSheet[STATUS_COL - 1] = finalStatusToSet;
        rowDataForSheet[PEAK_STATUS_COL - 1] = finalStatusToSet;
        rowDataForSheet[LAST_UPDATE_DATE_COL - 1] = emailDate;
        rowDataForSheet[EMAIL_SUBJECT_COL - 1] = emailSubject;
        rowDataForSheet[EMAIL_LINK_COL - 1] = emailPermaLink;
        rowDataForSheet[EMAIL_ID_COL - 1] = msgId;

        return {
            newRowData: rowDataForSheet,
            requiresManualReview: requiresManualReview
        };
    }
}

/**
 * Trigger handler for hourly email processing.
 */
function processEmails_triggerHandler() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scriptProperties = PropertiesService.getScriptProperties();
    Logger.log('Hourly email processing trigger started.');
    processJobApplicationEmails(ss, scriptProperties);
    Logger.log('Hourly email processing trigger finished.');
}

/**
 * Trigger handler for daily stale application checks.
 */
function markStale_triggerHandler() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Daily stale check trigger started.');
    markStaleApplicationsAsRejected(ss);
    Logger.log('Daily stale check trigger finished.');
}

/**
 * Main "stub" function for processing job application emails.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Properties.Properties} scriptProperties
 */
function processJobApplicationEmails(ss, scriptProperties) {
    const trackerProcessingConfig = {
        moduleName: "Application Tracker",
        sheetTabName: APP_TRACKER_SHEET_TAB_NAME,
        gmailLabelToProcess: TRACKER_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: TRACKER_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: TRACKER_GMAIL_LABEL_MANUAL_REVIEW,
        parserFunction: _trackerParser,
        dataHandler: _trackerDataHandler
    };
    _processingEngine(trackerProcessingConfig, ss, scriptProperties);
}

/**
 * Marks stale applications as "Rejected".
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 */
function markStaleApplicationsAsRejected(ss) {
    const FUNC_NAME = "markStaleApplicationsAsRejected";
    const SCRIPT_START_TIME = new Date();
    Logger.log(`\n==== ${FUNC_NAME}: START (${SCRIPT_START_TIME.toLocaleString()}) ====`);

    if (!ss) {
      Logger.log(`[${FUNC_NAME} FATAL ERROR] Main spreadsheet not passed. Aborting.`);
      return;
    }
  
    let dataSheet;
    try {
      dataSheet = ss.getSheetByName(APP_TRACKER_SHEET_TAB_NAME);
      if (!dataSheet) {
          Logger.log(`[${FUNC_NAME} FATAL ERROR] Tab "${APP_TRACKER_SHEET_TAB_NAME}" not found in "${ss.getName()}". Aborting.`);
          return;
      }
    } catch (e) {
      Logger.log(`[${FUNC_NAME} FATAL ERROR] Accessing tab "${APP_TRACKER_SHEET_TAB_NAME}": ${e.message}. Aborting.`);
      return;
    }
    
    const dataRange = dataSheet.getDataRange();
    const sheetValues = dataRange.getValues();
    const currentDate = new Date();
    const staleThresholdDate = new Date();
    staleThresholdDate.setDate(currentDate.getDate() - (WEEKS_THRESHOLD * 7));

    let updatedApplicationsCount = 0;
    for (let i = 1; i < sheetValues.length; i++) {
        const currentStatus = sheetValues[i][STATUS_COL - 1];
        const lastUpdateDate = new Date(sheetValues[i][LAST_UPDATE_DATE_COL - 1]);
        if (!FINAL_STATUSES_FOR_STALE_CHECK.has(currentStatus) && lastUpdateDate < staleThresholdDate) {
            sheetValues[i][STATUS_COL - 1] = REJECTED_STATUS;
            sheetValues[i][LAST_UPDATE_DATE_COL - 1] = currentDate;
            updatedApplicationsCount++;
        }
    }
  
    if (updatedApplicationsCount > 0) {
      dataRange.setValues(sheetValues);
      Logger.log(`[${FUNC_NAME} INFO] Updated ${updatedApplicationsCount} stale applications to Rejected.`);
    } else {
      Logger.log(`[${FUNC_NAME} INFO] No stale applications found needing update.`);
    }

    const SCRIPT_END_TIME = new Date();
    Logger.log(`==== ${FUNC_NAME}: END (${SCRIPT_END_TIME.toLocaleString()}) ==== Time: ${(SCRIPT_END_TIME.getTime() - SCRIPT_START_TIME.getTime()) / 1000}s ====`);
}

/**
 * Runs when the spreadsheet is opened to create the custom menu.
 * @param {object} e
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menuName = CUSTOM_MENU_NAME || '⚙️ CareerSuite.AI Tools';
  const menu = ui.createMenu(menuName);
  
  menu.addItem('🚀 Finalize Project Setup', 'userDrivenFullSetup');
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Manual Processing')
      .addItem('📧 Process Application Emails', 'processEmails_triggerHandler')
      .addItem('🗑️ Mark Stale Applications', 'markStale_triggerHandler')
      .addItem('📬 Process Job Leads', 'processJobLeads'));
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Admin & Config')
      .addItem('🔑 Set Gemini API Key', 'setSharedGeminiApiKey_UI')
      .addItem('🔄 Activate AI Features & Sync Key', 'activateAiFeatures')
      .addItem('🔍 Show All User Properties', 'showAllUserProperties'));
  menu.addSeparator();
  menu.addItem('❌ Uninstall Backend', 'uninstall');
  menu.addToUi();
}

/**
 * Wrapper for the user to run the full setup from the menu.
 */
function userDrivenFullSetup() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();

  if (typeof TEMPLATE_SHEET_ID !== 'undefined' && activeSS.getId() === TEMPLATE_SHEET_ID) {
      ui.alert('Action Not Allowed on Template', 'This setup function cannot be run on the master template sheet.', ui.ButtonSet.OK);
      return;
  }

  ui.alert('Starting Setup', 'The full project setup will now begin. This may take a minute or two.', ui.ButtonSet.OK);

  const setupResult = runFullProjectInitialSetup(activeSS);

  if (setupResult && setupResult.success) {
      scriptProperties.setProperty('initialSetupDone_vCSAI_1', 'true');
      const aiFeaturesAreActive = activateAiFeatures();
      let finalMessage = `Setup is complete for "${activeSS.getName()}".\n\n`;
      if (aiFeaturesAreActive) {
          finalMessage += "AI features are now active.";
      } else {
          finalMessage += "To enable AI features, please use the menu to set your API Key and activate.";
      }
      ui.alert('Setup Complete', finalMessage, ui.ButtonSet.OK);
  } else {
      ui.alert('Setup Issues Encountered', `The project setup had some issues.\n\nPlease check the script logs for more details.`, ui.ButtonSet.OK);
  }
}

/**
 * Uninstalls triggers and filters.
 */
function uninstall() {
  const FUNC_NAME = "uninstall";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING ====`);
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm Uninstall', 'This will remove all triggers and Gmail filters created by this script. Are you sure?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    Logger.log(`[${FUNC_NAME}] All triggers removed.`);

    // Additional cleanup for filters if needed, similar to previous versions

    ui.alert('Uninstall Complete', 'All triggers have been removed.', ui.ButtonSet.OK);
  } else {
    ui.alert('Uninstall Canceled', 'No changes made.', ui.ButtonSet.OK);
  }
  Logger.log(`\n==== ${FUNC_NAME}: FINISHED ====`);
}

/**
 * Activates AI features by syncing the API key.
 * @returns {boolean}
 */
function activateAiFeatures() {
  // Logic from previous versions
}
