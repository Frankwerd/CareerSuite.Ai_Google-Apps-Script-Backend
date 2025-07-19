/**
 * @file Contains the primary functions for the Job Leads Tracker module,
 * including initial setup of the leads sheet/labels/filters and the
 * ongoing processing of job lead emails.
 */

/**
 * Sets up the Job Leads Tracker module.
 * This function ensures the "Potential Job Leads" sheet exists and is formatted,
 * creates the necessary Gmail labels and filter, and sets up a daily trigger
 * for processing new job leads.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} passedSpreadsheet The spreadsheet object to set up.
 * @returns {{success: boolean, messages: string[]}} An object containing the setup result and messages.
 */
function runInitialSetup_JobLeadsModule(passedSpreadsheet) {
    const leadsConfig = {
        activeSS: passedSpreadsheet,
        moduleName: "Job Leads Tracker",
        sheetTabName: LEADS_SHEET_TAB_NAME,
        sheetHeaders: LEADS_SHEET_HEADERS,
        columnWidths: LEADS_SHEET_COLUMN_WIDTHS,
        bandingTheme: SpreadsheetApp.BandingTheme.YELLOW,
        tabColor: BRAND_COLORS.HUNYADI_YELLOW,
        gmailLabelParent: LEADS_GMAIL_LABEL_PARENT,
        gmailLabelToProcess: LEADS_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: LEADS_GMAIL_LABEL_PROCESSED,
        gmailFilterQuery: LEADS_GMAIL_FILTER_QUERY,
        triggerFunctionName: 'processJobLeads',
        triggerIntervalHours: 3
    };
    return _setupModule(leadsConfig);
}

/**
 * Processes emails that have been labeled for job lead extraction.
 * This function is designed to be run on a time-driven trigger. It fetches emails,
 * calls the Gemini API to extract job leads, and writes the results to the spreadsheet.
 */
function _leadsParser(subject, body, key) {
    return callGemini_forJobLeads(body, key);
}

function _leadsDataHandler(geminiResult, message, companyIndex, dataSheet) {
    const allTheNewRows = [];
    const requiresManualReview = false; // Not really applicable for leads

    if (geminiResult && geminiResult.success) {
        const extractedJobsArray = parseGeminiResponse_forJobLeads(geminiResult.data);
        if (extractedJobsArray && extractedJobsArray.length > 0) {
            Logger.log(`[_leadsDataHandler INFO] Gemini extracted ${extractedJobsArray.length} job(s) from msg ${message.getId()}.`);
            for (const jobData of extractedJobsArray) {
                if (jobData && jobData.jobTitle && String(jobData.jobTitle).toLowerCase() !== 'n/a' && String(jobData.jobTitle).toLowerCase() !== 'error') {
                    const rowDataForSheet = new Array(TOTAL_COLUMNS_IN_LEADS_SHEET).fill("");
                    rowDataForSheet[LEADS_DATE_ADDED_COL - 1] = message.getDate();
                    rowDataForSheet[LEADS_COMPANY_COL - 1] = jobData.company || "N/A";
                    rowDataForSheet[LEADS_JOB_TITLE_COL - 1] = jobData.jobTitle || "N/A";
                    rowDataForSheet[LEADS_LOCATION_COL - 1] = jobData.location || "N/A";
                    rowDataForSheet[LEADS_SALARY_PAY_COL - 1] = jobData.salaryPay || "N/A";
                    rowDataForSheet[LEADS_SOURCE_LINK_COL - 1] = jobData.jobUrl || "N/A";
                    rowDataForSheet[LEADS_NOTES_COL - 1] = jobData.notes || "";
                    rowDataForSheet[LEADS_STATUS_COL - 1] = DEFAULT_LEAD_STATUS;
                    rowDataForSheet[LEADS_FOLLOW_UP_COL - 1] = "";
                    rowDataForSheet[LEADS_EMAIL_SUBJECT_COL - 1] = message.getSubject().substring(0, 500);
                    rowDataForSheet[LEADS_EMAIL_ID_COL - 1] = message.getId();
                    rowDataForSheet[LEADS_PROCESSED_TIMESTAMP_COL - 1] = new Date();
                    allTheNewRows.push(rowDataForSheet);
                } else {
                    if (DEBUG_MODE) Logger.log(`[_leadsDataHandler DEBUG] Job from msg ${message.getId()} was N/A/error or missing title. Skipping sheet write: ${JSON.stringify(jobData)}`);
                }
            }
        } else {
            Logger.log(`[_leadsDataHandler INFO] Msg ${message.getId()}: Gemini API call success, but parsing response yielded no distinct job listings.`);
        }
    } else {
        const errorInfo = {
            moduleName: "Job Leads Tracker",
            errorType: "Gemini API Error",
            details: geminiResult.error || "Unknown API error",
            messageSubject: message.getSubject(),
            messageId: message.getId()
        };
        _writeErrorToSheet(dataSheet, errorInfo);
        Logger.log(`[_leadsDataHandler ERROR] Gemini API FAILED for msg ${message.getId()}. Error: ${geminiResult ? geminiResult.error : 'Null or unexpected API response object'}`);
    }

    return { newRowData: allTheNewRows, requiresManualReview };
}

function processJobLeads() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scriptProperties = PropertiesService.getScriptProperties();
    const leadsProcessingConfig = {
        moduleName: "Job Leads Tracker",
        sheetTabName: LEADS_SHEET_TAB_NAME,
        gmailLabelToProcess: LEADS_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: LEADS_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: null,
        parserFunction: _leadsParser,
        dataHandler: _leadsDataHandler,
    };
    _processingEngine(leadsProcessingConfig, ss, scriptProperties);
}
