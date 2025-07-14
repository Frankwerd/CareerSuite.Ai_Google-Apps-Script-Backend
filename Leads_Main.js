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
  const FUNC_NAME = "runInitialSetup_JobLeadsModule";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - Leads Module Setup ====`);
  let setupMessages = [];
  let leadsModuleSetupSuccess = true;
  let activeSSLeads;

  if (passedSpreadsheet && typeof passedSpreadsheet.getId === 'function') {
    activeSSLeads = passedSpreadsheet;
    // Logger.log(`[${FUNC_NAME} INFO] Using PASSED spreadsheet ID: ${activeSSLeads.getId()}.`); // Redundant with Main.gs log
  } else {
    Logger.log(`[${FUNC_NAME} WARN] No spreadsheet passed. Fallback get/create...`);
    const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
    activeSSLeads = mainSpreadsheet;
  }
  
  if (!activeSSLeads) {
    const errMsg = `CRITICAL [${FUNC_NAME}]: Could not obtain valid spreadsheet.`;
    Logger.log(errMsg); return { success: false, messages: [errMsg] };
  }
  setupMessages.push(`Using Spreadsheet: "${activeSSLeads.getName()}".`);

  try {
    // --- Step 1: Get/Create Leads Sheet Tab & Format ---
    Logger.log(`[${FUNC_NAME} INFO] Setting up sheet: "${LEADS_SHEET_TAB_NAME}"`);
    let leadsSheet = activeSSLeads.getSheetByName(LEADS_SHEET_TAB_NAME);
    if (!leadsSheet) {
      leadsSheet = activeSSLeads.insertSheet(LEADS_SHEET_TAB_NAME);
      Logger.log(`[${FUNC_NAME} INFO] CREATED new tab: "${LEADS_SHEET_TAB_NAME}".`);
    } else { Logger.log(`[${FUNC_NAME} INFO] Found EXISTING tab: "${LEADS_SHEET_TAB_NAME}".`); }
    
    if (!leadsSheet) throw new Error(`Get/Create FAILED for sheet: "${LEADS_SHEET_TAB_NAME}".`);
    
    // Call the generic formatting function from SheetUtils.gs
    // LEADS_SHEET_COLUMN_WIDTHS and LEADS_SHEET_HEADERS are from Config.gs
    if (!setupSheetFormatting(leadsSheet, 
                          LEADS_SHEET_HEADERS, 
                          LEADS_SHEET_COLUMN_WIDTHS,
                          true, // applyBandingFlag = true
                          SpreadsheetApp.BandingTheme.YELLOW // <<< SPECIFY YELLOW THEME (or CYAN, GREEN, GREY if yellow isn't distinct enough)
                         )) { /* throw error */ }
    leadsSheet.setTabColor(BRAND_COLORS.HUNYADI_YELLOW); // From Config.gs
    setupMessages.push(`Sheet "${LEADS_SHEET_TAB_NAME}": Setup OK. Color: Hunyadi Yellow.`);

    // --- Step 2: Gmail Label and Filter Setup ---
    Logger.log(`[${FUNC_NAME} INFO] Setting up Gmail labels & filters for Leads...`);
    getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(100);
    getOrCreateLabel(LEADS_GMAIL_LABEL_PARENT); Utilities.sleep(100);
    const needsProcessLabelObject = getOrCreateLabel(LEADS_GMAIL_LABEL_TO_PROCESS); Utilities.sleep(100);
    const doneProcessLabelObject = getOrCreateLabel(LEADS_GMAIL_LABEL_PROCESSED); Utilities.sleep(100);
    if (!needsProcessLabelObject || !doneProcessLabelObject) throw new Error("Failed to create/verify core Leads Gmail labels.");
    
    let needsProcessLeadLabelId = null;
    const advGmailService = Gmail; // Assumes Advanced Gmail API Service is enabled
    Utilities.sleep(300);
    const labelsListRespLeads = advGmailService.Users.Labels.list('me');
    if (labelsListRespLeads.labels && labelsListRespLeads.labels.length > 0) {
        const targetLabelInfoLeads = labelsListRespLeads.labels.find(l => l.name === LEADS_GMAIL_LABEL_TO_PROCESS);
        if (targetLabelInfoLeads && targetLabelInfoLeads.id) needsProcessLeadLabelId = targetLabelInfoLeads.id;
    }
    if (!needsProcessLeadLabelId) throw new Error(`CRITICAL: Could not get ID for Gmail label "${LEADS_GMAIL_LABEL_TO_PROCESS}".`);
    setupMessages.push(`Leads Labels & 'To Process' ID: OK (${needsProcessLeadLabelId}).`);

    const filterQueryLeadsConst = LEADS_GMAIL_FILTER_QUERY; // from Config.gs
    let leadsFilterExists = false;
    const existingLeadsFiltersResponse = advGmailService.Users.Settings.Filters.list('me'); // Call it once

    // Robust check for the response structure
    if (existingLeadsFiltersResponse && existingLeadsFiltersResponse.filter && Array.isArray(existingLeadsFiltersResponse.filter)) {
        leadsFilterExists = existingLeadsFiltersResponse.filter.some(f => 
            f.criteria?.query === filterQueryLeadsConst && f.action?.addLabelIds?.includes(needsProcessLeadLabelId));
    } else if (existingLeadsFiltersResponse && !existingLeadsFiltersResponse.hasOwnProperty('filter')) {
        Logger.log(`[${FUNC_NAME} INFO] No 'filter' property in Gmail response (user likely has no filters). Assuming filter does not exist.`);
        leadsFilterExists = false; // Explicitly set
    } else {
        // This case handles if existingLeadsFiltersResponse is null or unexpected
        Logger.log(`[${FUNC_NAME} WARN] Unexpected response or null from Gmail Filters.list('me'). Assuming filter does not exist. Response: ${JSON.stringify(existingLeadsFiltersResponse)}`);
        leadsFilterExists = false; // Explicitly set
    }
    
    if (!leadsFilterExists) {
        const leadsFilterResource = { criteria: { query: filterQueryLeadsConst }, action: { addLabelIds: [needsProcessLeadLabelId], removeLabelIds: ['INBOX'] } };
        const createdLeadsFilter = advGmailService.Users.Settings.Filters.create(leadsFilterResource, 'me');
        if (!createdLeadsFilter || !createdLeadsFilter.id) {
             throw new Error(`Gmail filter creation for leads FAILED or no ID. Response: ${JSON.stringify(createdLeadsFilter)}`);
        }
        setupMessages.push("Leads Filter: CREATED.");
    } else { 
        setupMessages.push("Leads Filter: Exists."); 
    }
    
    // --- Step 3: Store Configuration (ScriptProperties for label IDs) ---
    const scriptPropsLeads = PropertiesService.getScriptProperties();
    if (needsProcessLeadLabelId) scriptPropsLeads.setProperty(LEADS_USER_PROPERTY_TO_PROCESS_LABEL_ID, needsProcessLeadLabelId); else leadsModuleSetupSuccess = false;
    
    let doneProcessLeadLabelId = null;
    if(doneProcessLabelObject) { // Get ID for "Processed" label
        const doneLabelInfoLeads = labelsListRespLeads.labels?.find(l => l.name === LEADS_GMAIL_LABEL_PROCESSED);
        if (doneLabelInfoLeads?.id) {
             doneProcessLeadLabelId = doneLabelInfoLeads.id;
             scriptPropsLeads.setProperty(LEADS_USER_PROPERTY_PROCESSED_LABEL_ID, doneProcessLeadLabelId);
        }
    }
    if (!doneProcessLeadLabelId) leadsModuleSetupSuccess = false; // Critical if ID not found/stored for processing
    setupMessages.push(`ScriptProperties for Leads labels updated (ToProcessID: ${needsProcessLeadLabelId}, ProcessedID: ${doneProcessLeadLabelId}).`);

    // --- Step 4: Create Time-Driven Trigger ---
    if (createTimeDrivenTrigger('processJobLeads', 3)) { // Assumed in Triggers.gs, runs every 3 hours
        setupMessages.push("Trigger 'processJobLeads': CREATED.");
    } else { 
        setupMessages.push("Trigger 'processJobLeads': Exists/Verified."); 
    }

  } catch (e) {
    Logger.log(`[${FUNC_NAME} CRITICAL ERROR]: ${e.toString()}\nStack: ${e.stack || 'No stack'}`);
    setupMessages.push(`CRITICAL ERROR: ${e.message}`); leadsModuleSetupSuccess = false;
    try{ SpreadsheetApp.getUi().alert('Leads Module Setup Error', `Error: ${e.message}. Check logs.`, SpreadsheetApp.getUi().ButtonSet.OK); } catch(uiErr){ Logger.log("UI alert fail: " + uiErr.message);}
  }
  Logger.log(`Job Leads Module Setup: ${leadsModuleSetupSuccess ? "SUCCESSFUL" : "ISSUES"}.`);
  return { success: leadsModuleSetupSuccess, messages: setupMessages };
}

/**
 * Processes emails that have been labeled for job lead extraction.
 * This function is designed to be run on a time-driven trigger. It fetches emails,
 * calls the Gemini API to extract job leads, and writes the results to the spreadsheet.
 */
function processJobLeads() {
    const FUNC_NAME = "processJobLeads";
    const SCRIPT_START_TIME = new Date();
    Logger.log(`\n==== ${FUNC_NAME}: STARTING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

    const scriptProperties = PropertiesService.getScriptProperties();
    const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY);

    if (geminiApiKey) {
        Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Retrieved key for "${GEMINI_API_KEY_PROPERTY}" from ScriptProperties. Value (masked): ${geminiApiKey.substring(0, 4)}...${geminiApiKey.substring(geminiApiKey.length - 4)}`);
        if (geminiApiKey.trim() !== "" && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30) {
            Logger.log(`[${FUNC_NAME} INFO] Gemini API Key (ScriptProperties) is valid for Leads processing.`);
        } else {
            Logger.log(`[${FUNC_NAME} WARN] Gemini API Key (ScriptProperties) found but INvalid. callGemini_forJobLeads might use mock data or fail if this key is passed without its own internal placeholder check.`);
            Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Reason: Key failed validation. Length: ${geminiApiKey.length}, StartsWith AIza: ${geminiApiKey.startsWith("AIza")}`);
        }
    } else {
        Logger.log(`[${FUNC_NAME} WARN] Gemini API Key NOT FOUND in ScriptProperties for "${GEMINI_API_KEY_PROPERTY}". callGemini_forJobLeads will use mock/fail.`);
    }

    const { spreadsheet: activeSS } = getOrCreateSpreadsheetAndSheet();
    if (!activeSS) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Main spreadsheet could not be determined. Aborting.`);
        return;
    }

    const { sheet: leadsDataSheet, headerMap: leadsHeaderMap } = getSheetAndHeaderMapping_forLeads(activeSS.getId(), LEADS_SHEET_TAB_NAME);
    if (!leadsDataSheet || !leadsHeaderMap || Object.keys(leadsHeaderMap).length === 0) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Leads sheet "${LEADS_SHEET_TAB_NAME}" or headers not found/mapped in SS ID ${activeSS.getId()}. Aborting.`);
        return;
    }
    Logger.log(`[${FUNC_NAME} INFO] Processing leads against: "${activeSS.getName()}", Leads Tab: "${leadsDataSheet.getName()}"`);

    const needsProcessLabelName = LEADS_GMAIL_LABEL_TO_PROCESS;
    const doneProcessLabelName = LEADS_GMAIL_LABEL_PROCESSED;

    const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
    const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName);
    if (!needsProcessLabel) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Label "${needsProcessLabelName}" not found. Aborting.`);
        return;
    }
    if (!doneProcessLabel) {
        Logger.log(`[${FUNC_NAME} WARN] Label "${doneProcessLabelName}" not found. Processed leads will only be unlabelled from 'To Process'.`);
    }

    const processedLeadEmailIds = getProcessedEmailIdsFromSheet_forLeads(leadsDataSheet, leadsHeaderMap);
    Logger.log(`[${FUNC_NAME} INFO] Preloaded ${processedLeadEmailIds.size} email IDs already processed for leads.`);

    const LEADS_THREAD_LIMIT = 10;
    const LEADS_MESSAGE_LIMIT_PER_RUN = 15;
    let messagesProcessedThisRunCounter = 0;
    const leadThreadsToProcess = needsProcessLabel.getThreads(0, LEADS_THREAD_LIMIT);
    Logger.log(`[${FUNC_NAME} INFO] Found ${leadThreadsToProcess.length} threads in "${needsProcessLabelName}".`);

    const newJobs = [];

    for (const thread of leadThreadsToProcess) {
        if (messagesProcessedThisRunCounter >= LEADS_MESSAGE_LIMIT_PER_RUN) {
            Logger.log(`[${FUNC_NAME} INFO] Message processing limit (${LEADS_MESSAGE_LIMIT_PER_RUN}) reached for this run.`);
            break;
        }
        const scriptRunTimeSeconds = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
        if (scriptRunTimeSeconds > 320) {
            Logger.log(`[${FUNC_NAME} WARN] Execution time limit (${scriptRunTimeSeconds}s) approaching. Stopping further thread processing.`);
            break;
        }

        const messagesInThread = thread.getMessages();
        let threadContainedAtLeastOneNewMessage = false;
        let allNewMessagesInThisThreadProcessedSuccessfully = true;

        for (const message of messagesInThread) {
            if (messagesProcessedThisRunCounter >= LEADS_MESSAGE_LIMIT_PER_RUN) break;

            const msgId = message.getId();
            if (processedLeadEmailIds.has(msgId)) {
                if (DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Msg ID ${msgId} in thread ${thread.getId()} already processed. Skipping.`);
                continue;
            }

            threadContainedAtLeastOneNewMessage = true;
            Logger.log(`\n--- [${FUNC_NAME}] Processing NEW Lead Msg ID: ${msgId}, Thread: ${thread.getId()}, Subject: "${message.getSubject()}" ---`);
            messagesProcessedThisRunCounter++;
            let currentMessageHandledNoErrors = false;

            try {
                let emailBody = message.getPlainBody();
                if (typeof emailBody !== 'string' || emailBody.trim() === "") {
                    Logger.log(`[${FUNC_NAME} WARN] Msg ${msgId}: Body is empty or not a string. Skipping AI call for this message.`);
                    currentMessageHandledNoErrors = true;
                    processedLeadEmailIds.add(msgId);
                    continue;
                }

                const geminiApiResponse = callGemini_forJobLeads(emailBody, geminiApiKey);

                if (geminiApiResponse && geminiApiResponse.success) {
                    const extractedJobsArray = parseGeminiResponse_forJobLeads(geminiApiResponse.data);
                    if (extractedJobsArray && extractedJobsArray.length > 0) {
                        Logger.log(`[${FUNC_NAME} INFO] Gemini extracted ${extractedJobsArray.length} job(s) from msg ${msgId}.`);
                        let atLeastOneValidJobWrittenThisMessage = false;
                        for (const jobData of extractedJobsArray) {
                            if (jobData && jobData.jobTitle && String(jobData.jobTitle).toLowerCase() !== 'n/a' && String(jobData.jobTitle).toLowerCase() !== 'error') {
                                jobData.dateAdded = message.getDate();
                                jobData.sourceEmailSubject = message.getSubject().substring(0, 500);
                                jobData.sourceEmailId = msgId;
                                jobData.status = "New";
                                jobData.processedTimestamp = new Date();
                                newJobs.push(jobData);
                                atLeastOneValidJobWrittenThisMessage = true;
                            } else {
                                if (DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Job from msg ${msgId} was N/A/error or missing title. Skipping sheet write: ${JSON.stringify(jobData)}`);
                            }
                        }
                        if (atLeastOneValidJobWrittenThisMessage) currentMessageHandledNoErrors = true;
                        else {
                            Logger.log(`[${FUNC_NAME} INFO] Msg ${msgId}: Gemini success, and parsed jobs, but no valid/writable jobs after filtering. Considered handled.`);
                            currentMessageHandledNoErrors = true;
                        }
                    } else {
                        Logger.log(`[${FUNC_NAME} INFO] Msg ${msgId}: Gemini API call success, but parsing response yielded no distinct job listings.`);
                        currentMessageHandledNoErrors = true;
                    }
                } else {
                    Logger.log(`[${FUNC_NAME} ERROR] Gemini API FAILED for msg ${msgId}. Error: ${geminiApiResponse ? geminiApiResponse.error : 'Null or unexpected API response object'}`);
                    writeErrorEntryToSheet_forLeads(leadsDataSheet, message, "Gemini API Call Fail (Leads)", geminiApiResponse ? String(geminiApiResponse.error).substring(0, 500) : "Unknown API response error", leadsHeaderMap);
                    allNewMessagesInThisThreadProcessedSuccessfully = false;
                }
            } catch (e) {
                Logger.log(`[${FUNC_NAME} SCRIPT ERROR] Exception processing Msg ${msgId}: ${e.toString()}\nStack: ${e.stack}`);
                writeErrorEntryToSheet_forLeads(leadsDataSheet, message, "Script Error (Leads)", String(e.toString()).substring(0, 500), leadsHeaderMap);
                allNewMessagesInThisThreadProcessedSuccessfully = false;
            }

            if (currentMessageHandledNoErrors) {
                processedLeadEmailIds.add(msgId);
            }
            Utilities.sleep(1500 + Math.floor(Math.random() * 1000));
        }

        if (threadContainedAtLeastOneNewMessage) {
            if (allNewMessagesInThisThreadProcessedSuccessfully) {
                if (doneProcessLabel) {
                    try {
                        thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
                        Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} successfully processed & moved to "${doneProcessLabelName}".`);
                    } catch (eRelabel) {
                        Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} relabel (success case) error: ${eRelabel.message}`);
                    }
                } else {
                    try {
                        thread.removeLabel(needsProcessLabel);
                        Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} processed. Removed from "${needsProcessLabelName}", but "Done" label object is missing.`);
                    } catch (eRemOnly) {
                        Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} removeLabel error: ${eRemOnly.message}`);
                    }
                }
            } else {
                Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} contained new messages but encountered errors during processing. NOT moved from "${needsProcessLabelName}". Will be re-attempted.`);
            }
        } else if (messagesInThread.length > 0) {
            Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} contained only previously processed messages. Ensuring it's labeled correctly.`);
            if (doneProcessLabel && !thread.getLabels().map(l => l.getName()).includes(doneProcessLabelName)) {
                try {
                    thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
                } catch (eOldDone) {
                }
            } else if (!doneProcessLabel) {
                try {
                    thread.removeLabel(needsProcessLabel);
                } catch (eOldRem) {
                }
            }
        } else if (messagesInThread.length === 0) {
            Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} was empty. Removing from "${needsProcessLabelName}".`);
            try {
                thread.removeLabel(needsProcessLabel);
            } catch (eEmptyThread) { /*Minor*/
            }
        }
        Utilities.sleep(500);
    }

    if (newJobs.length > 0) {
        const rows = newJobs.map(job => {
            const row = [];
            for (const header of LEADS_SHEET_HEADERS) {
                row.push(job[header] || "");
            }
            return row;
        });
        leadsDataSheet.getRange(leadsDataSheet.getLastRow() + 1, 1, rows.length, LEADS_SHEET_HEADERS.length).setValues(rows);
        Logger.log(`[${FUNC_NAME} INFO] Batch appended ${newJobs.length} new jobs to the sheet.`);
    }

    Logger.log(`\n==== ${FUNC_NAME}: FINISHED (${new Date().toLocaleString()}) === Messages Attempted This Run: ${messagesProcessedThisRunCounter}. Total Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000}s ====`);
}
