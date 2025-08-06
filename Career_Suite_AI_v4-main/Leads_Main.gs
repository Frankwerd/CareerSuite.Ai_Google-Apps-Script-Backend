// File: Leads_Main.gs
// Description: Contains the primary functions for the Job Leads Tracker module,
// including initial setup of the leads sheet/labels/filters and the
// ongoing processing of job lead emails.

/**
 * Sets up the Job Leads Tracker module:
 * - Ensures the "Potential Job Leads" sheet exists in the main spreadsheet and formats it.
 * - Creates necessary Gmail labels ("Master Job Manager/Job Application Potential/*").
 * - Creates a Gmail filter for "job alert" emails to apply the "NeedsProcess" label.
 * - Sets up a daily trigger for processing new job leads.
 * Designed to be run manually from the Apps Script editor once.
 */

function runInitialSetup_JobLeadsModule() {
  Logger.log("runInitialSetup_JobLeadsModule: User initiated from Apps Script editor. Proceeding with setup.");
  Logger.log(
      'This will:\n' +
      '1. Create/Verify the "Potential Job Leads" tab in your main spreadsheet (using Tab Name: "' + LEADS_SHEET_TAB_NAME + '").\n' +
      '2. Style the "Potential Job Leads" tab with headers and formatting.\n' +
      '3. Create Gmail labels: "' + MASTER_GMAIL_LABEL_PARENT + '", "' + LEADS_GMAIL_LABEL_PARENT + '", "' + LEADS_GMAIL_LABEL_NEEDS_PROCESS + '", and "' + LEADS_GMAIL_LABEL_DONE_PROCESS + '".\n' +
      '4. Create a Gmail filter for query: "' + LEADS_GMAIL_FILTER_QUERY + '" emails.\n' +
      '5. Set up a daily trigger to automatically process new job leads (function: "processJobLeads").'
  );


  let ui = null; // Define ui, it might be used for later, less critical alerts.
  try { ui = SpreadsheetApp.getUi(); } catch(e) { /* No UI context, that's fine here */ }


  try {
    Logger.log('Starting runInitialSetup_JobLeadsModule core logic...');

    // --- Step 1: Get Main Spreadsheet & Leads Sheet Tab ---
    const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From main SheetUtils.gs
    if (!mainSpreadsheet) {
      throw new Error("Could not get or create the main spreadsheet. Ensure FIXED_SPREADSHEET_ID or TARGET_SPREADSHEET_FILENAME in Config.gs is correct.");
    }
    Logger.log(`MAIN SPREADSHEET: Using "${mainSpreadsheet.getName()}", ID: ${mainSpreadsheet.getId()}, URL: ${mainSpreadsheet.getUrl()}`);

    let leadsSheet = mainSpreadsheet.getSheetByName(LEADS_SHEET_TAB_NAME); // LEADS_SHEET_TAB_NAME from Config.gs
    if (!leadsSheet) {
      leadsSheet = mainSpreadsheet.insertSheet(LEADS_SHEET_TAB_NAME);
      Logger.log(`Created new tab: "${LEADS_SHEET_TAB_NAME}" in spreadsheet "${mainSpreadsheet.getName()}".`);
    } else {
      Logger.log(`Found existing tab: "${LEADS_SHEET_TAB_NAME}" in spreadsheet "${mainSpreadsheet.getName()}".`);
    }
    try { mainSpreadsheet.setActiveSheet(leadsSheet); } // Make it active for subsequent operations
    catch(e) { Logger.log("Could not set active sheet, but proceeding. Error: " + e.message); }


    // --- Step 2: Setup Leads Sheet with Headers and Styling ---
    if (leadsSheet.getLastRow() === 0 || (leadsSheet.getLastRow() === 1 && leadsSheet.getLastColumn() <=1 && leadsSheet.getRange(1,1).isBlank())) {
        leadsSheet.clearContents();
        leadsSheet.clearFormats();
        Logger.log(`Cleared contents and formats for "${LEADS_SHEET_TAB_NAME}" as it appeared new/empty.`);
    } else {
        Logger.log(`"${LEADS_SHEET_TAB_NAME}" has existing content. Headers will be verified/added if missing. Styling will be applied.`);
    }

    if (leadsSheet.getLastRow() < 1 || leadsSheet.getRange(1,1).isBlank()) {
        leadsSheet.appendRow(LEADS_SHEET_HEADERS); // From Config.gs
        Logger.log(`Headers added to "${LEADS_SHEET_TAB_NAME}".`);
    } else {
        Logger.log(`Headers appear to exist in "${LEADS_SHEET_TAB_NAME}". Skipping appendRow.`);
    }

    const headerRange = leadsSheet.getRange(1, 1, 1, LEADS_SHEET_HEADERS.length);
    headerRange.setFontWeight("bold").setHorizontalAlignment("center");
    leadsSheet.setFrozenRows(1);
    Logger.log("Header row styled and frozen for leads sheet.");

    const bandingRange = leadsSheet.getRange(1, 1, leadsSheet.getMaxRows(), LEADS_SHEET_HEADERS.length);
    try {
        const existingSheetBandings = leadsSheet.getBandings();
        for (let k = 0; k < existingSheetBandings.length; k++) existingSheetBandings[k].remove();
        bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
        Logger.log("Applied LIGHT_GREY alternating row colors (banding) to leads sheet.");
    } catch (e) { Logger.log("Error applying banding to leads sheet: " + e.toString()); }

    const columnWidthsLeads = { // Defined in original script, kept local here
        "Date Added": 100, "Job Title": 220, "Company": 150, "Location": 130,
        "Source Email Subject": 220, "Link to Job Posting": 250, "Status": 80,
        "Source Email ID": 130, "Processed Timestamp": 100, "Notes": 250
    };
    for (let i = 0; i < LEADS_SHEET_HEADERS.length; i++) {
        const headerName = LEADS_SHEET_HEADERS[i]; const columnIndex = i + 1;
        try {
            if (columnWidthsLeads[headerName]) {
                leadsSheet.setColumnWidth(columnIndex, columnWidthsLeads[headerName]);
            }
        }
        catch (e) { Logger.log("Error setting column " + columnIndex + " (" + headerName + ") width for leads sheet: " + e.toString()); }
    }
    Logger.log("Set column widths for leads sheet.");

    const totalColumnsInLeadsSheet = leadsSheet.getMaxColumns();
    if (LEADS_SHEET_HEADERS.length < totalColumnsInLeadsSheet) {
      try {
        leadsSheet.hideColumns(LEADS_SHEET_HEADERS.length + 1, totalColumnsInLeadsSheet - LEADS_SHEET_HEADERS.length);
        Logger.log(`Hid unused columns in leads sheet from column ${LEADS_SHEET_HEADERS.length + 1}.`);
      }
      catch (e) { Logger.log("Error hiding columns in leads sheet: " + e.toString()); }
    }
    Logger.log("Leads sheet styling applied.");

    // --- Step 3: Gmail Label and Filter Setup ---
    Logger.log(`[LEADS_SETUP DEBUG] Ensuring parent labels exist for leads module...`);
    // Call getOrCreateLabel (from GmailUtils.gs) to ensure labels are present.
    // These calls also have internal logging (including the GMAIL_UTIL DEBUG RETURN CHECK).
    getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); 
    Utilities.sleep(500); // Pause after master parent
    getOrCreateLabel(LEADS_GMAIL_LABEL_PARENT);  // From Config.gs
    Utilities.sleep(500); // Pause after leads parent

    const needsProcessLabelNameConst = LEADS_GMAIL_LABEL_NEEDS_PROCESS; // From Config.gs
    const doneProcessLabelNameConst = LEADS_GMAIL_LABEL_DONE_PROCESS;   // From Config.gs

    // Ensure the specific processing labels exist.
    // We don't need to capture their objects here if we use Advanced Service for ID.
    getOrCreateLabel(needsProcessLabelNameConst);
    Utilities.sleep(500); // Pause after creating/verifying "NeedsProcess"
    getOrCreateLabel(doneProcessLabelNameConst);
    Utilities.sleep(500); // Pause after creating/verifying "DoneProcess"
    Logger.log(`[LEADS_SETUP INFO] Called getOrCreateLabel for all specific leads module labels ("${needsProcessLabelNameConst}", "${doneProcessLabelNameConst}").`);

    let needsProcessLeadLabelId = null;

    // --- Get Label ID using Advanced Gmail Service ---
    Logger.log(`[LEADS_SETUP INFO] Attempting to get Label ID for "${needsProcessLabelNameConst}" using Advanced Gmail Service.`);
    try {
        // Ensure Gmail API Advanced Service is enabled in your project (Services + > Gmail API > Add)
        const advancedGmailService = Gmail; // This is how you reference the advanced service
        if (!advancedGmailService || !advancedGmailService.Users || !advancedGmailService.Users.Labels) {
            const advServiceErrorMsg = "Gmail API Advanced Service (Gmail) is not available or not properly enabled. Please enable it via Services + > Gmail API > Add.";
            Logger.log(`[LEADS_SETUP CRITICAL ERROR] ${advServiceErrorMsg}`);
            throw new Error(advServiceErrorMsg); // This will be caught by the outer try-catch of runInitialSetup_JobLeadsModule
        }
        
        const labelsListResponse = advancedGmailService.Users.Labels.list('me');
        
        if (labelsListResponse && labelsListResponse.labels && labelsListResponse.labels.length > 0) {
            const targetLabelInfo = labelsListResponse.labels.find(l => l.name === needsProcessLabelNameConst);
            if (targetLabelInfo && targetLabelInfo.id) {
                needsProcessLeadLabelId = targetLabelInfo.id;
                Logger.log(`[LEADS_SETUP INFO] Successfully retrieved Label ID via Advanced Service: "${needsProcessLeadLabelId}" for label "${needsProcessLabelNameConst}".`);
            } else {
                Logger.log(`[LEADS_SETUP WARN] Label "${needsProcessLabelNameConst}" not found in list from Advanced Gmail Service. Ensure it was created by getOrCreateLabel.`);
                // This could happen if getOrCreateLabel failed silently or there's a significant propagation delay.
            }
        } else {
            Logger.log(`[LEADS_SETUP WARN] No labels returned by Advanced Gmail Service list for user 'me', or labels array is empty. Response: ${JSON.stringify(labelsListResponse)}`);
        }
    } catch (e) {
        Logger.log(`[LEADS_SETUP ERROR] Error using Advanced Gmail Service to get label ID for "${needsProcessLabelNameConst}": ${e.message}\nStack: ${e.stack}`);
        // If Advanced Service itself fails (e.g., not enabled), needsProcessLeadLabelId will remain null.
    }
    // --- End Get Label ID using Advanced Gmail Service ---

    if (!needsProcessLeadLabelId) {
         const errorMsg = `CRITICAL: Could not obtain ID for Gmail label "${needsProcessLabelNameConst}" using Advanced Gmail Service. Filter creation will fail. Please check Gmail manually for the label, and ensure the Advanced Gmail Service is enabled and functioning.`;
         Logger.log(errorMsg);
         if (ui) { // ui was defined at the start of runInitialSetup_JobLeadsModule
            ui.alert("Label ID Error (Advanced Service)", errorMsg, ui.ButtonSet.OK);
         }
         throw new Error(errorMsg); // Stop further execution of this setup if ID is not found
    }
    Logger.log(`[LEADS_SETUP INFO] Using Gmail Label ID "${needsProcessLeadLabelId}" (obtained via Advanced Service) for "${needsProcessLabelNameConst}" for filter creation.`);

    // --- Filter Creation Logic (using needsProcessLeadLabelId from Advanced Service) ---
    Logger.log(`[LEADS_SETUP INFO] Proceeding to create/verify Gmail filter...`);
    try {
        const gmailApiServiceForFilter = Gmail; // Use the same advanced service reference
        let filterExists = false;
        const existingFiltersResponse = gmailApiServiceForFilter.Users.Settings.Filters.list('me');
        const existingFiltersList = existingFiltersResponse.filter; // The actual array of filters

        if (existingFiltersList && existingFiltersList.length > 0) {
            for (const filter of existingFiltersList) {
                if (filter.criteria && filter.criteria.query === LEADS_GMAIL_FILTER_QUERY && // From Config.gs
                    filter.action && filter.action.addLabelIds && filter.action.addLabelIds.includes(needsProcessLeadLabelId)) {
                    filterExists = true;
                    break;
                }
            }
        }

        if (!filterExists) {
            const filterResource = {
                criteria: { query: LEADS_GMAIL_FILTER_QUERY }, // From Config.gs
                action: { addLabelIds: [needsProcessLeadLabelId], removeLabelIds: ['INBOX'] } // Example action
            };
            gmailApiServiceForFilter.Users.Settings.Filters.create(filterResource, 'me');
            Logger.log(`[LEADS_SETUP INFO] Gmail filter CREATED for query "${LEADS_GMAIL_FILTER_QUERY}" to apply label ID "${needsProcessLeadLabelId}".`);
        } else {
            Logger.log(`[LEADS_SETUP INFO] Gmail filter for query "${LEADS_GMAIL_FILTER_QUERY}" and label ID "${needsProcessLeadLabelId}" appears to ALREADY EXIST.`);
        }
    } catch (e) {
        if (e.message && e.message.toLowerCase().includes("filter already exists")) {
            Logger.log(`[LEADS_SETUP WARN] Gmail filter (query: "${LEADS_GMAIL_FILTER_QUERY}") likely already exists (API reported).`);
        } else {
            Logger.log(`[LEADS_SETUP ERROR] Error creating/checking Gmail filter: ${e.toString()}. Ensure Gmail API Advanced Service is enabled and permissions are correct.`);
            if (ui) ui.alert('Filter Creation Issue', `Could not create Gmail filter. Error: ${e.message}. Check logs. Ensure Gmail API Advanced Service is enabled.`, ui.ButtonSet.OK);
            // Depending on criticality, you might re-throw 'e' here to halt the entire setup.
            // For now, it logs and continues to UserProperty/Trigger setup.
        }
    }
    Logger.log("[LEADS_SETUP INFO] Gmail label and filter setup (Step 3) for leads module completed.");

    // In Leads_Main.gs -> runInitialSetup_JobLeadsModule()

// ... (Step 3: Gmail Label and Filter Setup - WHICH IS NOW WORKING FOR NEEDS_PROCESS_ID - is above this)

    // --- Step 4: Store Configuration (UserProperties) ---
    Logger.log(`[LEADS_SETUP INFO] Storing configuration to UserProperties...`);
    const userProps = PropertiesService.getUserProperties();

    // Store NeedsProcessLabel ID (already reliably fetched above using Advanced Service)
    if (needsProcessLeadLabelId) { // needsProcessLeadLabelId was populated using Advanced Service in Step 3
        userProps.setProperty(LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID, needsProcessLeadLabelId);
        // userProps.setProperty(LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_NAME, needsProcessLabelNameConst); // Optional
        Logger.log(`[LEADS_SETUP INFO] Stored LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID: ${needsProcessLeadLabelId}`);
    } else {
        Logger.log(`[LEADS_SETUP WARN] needsProcessLeadLabelId was not available to store in UserProperties.`);
    }

    // Get and Store DoneProcessLabel ID using Advanced Gmail Service
    const doneProcessLabelNameConstForProps = LEADS_GMAIL_LABEL_DONE_PROCESS; // From Config.gs
    let doneProcessLeadLabelId = null;
    try {
        const advancedGmailService = Gmail; // Assuming already checked for availability in Step 3
        const labelsListResponse = advancedGmailService.Users.Labels.list('me');
        if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
            const targetLabelInfo = labelsListResponse.labels.find(l => l.name === doneProcessLabelNameConstForProps);
            if (targetLabelInfo && targetLabelInfo.id) {
                doneProcessLeadLabelId = targetLabelInfo.id;
                userProps.setProperty(LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID, doneProcessLeadLabelId);
                // userProps.setProperty(LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_NAME, doneProcessLabelNameConstForProps); // Optional
                Logger.log(`[LEADS_SETUP INFO] Successfully retrieved and stored LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID: ${doneProcessLeadLabelId} for "${doneProcessLabelNameConstForProps}".`);
            } else {
                Logger.log(`[LEADS_SETUP WARN] Label "${doneProcessLabelNameConstForProps}" not found via Advanced Service for UserProperties storage.`);
            }
        } else {
            Logger.log(`[LEADS_SETUP WARN] No labels returned by Advanced Gmail Service list (for DoneProcessLabel UserProperty).`);
        }
    } catch (e) {
        Logger.log(`[LEADS_SETUP ERROR] Error using Advanced Gmail Service for "${doneProcessLabelNameConstForProps}" ID for UserProperties: ${e.message}`);
    }
    if (!doneProcessLeadLabelId) { // If still not found after trying.
         Logger.log(`[LEADS_SETUP WARN] doneProcessLeadLabelId for "${doneProcessLabelNameConstForProps}" could not be obtained or stored in UserProperties.`);
    }
    Logger.log('[LEADS_SETUP INFO] UserProperties storage attempt for Step 4 finished.');

    // --- Step 5: Create Time-Driven Trigger --- 
    // ... (This part should be fine)

    // --- Step 5: Create Time-Driven Trigger ---
    const triggerFunctionNameForLeads = 'processJobLeads';
    let triggerWasCreated = false;
    try {
        const existingTriggers = ScriptApp.getProjectTriggers();
        let triggerAlreadyExists = false;
        for (let i = 0; i < existingTriggers.length; i++) {
            if (existingTriggers[i].getHandlerFunction() === triggerFunctionNameForLeads) {
                Logger.log(`Deleting existing trigger for ${triggerFunctionNameForLeads} to ensure single instance.`);
                ScriptApp.deleteTrigger(existingTriggers[i]);
                // If you want to be certain only one exists, break after deleting one.
                // Or, if multiple could exist due to past errors, let it loop to delete all.
            }
        }
        ScriptApp.newTrigger(triggerFunctionNameForLeads)
            .timeBased()
            .everyDays(1)
            .atHour(3) // Approx 3 AM in script's timezone
            .create();
        triggerWasCreated = true;
        Logger.log(`Successfully created a new daily trigger for ${triggerFunctionNameForLeads} to run around 3 AM.`);
    } catch (e) {
        Logger.log(`Error creating time-driven trigger for ${triggerFunctionNameForLeads}: ${e.toString()}`);
    }

    Logger.log('Initial setup for Job Leads Module complete.');
    if (ui) { // Only show UI alert if UI context exists
        if (triggerWasCreated) {
            ui.alert('Job Leads Module Setup Complete!', `The "Potential Job Leads" tab and associated Gmail configurations are ready.\nA new daily processing trigger for "${triggerFunctionNameForLeads}" has been created.`, ui.ButtonSet.OK);
        } else {
            ui.alert('Job Leads Module Setup Complete (Trigger Note)', `The "Potential Job Leads" tab and Gmail configurations are ready.\nCould not create/verify trigger for "${triggerFunctionNameForLeads}" (check logs), or it might have existed.`, ui.ButtonSet.OK);
        }
    } else {
        Logger.log("Job Leads Module Setup Complete (No UI Alert as no UI context). Trigger creation logged above.");
    }

  } catch (e) {
    Logger.log(`CRITICAL ERROR in Job Leads Module initial setup: ${e.toString()}\nStack: ${e.stack || 'No stack available'}`);
    if (ui) ui.alert('Error During Leads Setup', `A critical error occurred: ${e.message || e}. Setup may be incomplete. Check Apps Script logs.`, ui.ButtonSet.OK);
  }
}

/**
 * Processes emails labeled for job leads, extracts job information using Gemini,
 * and writes the data to the "Potential Job Leads" sheet.
 * Intended to be run by a time-driven trigger.
 */

function processJobLeads() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING JOB LEAD PROCESSING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  const userProperties = PropertiesService.getUserProperties();
  // Use GEMINI_API_KEY_PROPERTY from Config.gs (shared with main tracker)
  const geminiApiKey = userProperties.getProperty(GEMINI_API_KEY_PROPERTY);

  // Get main spreadsheet ID (where "Potential Job Leads" tab resides)
  // This relies on the main tracker's config for spreadsheet access.
  const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From main SheetUtils.gs

  if (!mainSpreadsheet) {
      Logger.log('[FATAL] Main spreadsheet not found for leads processing. Ensure FIXED_SPREADSHEET_ID or TARGET_SPREADSHEET_FILENAME in Config.gs is correct. Aborting.');
      return;
  }
  const targetSpreadsheetId = mainSpreadsheet.getId(); // Use the ID of the obtained main spreadsheet

  const needsProcessLabelName = LEADS_GMAIL_LABEL_NEEDS_PROCESS; // From Config.gs
  const doneProcessLabelName = LEADS_GMAIL_LABEL_DONE_PROCESS;   // From Config.gs

  if (!geminiApiKey) {
    Logger.log('[FATAL] Gemini API Key not found in UserProperties under key: "' + GEMINI_API_KEY_PROPERTY + '". Aborting job leads processing.');
    return;
  }
  if (!needsProcessLabelName || !doneProcessLabelName) {
    Logger.log('[FATAL] Gmail label configuration for leads is missing in Config.gs. Aborting job leads processing.');
    return;
  }
  Logger.log(`[INFO] Job Leads Config OK. API Key: ${geminiApiKey ? geminiApiKey.substring(0,5) + "..." : "NOT SET"}, Main SS ID: ${targetSpreadsheetId}`);

  // getSheetAndHeaderMapping_forLeads will be defined in Leads_SheetUtils.gs
  const { sheet: dataSheet, headerMap } = getSheetAndHeaderMapping_forLeads(targetSpreadsheetId, LEADS_SHEET_TAB_NAME); // LEADS_SHEET_TAB_NAME from Config.gs
  if (!dataSheet || !headerMap || Object.keys(headerMap).length === 0) {
    Logger.log(`[FATAL] Leads sheet "${LEADS_SHEET_TAB_NAME}" or its headers not found in spreadsheet ID ${targetSpreadsheetId}. Aborting job leads processing.`);
    return;
  }

  const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
  const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName);
  if (!needsProcessLabel) {
      Logger.log(`[FATAL] Gmail label "${needsProcessLabelName}" not found. Aborting job leads processing.`);
      return;
  }
  if (!doneProcessLabel) {
      // This is less critical for processing, but good for completion.
      Logger.log(`[WARN] Gmail label "${doneProcessLabelName}" not found. Processing will continue, but threads may not be re-labeled correctly.`);
  }

  // getProcessedEmailIdsFromSheet_forLeads will be in Leads_SheetUtils.gs

  const processedEmailIds = getProcessedEmailIdsFromSheet_forLeads(dataSheet, headerMap);
  Logger.log(`[INFO] Preloaded ${processedEmailIds.size} email IDs already processed for leads from sheet "${dataSheet.getName()}".`);

  const LEADS_THREAD_LIMIT = 10; // Consider moving to Config.gs
  const LEADS_MESSAGE_LIMIT_PER_RUN = 15; // Consider moving to Config.gs
  let messagesProcessedThisRun = 0;
  const threads = needsProcessLabel.getThreads(0, LEADS_THREAD_LIMIT);
  Logger.log(`[INFO] Found ${threads.length} threads in "${needsProcessLabelName}".`);

  for (const thread of threads) {
    if (messagesProcessedThisRun >= LEADS_MESSAGE_LIMIT_PER_RUN) {
      Logger.log(`[INFO] Leads message limit (${LEADS_MESSAGE_LIMIT_PER_RUN}) reached for this run.`);
      break;
    }
    if ((new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 320) { // Approx 5m20s, adjust as needed
      Logger.log(`[WARN] Time limit approaching for leads processing. Stopping loop.`);
      break;
    }

    const messages = thread.getMessages();
    let threadContainedUnprocessedMessages = false;
    let allMessagesInThreadSuccessfullyProcessedThisRun = true; // Assume success until a message fails

    for (const message of messages) {
      if (messagesProcessedThisRun >= LEADS_MESSAGE_LIMIT_PER_RUN) break;
      const msgId = message.getId();

      if (processedEmailIds.has(msgId)) {
        // Logger.log(`[DEBUG] Msg ${msgId} in thread ${thread.getId()} already fully processed for leads. Skipping.`);
        continue;
      }
      threadContainedUnprocessedMessages = true; // Found at least one message not in our processed ID set

      Logger.log(`\n--- Processing Lead Msg ID: ${msgId}, Subject: "${message.getSubject()}" ---`);
      messagesProcessedThisRun++;
      let currentMessageProcessedSuccessfully = false; // Flag for current message's outcome

      try {
        let emailBody = message.getPlainBody();
        if (typeof emailBody !== 'string' || emailBody.trim() === "") {
          Logger.log(`[WARN] Msg ${msgId}: Invalid or empty body for leads. Skipping Gemini call for this message.`);
          currentMessageProcessedSuccessfully = true; // Successfully determined there's nothing to parse
          // No error entry needed here, it's just an empty email.
          continue;
        }

        // callGemini_forJobLeads and parseGeminiResponse_forJobLeads will be in GeminiService.gs
        const geminiApiResponse = callGemini_forJobLeads(emailBody, geminiApiKey);

        if (geminiApiResponse && geminiApiResponse.success) {
          const extractedJobsArray = parseGeminiResponse_forJobLeads(geminiApiResponse.data); // Pass raw API data

          if (extractedJobsArray && extractedJobsArray.length > 0) {
            Logger.log(`[INFO] Gemini extracted ${extractedJobsArray.length} potential job listings from msg ${msgId}.`);
            let atLeastOneGoodJobWrittenThisMessage = false;
            for (const jobData of extractedJobsArray) {
              if (jobData && jobData.jobTitle && jobData.jobTitle.toLowerCase() !== 'n/a' && jobData.jobTitle.toLowerCase() !== 'error') {
                // Prepare jobData for sheet writing
                jobData.dateAdded = new Date();
                jobData.sourceEmailSubject = message.getSubject();
                jobData.sourceEmailId = msgId;
                jobData.status = "New"; // Default status for new leads
                jobData.processedTimestamp = new Date();
                // writeJobDataToSheet_forLeads will be in Leads_SheetUtils.gs
                writeJobDataToSheet_forLeads(dataSheet, jobData, headerMap);
                atLeastOneGoodJobWrittenThisMessage = true;
              } else {
                Logger.log(`[INFO] A job object from msg ${msgId} was N/A or error. Skipping write for this specific item: ${JSON.stringify(jobData)}`);
              }
            }
            if (atLeastOneGoodJobWrittenThisMessage) {
                currentMessageProcessedSuccessfully = true;
            } else {
                // No good jobs were extracted, but Gemini call was successful.
                // Could be Gemini correctly found no jobs, or all were N/A.
                Logger.log(`[INFO] Msg ${msgId}: Gemini call successful, but no valid job listings written (all N/A or empty array).`);
                currentMessageProcessedSuccessfully = true; // Still consider message processed.
            }
          } else { // Gemini call success, but extractedJobsArray is null or empty
            Logger.log(`[INFO] Msg ${msgId}: Gemini call successful, but parsing yielded no job listings array or it was empty.`);
            // This could be a valid case where the email contains no jobs.
            currentMessageProcessedSuccessfully = true; // Consider the message processed.
          }
        } else { // Gemini API call failed
          Logger.log(`[ERROR] Gemini API call FAILED for msg ${msgId}. Details: ${geminiApiResponse ? geminiApiResponse.error : 'Response object was null'}`);
          // writeErrorEntryToSheet_forLeads will be in Leads_SheetUtils.gs
          writeErrorEntryToSheet_forLeads(dataSheet, message, "Gemini API Call/Parse Failed", geminiApiResponse ? geminiApiResponse.error : "Unknown API error", headerMap);
          allMessagesInThreadSuccessfullyProcessedThisRun = false; // Mark thread as having an issue
        }
      } catch (e) {
        Logger.log(`[FATAL SCRIPT ERROR] Processing msg ${msgId} for leads: ${e.toString()}\nStack: ${e.stack}`);
        // writeErrorEntryToSheet_forLeads will be in Leads_SheetUtils.gs
        writeErrorEntryToSheet_forLeads(dataSheet, message, "Script error during lead processing", e.toString(), headerMap);
        allMessagesInThreadSuccessfullyProcessedThisRun = false; // Mark thread as having an issue
      }

      if (currentMessageProcessedSuccessfully) {
        // Optionally add msgId to a temporary set for this run if you want to ensure labels are only changed if *all* new messages in a thread are fine.
        // For now, relying on allMessagesInThreadSuccessfullyProcessedThisRun
      }
      Utilities.sleep(1500 + Math.floor(Math.random() * 1000)); // Pause between messages
    } // End loop over messages in a thread

    // After processing all messages in a thread (or hitting limit):
    if (threadContainedUnprocessedMessages && allMessagesInThreadSuccessfullyProcessedThisRun) {
      if (doneProcessLabel) { // Check if doneProcessLabel was found
        thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
        Logger.log(`[INFO] Thread ID ${thread.getId()} successfully processed for leads and moved to "${doneProcessLabelName}".`);
        // Add all message IDs from this thread to our main processedEmailIds set to prevent re-processing in future runs.
        messages.forEach(m => processedEmailIds.add(m.getId()));
      } else {
        thread.removeLabel(needsProcessLabel); // Remove from "NeedsProcess" even if "DoneProcess" is missing
        Logger.log(`[WARN] Thread ID ${thread.getId()} processed for leads. Removed from "${needsProcessLabelName}", but "${doneProcessLabelName}" label was not found to apply.`);
      }
    } else if (threadContainedUnprocessedMessages) { // Some messages failed
      Logger.log(`[WARN] Thread ID ${thread.getId()} had processing issues with one or more messages. NOT moved from "${needsProcessLabelName}". Will be retried next run.`);
    } else if (!threadContainedUnprocessedMessages && messages.length > 0) { // Thread had messages, but all were already processed
        if (doneProcessLabel) {
            thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
            Logger.log(`[INFO] Thread ID ${thread.getId()} contained only previously processed lead messages. Ensured it is in "${doneProcessLabelName}".`);
        } else {
            thread.removeLabel(needsProcessLabel);
            Logger.log(`[INFO] Thread ID ${thread.getId()} contained only previously processed lead messages. Removed from "${needsProcessLabelName}".`);
        }
    } else { // Thread was empty or became empty
      Logger.log(`[INFO] Thread ID ${thread.getId()} appears empty or all its messages were skipped. Removing from "${needsProcessLabelName}".`);
      try { thread.removeLabel(needsProcessLabel); }
      catch(e) { Logger.log(`[DEBUG] Minor error removing label from (likely) already unlabelled/empty thread ${thread.getId()}: ${e}`);}
    }
    Utilities.sleep(500); // Pause between threads
  } // End loop over threads

  Logger.log(`\n==== JOB LEAD PROCESSING FINISHED (${new Date().toLocaleString()}) === Total Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}

