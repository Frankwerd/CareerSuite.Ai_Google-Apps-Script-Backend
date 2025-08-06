// --- SCRIPT-WIDE CONSTANTS ---
const GEMINI_API_KEY_PROPERTY = 'geminiApiKey';
const SCRIPT_PROPERTY_NEEDS_PROCESS_LABEL_NAME = 'gmailNeedsProcessLabelName';
const SCRIPT_PROPERTY_NEEDS_PROCESS_LABEL_ID = 'gmailNeedsProcessLabelId';
const SCRIPT_PROPERTY_DONE_PROCESS_LABEL_NAME = 'gmailDoneProcessLabelName';
const SCRIPT_PROPERTY_DONE_PROCESS_LABEL_ID = 'gmailDoneProcessLabelId';
const SCRIPT_PROPERTY_SPREADSHEET_ID = 'lastCreatedJobTrackerSpreadsheetId';
const TARGET_SHEET_NAME = 'Potential Job Leads';
let COLUMN_MAPPING = {}; // Populated dynamically

/**
 * Creates a new Spreadsheet file & sets up Gmail labels/filters,
 * applies sheet styling, and creates a daily trigger for processing.
 * Designed to be run from the Apps Script editor.
 */
function runInitialSetup_createNewSpreadsheetWithJobLabels() {
  let ui;
  let proceed = true; 
  try {
    ui = SpreadsheetApp.getUi(); 
    const confirmation_response = ui.alert(
      'Confirm Full Setup', // Updated dialog title
      'This will:\n' +
      '1. Create a new Google Spreadsheet for job leads with styling.\n' +
      '2. Create Gmail labels ("Job Application Potential", ".../NeedsProcess", ".../DoneProcess").\n' +
      '3. Create a Gmail filter for "job alert" emails.\n' +
      '4. Set up a daily trigger to automatically process new job leads.\n\n' + // Added info about trigger
      'Do you want to proceed?',
      ui.ButtonSet.YES_NO
    );
    if (confirmation_response !== ui.Button.YES) {
      Logger.log('User cancelled operation via UI.');
      if (ui) ui.alert('Operation Cancelled', 'Setup was cancelled.', ui.ButtonSet.OK);
      proceed = false;
    }
  } catch (e) {
    Logger.log('Initial Setup: UI for confirmation failed, proceeding with setup. Error: ' + e.message);
  }

  if (!proceed) return;

  try {
    Logger.log('Starting runInitialSetup_createNewSpreadsheetWithJobLabels...');

    // --- Step 1: Create Spreadsheet ---
    const newSpreadsheetName = 'Job Leads Tracker - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);
    const newSpreadsheetId = newSpreadsheet.getId();
    const newSpreadsheetUrl = newSpreadsheet.getUrl();
    Logger.log(`NEW SPREADSHEET: Name: ${newSpreadsheetName}, ID: ${newSpreadsheetId}, URL: ${newSpreadsheetUrl}`);

    // --- Step 2: Setup Sheet with Headers and Styling ---
    const leadsSheetName = TARGET_SHEET_NAME; // Assumes TARGET_SHEET_NAME is a global const
    let defaultSheet = newSpreadsheet.getSheetByName('Sheet1');
    let leadsSheet = defaultSheet ? defaultSheet.setName(leadsSheetName) : newSpreadsheet.insertSheet(leadsSheetName);
    
    const headers = [
        "Date Added", "Job Title", "Company", "Location", "Source Email Subject",
        "Link to Job Posting", 
        "Status", 
        "Source Email ID", "Processed Timestamp",
        "Notes" 
    ];
    leadsSheet.appendRow(headers);
    const headerRange = leadsSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    Logger.log(`Headers added and bolded in "${leadsSheetName}".`);

    const numActualColumns = headers.length;
    Utilities.sleep(1500); 

    // 1. Apply Alternating Row Colors (Banding)
    const bandingRange = leadsSheet.getRange(1, 1, leadsSheet.getMaxRows(), numActualColumns);
    try {
        const existingSheetBandings = leadsSheet.getBandings();
        for (let k = 0; k < existingSheetBandings.length; k++) existingSheetBandings[k].remove();
        bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); 
        Logger.log("Applied PREDEFINED LIGHT_GREY alternating row colors (banding).");
    } catch (e) { Logger.log("Error applying PREDEFINED banding: " + e.toString() + "\nStack: " + e.stack); }
    
    // 2. Set Specific Column Widths
    const columnWidths = {
        "Date Added": 100, "Job Title": 220, "Company": 150, "Location": 130,
        "Source Email Subject": 220, "Link to Job Posting": 250, "Status": 80,
        "Source Email ID": 130, "Processed Timestamp": 100, "Notes": 250
    };
    for (let i = 0; i < headers.length; i++) {
        const headerName = headers[i]; const columnIndex = i + 1; 
        try { if (columnWidths[headerName]) leadsSheet.setColumnWidth(columnIndex, columnWidths[headerName]); else leadsSheet.autoResizeColumn(columnIndex); } 
        catch (e) { Logger.log("Error setting/resizing column " + columnIndex + " (" + headerName + "): " + e.toString()); }
    }
    Logger.log("Set column widths.");

    // 3. Hide unused columns to the right
    const totalColumnsInSheet = leadsSheet.getMaxColumns(); 
    if (numActualColumns < totalColumnsInSheet) { 
      try { leadsSheet.hideColumns(numActualColumns + 1, totalColumnsInSheet - numActualColumns); Logger.log(`Hid unused columns from ${numActualColumns + 1} to ${totalColumnsInSheet}.`); } 
      catch (e) { Logger.log("Error hiding columns: " + e.toString() + "\n" + e.stack); }
    } else { Logger.log("No unused columns to hide."); }

    // 4. Set Frozen Rows (Applied last among styling operations)
    try { leadsSheet.setFrozenRows(1); Logger.log("Frozen header row.");} 
    catch (e) { Logger.log("Error freezing rows: " + e.toString() + "\n" + e.stack);}
    Logger.log("Sheet styling applied.");


    // --- Step 3: Gmail Label and Filter Setup ---
    const parentLabelName = 'Job Application Potential';
    const needsProcessLabelName = `${parentLabelName}/NeedsProcess`;
    const doneProcessLabelName = `${parentLabelName}/DoneProcess`;
    let needsProcessLabelId = null, doneProcessLabelId = null;

    if (!GmailApp.getUserLabelByName(parentLabelName)) GmailApp.createLabel(parentLabelName);
    if (!GmailApp.getUserLabelByName(needsProcessLabelName)) GmailApp.createLabel(needsProcessLabelName);
    if (!GmailApp.getUserLabelByName(doneProcessLabelName)) GmailApp.createLabel(doneProcessLabelName);
    Logger.log(`Gmail labels ensured: "${parentLabelName}", "${needsProcessLabelName}", "${doneProcessLabelName}".`);

    const gmailApiService = Gmail; 
    if (!gmailApiService || !gmailApiService.Users || !gmailApiService.Users.Labels) {
        throw new Error("Gmail API Advanced Service is not properly enabled or available.");
    }
    const labelsResponse = gmailApiService.Users.Labels.list('me');
    const allLabels = labelsResponse.labels;
    if (allLabels && allLabels.length > 0) {
        for (let lbl of allLabels) {
            if (lbl.name === needsProcessLabelName) needsProcessLabelId = lbl.id;
            if (lbl.name === doneProcessLabelName) doneProcessLabelId = lbl.id;
            if (needsProcessLabelId && doneProcessLabelId) break;
        }
    }
    if (!needsProcessLabelId) throw new Error(`Could not get ID for label "${needsProcessLabelName}".`);
    Logger.log(`ID for "${needsProcessLabelName}": ${needsProcessLabelId}`);
    if (doneProcessLabelId) Logger.log(`ID for "${doneProcessLabelName}": ${doneProcessLabelId}`);
    else Logger.log(`Warning: Could not get ID for "${doneProcessLabelName}".`);

    const filterQuery = 'subject:("job alert") OR subject:(jobalert)';
    const filterResource = { criteria: { query: filterQuery }, action: { addLabelIds: [needsProcessLabelId], removeLabelIds: ['INBOX'] } };
    try {
      gmailApiService.Users.Settings.Filters.create(filterResource, 'me');
      Logger.log(`Gmail filter instruction sent for "${needsProcessLabelName}".`);
    } catch (e) {
      if (e.message && e.message.includes("Filter already exists")) Logger.log("Gmail filter likely already exists.");
      else throw e;
    }
    Logger.log("Gmail setup completed.");

    // --- Step 4: Store Configuration ---
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty(SCRIPT_PROPERTY_SPREADSHEET_ID, newSpreadsheetId);
    userProps.setProperty('lastCreatedJobTrackerSpreadsheetUrl', newSpreadsheetUrl); 
    userProps.setProperty(SCRIPT_PROPERTY_NEEDS_PROCESS_LABEL_NAME, needsProcessLabelName);
    if (needsProcessLabelId) userProps.setProperty(SCRIPT_PROPERTY_NEEDS_PROCESS_LABEL_ID, needsProcessLabelId);
    userProps.setProperty(SCRIPT_PROPERTY_DONE_PROCESS_LABEL_NAME, doneProcessLabelName);
    if (doneProcessLabelId) userProps.setProperty(SCRIPT_PROPERTY_DONE_PROCESS_LABEL_ID, doneProcessLabelId);
    Logger.log('Configuration saved to UserProperties.');
    
    // --- Step 5: Create Time-Driven Trigger ---
    const triggerFunctionName = 'processJobLeadsWithGemini';
    try {
        // Delete any existing triggers for this function to avoid duplicates
        const existingTriggers = ScriptApp.getProjectTriggers();
        for (let i = 0; i < existingTriggers.length; i++) {
            if (existingTriggers[i].getHandlerFunction() === triggerFunctionName) {
                ScriptApp.deleteTrigger(existingTriggers[i]);
                Logger.log(`Deleted existing trigger for ${triggerFunctionName}.`);
            }
        }

        // Create a new trigger to run daily
        ScriptApp.newTrigger(triggerFunctionName)
            .timeBased()
            .everyDays(1)    // Run once a day
            .atHour(3)       // At 3 AM (script's timezone, adjust as needed: 0-23)
            .create();
        Logger.log(`Successfully created a daily trigger for ${triggerFunctionName} to run around 3 AM.`);
        if (ui) ui.alert('Trigger Created', `A daily trigger has been set for "${triggerFunctionName}" to run around 3 AM.`, ui.ButtonSet.OK);

    } catch (e) {
        Logger.log(`Error creating time-driven trigger for ${triggerFunctionName}: ${e.toString()}\nStack: ${e.stack}`);
        if (ui) ui.alert('Trigger Creation Failed', `Could not set up daily trigger for "${triggerFunctionName}". Error: ${e.message}`, ui.ButtonSet.OK);
    }
    // --- End of Trigger Setup ---
    
    Logger.log('Initial setup including trigger OK. New spreadsheet URL: ' + newSpreadsheetUrl);
    if (ui) ui.alert('Setup Complete!', `Full setup is complete!\n\n- New Spreadsheet: "${newSpreadsheetName}" (URL: ${newSpreadsheetUrl})\n- Gmail configured.\n- Daily processing trigger active.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`Error in initial setup: ${e.toString()}\nStack: ${e.stack || 'No stack available'}`);
    if (ui) ui.alert('Error During Setup', `An error occurred: ${e.message || e}. Check Apps Script logs.`, ui.ButtonSet.OK);
  }
}

// --- API KEY MANAGEMENT FUNCTIONS ---
function setGeminiApiKey_UI() { 
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { Logger.log('setGeminiApiKey_UI: UI context error: ' + e.message); return; }
  const currentKey = PropertiesService.getUserProperties().getProperty(GEMINI_API_KEY_PROPERTY);
  const response = ui.prompt('Gemini API Key', `Enter Google AI Gemini API Key.${currentKey ? ' (Overwrite existing)' : ''}`, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);
      ui.alert('API Key Saved.');
    } else ui.alert('API Key Not Saved (empty).');
  } else ui.alert('API Key Setup Cancelled.');
}
function TEMPORARY_manualSetUserGeminiApiKey() {
  const YOUR_GEMINI_KEY_HERE = 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; 
  const propertyName = GEMINI_API_KEY_PROPERTY;
  if (YOUR_GEMINI_KEY_HERE === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || YOUR_GEMINI_KEY_HERE.trim() === '') {
    const msg = 'ERROR: Gemini API Key not set in TEMPORARY function. Edit the script code first with your key.';
    Logger.log(msg); try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log("UI Alert failed in temporary function: " + e.message); } return;
  }
  PropertiesService.getUserProperties().setProperty(propertyName, YOUR_GEMINI_KEY_HERE);
  const successMsg = `UserProperty "${propertyName}" MANUALLY SET for Gemini. IMPORTANT: Now remove or comment out the TEMPORARY_manualSetUserGeminiApiKey function, or at least clear the hardcoded API key from it for security.`;
  Logger.log(successMsg); try { SpreadsheetApp.getUi().alert(successMsg); } catch(e){ Logger.log("UI Alert failed in temporary function: " + e.message); }
}
function TEMPORARY_showUserProperties() {
  const userProps = PropertiesService.getUserProperties().getProperties();
  let logOutput = "Current UserProperties for this script:\n";
  if (Object.keys(userProps).length === 0) { logOutput += "  (No UserProperties are currently set)\n"; } 
  else { for (const key in userProps) { let value = userProps[key]; if (key.toLowerCase().includes('api') && value && value.length > 10) { value = value.substring(0, 5) + "..." + value.substring(value.length - 4); } logOutput += `  ${key}: ${value}\n`; } }
  Logger.log(logOutput); try { SpreadsheetApp.getUi().alert("User Properties Logged", "Check Apps Script logs (View > Logs or Executions) to see current UserProperties.", SpreadsheetApp.getUi().ButtonSet.OK); } catch(e){ Logger.log("UI Alert failed in temporary function: " + e.message); }
}

// --- GEMINI PROCESSING FUNCTIONS ---
function processJobLeadsWithGemini() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING GEMINI JOB LEAD PROCESSING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  const userProperties = PropertiesService.getUserProperties();
  const geminiApiKey = userProperties.getProperty(GEMINI_API_KEY_PROPERTY);
  const targetSpreadsheetId = userProperties.getProperty(SCRIPT_PROPERTY_SPREADSHEET_ID);
  const needsProcessLabelName = userProperties.getProperty(SCRIPT_PROPERTY_NEEDS_PROCESS_LABEL_NAME);
  const doneProcessLabelName = userProperties.getProperty(SCRIPT_PROPERTY_DONE_PROCESS_LABEL_NAME);

  if (!geminiApiKey) { Logger.log('[FATAL] Gemini API Key not set. Run TEMPORARY_manualSetUserGeminiApiKey.'); return; }
  if (!targetSpreadsheetId) { Logger.log('[FATAL] Target Spreadsheet ID not set. Run initial setup first.'); return; }
  if (!needsProcessLabelName || !doneProcessLabelName) { Logger.log('[FATAL] Gmail label config missing. Run initial setup.'); return; }
  Logger.log(`[INFO] Config OK. API Key: ${geminiApiKey.substring(0,10)}..., SS ID: ${targetSpreadsheetId}`);

  const { sheet: dataSheet, headerMap } = getSheetAndHeaderMapping(targetSpreadsheetId, TARGET_SHEET_NAME);
  if (!dataSheet || Object.keys(headerMap).length === 0) { Logger.log(`[FATAL] Sheet "${TARGET_SHEET_NAME}" / headers not found. Aborting.`); return; }
  COLUMN_MAPPING = headerMap;

  const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
  const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName);
  if (!needsProcessLabel || !doneProcessLabel) { Logger.log(`[FATAL] Processing Gmail labels missing. Aborting.`); return; }

  const processedEmailIds = getProcessedEmailIdsFromSheet(dataSheet);
  Logger.log(`[INFO] Preloaded ${processedEmailIds.size} email IDs (messages fully processed before).`);

  const THREAD_LIMIT = 10, MESSAGE_LIMIT = 15;
  let messagesProcessedThisRun = 0;
  const threads = needsProcessLabel.getThreads(0, THREAD_LIMIT);
  Logger.log(`[INFO] Found ${threads.length} threads in "${needsProcessLabelName}".`);

  for (const thread of threads) {
    if (messagesProcessedThisRun >= MESSAGE_LIMIT) { Logger.log(`[INFO] Msg limit (${MESSAGE_LIMIT}) reached for this run.`); break; }
    if ((new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 300) { Logger.log(`[WARN] Time limit approaching.`); break; }

    const messages = thread.getMessages();
    let threadContainedUnprocessedMessages = false;
    let allMessagesInThreadProcessedSuccessfullyThisRun = true;

    for (const message of messages) {
      if (messagesProcessedThisRun >= MESSAGE_LIMIT) break;
      const msgId = message.getId();

      if (processedEmailIds.has(msgId)) { 
        Logger.log(`[INFO] Skip msg ${msgId} (entire message previously marked as processed).`);
        continue; 
      }
      threadContainedUnprocessedMessages = true;

      Logger.log(`\n--- Processing Msg ${msgId}, Subject: "${message.getSubject()}" ---`);
      messagesProcessedThisRun++;
      let currentMessageHadSuccessfulExtractions = false;
      let currentMessageHadAnyProcessingAttempt = false;

      try {
        let emailBody = message.getPlainBody();
        if (typeof emailBody !== 'string' || emailBody.trim() === "") {
          Logger.log(`[WARN] Msg ${msgId}: Invalid/empty body. Type: ${typeof emailBody}. Skipping Gemini call.`);
          currentMessageHadSuccessfulExtractions = true; 
          currentMessageHadAnyProcessingAttempt = true;
          continue;
        }
        Logger.log(`[DEBUG] Msg ${msgId}: Body len: ${emailBody.length}, Start: "${emailBody.substring(0,100)}"`);
        
        const geminiResponse = callGeminiApiForJobData(emailBody, geminiApiKey);
        currentMessageHadAnyProcessingAttempt = true;
        
        if (geminiResponse && geminiResponse.success) {
          const extractedJobsArray = parseGeminiJobData(geminiResponse.data);

          if (extractedJobsArray && extractedJobsArray.length > 0) {
            Logger.log(`[INFO] Extracted ${extractedJobsArray.length} job listings from message ID ${msgId}.`);
            let allJobsInThisEmailWrittenOrNA = true;
            let atLeastOneGoodJobWritten = false;
            for (const jobData of extractedJobsArray) {
              if (jobData && jobData.jobTitle && jobData.jobTitle.toLowerCase() !== 'n/a') {
                jobData.dateAdded = new Date(); 
                jobData.sourceEmailSubject = message.getSubject(); 
                jobData.sourceEmailId = msgId; 
                jobData.status = "New"; 
                jobData.processedTimestamp = new Date();
                writeJobDataToSheet(dataSheet, jobData);
                Logger.log(`[SUCCESS] Job "${jobData.jobTitle}" from msg ${msgId} written to sheet.`);
                currentMessageHadSuccessfulExtractions = true; 
                atLeastOneGoodJobWritten = true;
              } else {
                Logger.log(`[WARN] A job object from msg ${msgId} was N/A or incomplete. Skipping write: ${JSON.stringify(jobData)}`);
                // If Gemini returns an object that is all N/A, we don't necessarily count it as a failure of the message.
                // We just don't write that specific N/A job.
              }
            }
            if (!atLeastOneGoodJobWritten && extractedJobsArray.length > 0) { // All were N/A
                 Logger.log(`[INFO] All ${extractedJobsArray.length} extracted job objects from msg ${msgId} were N/A. No data written, but considering message processed.`);
                 currentMessageHadSuccessfulExtractions = true; // Processed, but nothing to write for jobs.
            } else if (!currentMessageHadSuccessfulExtractions) { // Should not happen if above is true
                 writeErrorEntryToSheet(dataSheet, message, "Gemini parsing - all jobs N/A", geminiResponse.data);
            }
          } else { 
            Logger.log(`[WARN] Gemini parsing returned no job listings array or it was empty for msg ${msgId}.`); 
            writeErrorEntryToSheet(dataSheet, message, "Gemini parsing - no job array", geminiResponse.data);
          }
        } else { 
          Logger.log(`[ERROR] Gemini API call failed for msg ${msgId}. Details: ${geminiResponse ? geminiResponse.error : 'No response object'}`); 
          writeErrorEntryToSheet(dataSheet, message, "Gemini API call failed", geminiResponse ? geminiResponse.error : "Unknown API error");
        }
      } catch (e) { 
        Logger.log(`[FATAL SCRIPT ERROR] Processing msg ${msgId}: ${e.toString()}\nStack: ${e.stack}`); 
        writeErrorEntryToSheet(dataSheet, message, "Script error during processing", e.toString()); 
        allMessagesInThreadProcessedSuccessfullyThisRun = false; 
      }

      if (currentMessageHadSuccessfulExtractions && currentMessageHadAnyProcessingAttempt) {
        // Message processing attempt was made, and it either yielded good jobs or correctly yielded no jobs (e.g. all N/A from LLM)
      } else if (currentMessageHadAnyProcessingAttempt) { 
        allMessagesInThreadProcessedSuccessfullyThisRun = false;
      }
      Utilities.sleep(2000);
    } 

    if (threadContainedUnprocessedMessages && allMessagesInThreadProcessedSuccessfullyThisRun) {
        thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
        Logger.log(`[INFO] Thread ID ${thread.getId()} successfully processed and moved to "${doneProcessLabelName}".`);
        messages.forEach(m => processedEmailIds.add(m.getId()));
    } else if (threadContainedUnprocessedMessages) { 
        Logger.log(`[WARN] Thread ID ${thread.getId()} had processing issues with one or more messages. Not moved from "${needsProcessLabelName}".`);
    } else if (!threadContainedUnprocessedMessages && messages.length > 0) { 
        Logger.log(`[INFO] Thread ID ${thread.getId()} contained only previously processed messages. Moving to "${doneProcessLabelName}".`);
        thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
    } else { 
        Logger.log(`[INFO] Thread ID ${thread.getId()} was empty. Removing from "${needsProcessLabelName}".`);
        try { thread.removeLabel(needsProcessLabel); } catch(e) { Logger.log("Error removing label from empty thread: " + e);}
    }
  } 
  Logger.log(`\n==== GEMINI PROCESSING FINISHED (${new Date().toLocaleString()}) === Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}

// --- HELPER FUNCTIONS ---
function getSheetAndHeaderMapping(ssId, sheetName) { 
  try {
    const ss = SpreadsheetApp.openById(ssId); const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log(`[ERR] Sheet "${sheetName}" not in SS ID "${ssId}".`); return { sheet: null, headerMap: {} }; }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; const headerMap = {};
    headers.forEach((h, i) => { if (h && h.toString().trim() !== "") headerMap[h.toString().trim()] = i + 1; });
    if (Object.keys(headerMap).length === 0) Logger.log(`[ERR] No headers in "${sheetName}".`);
    else Logger.log(`[INFO] Headers for "${sheetName}": ${JSON.stringify(headerMap)}`);
    return { sheet: sheet, headerMap: headerMap };
  } catch (e) { Logger.log(`[ERR] Opening SS/sheet/headers. ID: ${ssId}, Err: ${e.toString()}`); return { sheet: null, headerMap: {} }; }
}

function getProcessedEmailIdsFromSheet(sheet) {
    const ids = new Set(); const emailIdColHeader = "Source Email ID";
    if (!sheet || !COLUMN_MAPPING[emailIdColHeader]) { Logger.log(`[WARN] No processed IDs: sheet or "${emailIdColHeader}" map missing.`); return ids; }
    const lastR = sheet.getLastRow(); if (lastR < 2) return ids;
    const emailIdColNum = COLUMN_MAPPING[emailIdColHeader];
    const rangeToRead = sheet.getRange(2, emailIdColNum, lastR - 1, 1);
    const emailIdValues = rangeToRead.getValues();
    emailIdValues.forEach(r => { if (r[0] && r[0].toString().trim() !== "") ids.add(r[0].toString().trim()); });
    return ids;
}

function callGeminiApiForJobData(emailBody, apiKey) {
  if (typeof emailBody !== 'string') { Logger.log(`[CRITICAL ERR in callGeminiApi] emailBody not string.`); return { success: false, data: null, error: `emailBody not string.` }; }
  Logger.log(`[API callGeminiApi] Body len: ${emailBody.length}. Start: "${emailBody.substring(0,50)}"`);

  const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  
  if (apiKey === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || apiKey.trim() === '') { 
      Logger.log("[WARN STUB Gemini] Placeholder API Key or empty. Using MOCK response for multiple jobs.");
      if (emailBody.toLowerCase().includes("multiple job listings inside") || emailBody.toLowerCase().includes("software engineer at google")) {
        return {
          success: true, 
          data: { 
            candidates: [{ 
              content: { 
                parts: [{
                  text: JSON.stringify([ 
                    { "jobTitle": "Software Engineer (Mock)", "company": "Tech Alpha (Mock)", "location": "Remote", "linkToJobPosting": "https://example.com/job/alpha"},
                    { "jobTitle": "Product Manager (Mock)", "company": "Innovate Beta (Mock)", "location": "New York, NY", "linkToJobPosting": "https://example.com/job/beta"}
                  ])
                }]
              }
            }]
          }
        };
      }
      return {success: true, data: { candidates: [{ content: { parts: [{text: JSON.stringify([{ "jobTitle": "N/A (Mock Single)", "company": "Some Corp (Mock)", "location": "Remote", "linkToJobPosting": "N/A"}]) }]}}]}, error: null};
  }

  const promptText = `From the following email content, identify each distinct job posting.
For EACH job posting found, extract the following details:
- Job Title
- Company
- Location (city, state, or remote)
- Link (a direct URL to the job application or description, if available)

If a field for a specific job is not found or not applicable, use the string "N/A" as its value.

Format your entire response as a single, valid JSON array where each element of the array is a JSON object representing one job posting.
Each JSON object should have the keys: "jobTitle", "company", "location", "linkToJobPosting".
If no job postings are found, return an empty JSON array: [].
Do not include any text before or after the JSON array. Ensure the JSON is strictly valid.

Example of a single job object:
{
  "jobTitle": "Software Engineer",
  "company": "Tech Corp",
  "location": "Remote",
  "linkToJobPosting": "https://example.com/job/123"
}

Email Content:
---
${emailBody.substring(0, 28000)} 
---
JSON Array Output:`; 

  const payload = {
    contents: [{ parts: [{ "text": promptText }] }],
    generationConfig: {
      temperature: 0.2, 
      maxOutputTokens: 8192 
    }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log(`[API REQUEST to Gemini] URL: ${GEMINI_API_URL.split('key=')[0]}key=YOUR_KEY...`);
    const response = UrlFetchApp.fetch(GEMINI_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      Logger.log(`[API SUCCESS ${responseCode}] Gemini raw response (first 300 chars): ${responseBody.substring(0,300)}...`);
      try {
        return { success: true, data: JSON.parse(responseBody), error: null };
      } catch (jsonParseError) {
        Logger.log(`[API ERROR] Failed to parse Gemini JSON response: ${jsonParseError}. Raw body: ${responseBody}`);
        return { success: false, data: null, error: `Failed to parse API JSON response: ${jsonParseError}. Response: ${responseBody}` };
      }
    } else {
      Logger.log(`[API ERROR ${responseCode}] Gemini error: ${responseBody}`);
      return { success: false, data: null, error: `API Error ${responseCode}: ${responseBody}` };
    }
  } catch (e) {
    Logger.log(`[API CATCH ERROR] Failed to call Gemini API: ${e.toString()}`);
    return { success: false, data: null, error: `Fetch Error: ${e.toString()}` };
  }
}

function parseGeminiJobData(apiResponseData) {
  Logger.log(`[PARSE GeminiJobData] Raw API Data from Gemini (first 300): ${JSON.stringify(apiResponseData).substring(0,300)}...`);
  let jobListings = []; 
  try {
    let jsonStringFromLLM = "";
    if (apiResponseData && apiResponseData.candidates && apiResponseData.candidates.length > 0 &&
        apiResponseData.candidates[0].content && apiResponseData.candidates[0].content.parts &&
        apiResponseData.candidates[0].content.parts.length > 0 &&
        typeof apiResponseData.candidates[0].content.parts[0].text === 'string') {
        jsonStringFromLLM = apiResponseData.candidates[0].content.parts[0].text.trim();
        if (jsonStringFromLLM.startsWith("```json")) jsonStringFromLLM = jsonStringFromLLM.substring(7);
        if (jsonStringFromLLM.startsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(3);
        if (jsonStringFromLLM.endsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(0, jsonStringFromLLM.length - 3);
        jsonStringFromLLM = jsonStringFromLLM.trim();
        Logger.log(`[PARSE GeminiJobData] Cleaned JSON String from LLM (first 500 chars): ${jsonStringFromLLM.substring(0,500)}...`);
        Logger.log(`[PARSE GeminiJobData] FULL Cleaned JSON String from LLM: ${jsonStringFromLLM}`);
    } else {
        Logger.log(`[WARN PARSE GeminiJobData] No parsable content string found in Gemini response structure.`);
        if (apiResponseData && apiResponseData.promptFeedback && apiResponseData.promptFeedback.blockReason) {
            Logger.log(`[WARN PARSE GeminiJobData] Prompt Feedback Block Reason: ${apiResponseData.promptFeedback.blockReason}`);
        }
        return jobListings;
    }

    try {
      const parsedData = JSON.parse(jsonStringFromLLM);
      if (Array.isArray(parsedData)) {
        parsedData.forEach(job => {
          jobListings.push({
            jobTitle: job.jobTitle || "N/A", company: job.company || "N/A", location: job.location || "N/A",
            linkToJobPosting: job.linkToJobPosting || "N/A"
            // Removed Key Skills, Summary
          });
        });
        Logger.log(`[PARSE GeminiJobData] Successfully parsed ${jobListings.length} job objects from JSON string.`);
      } else {
         Logger.log(`[WARN PARSE GeminiJobData] LLM output was not a JSON array. Output (first 200): ${jsonStringFromLLM.substring(0,200)}`);
         if (typeof parsedData === 'object' && parsedData !== null && (parsedData.jobTitle || parsedData.company) ) { 
             jobListings.push({ 
                jobTitle: parsedData.jobTitle || "N/A", company: parsedData.company || "N/A", location: parsedData.location || "N/A",
                linkToJobPosting: parsedData.linkToJobPosting || "N/A"
             });
             Logger.log(`[PARSE GeminiJobData] Fallback: Parsed as a single job object.`);
        }
      }
    } catch (jsonError) {
      Logger.log(`[ERROR PARSE GeminiJobData] Failed to parse JSON string from LLM: ${jsonError}. Problematic String (first 500 and last 500 if long):`);
      if (jsonStringFromLLM.length > 1000) { Logger.log(`START: ${jsonStringFromLLM.substring(0,500)}`); Logger.log(`END: ${jsonStringFromLLM.substring(jsonStringFromLLM.length - 500)}`); } 
      else { Logger.log(`FULL: ${jsonStringFromLLM}`); }
    }    
    return jobListings;
  } catch (e) {
    Logger.log(`[ERROR PARSE GeminiJobData] Outer error during parsing: ${e.toString()}. Data: ${JSON.stringify(apiResponseData).substring(0,300)}`);
    return jobListings;
  }
}

function writeJobDataToSheet(sheet, jobData) {
  Logger.log(`[WRITING] To sheet: "${jobData.jobTitle}" @ "${jobData.company}"`);
  if (!sheet || Object.keys(COLUMN_MAPPING).length === 0) { Logger.log("[ERR] No write: Sheet/COLUMN_MAPPING invalid."); return; }
  const numCols = Math.max(...Object.values(COLUMN_MAPPING), 1); 
  const newRow = new Array(numCols).fill("");
  const getData = (p, d = "") => jobData[p] !== undefined && jobData[p] !== null ? jobData[p] : d;

  if (COLUMN_MAPPING["Date Added"]) newRow[COLUMN_MAPPING["Date Added"]-1] = getData('dateAdded', new Date());
  if (COLUMN_MAPPING["Job Title"]) newRow[COLUMN_MAPPING["Job Title"]-1] = getData('jobTitle', "N/A");
  if (COLUMN_MAPPING["Company"]) newRow[COLUMN_MAPPING["Company"]-1] = getData('company', "N/A");
  if (COLUMN_MAPPING["Location"]) newRow[COLUMN_MAPPING["Location"]-1] = getData('location', "N/A");
  if (COLUMN_MAPPING["Source Email Subject"]) newRow[COLUMN_MAPPING["Source Email Subject"]-1] = getData('sourceEmailSubject');
  if (COLUMN_MAPPING["Link to Job Posting"]) newRow[COLUMN_MAPPING["Link to Job Posting"]-1] = getData('linkToJobPosting', "N/A");
  if (COLUMN_MAPPING["Status"]) newRow[COLUMN_MAPPING["Status"]-1] = getData('status', "New");
  if (COLUMN_MAPPING["Source Email ID"]) newRow[COLUMN_MAPPING["Source Email ID"]-1] = getData('sourceEmailId');
  if (COLUMN_MAPPING["Processed Timestamp"]) newRow[COLUMN_MAPPING["Processed Timestamp"]-1] = getData('processedTimestamp', new Date());
  if (COLUMN_MAPPING["Notes"]) newRow[COLUMN_MAPPING["Notes"]-1] = getData('notes'); // Added back
  
  sheet.appendRow(newRow); 
  Logger.log(`[SUCCESS WRITING] Appended: "${jobData.jobTitle}"`);
}

function writeErrorEntryToSheet(sheet, message, errorType, details) {
    const detailsString = typeof details === 'string' ? details.substring(0, 1000) : JSON.stringify(details).substring(0, 1000);
    Logger.log(`[INFO] Writing error entry for msg ${message.getId()}: ${errorType}. Details: ${detailsString}`);
    
    const errorData = { 
        dateAdded: new Date(), 
        jobTitle: "PROCESSING ERROR", 
        company: errorType.substring(0,250), 
        location: "N/A", 
        sourceEmailSubject: message.getSubject(), 
        linkToJobPosting: "N/A", 
        status: "Error", 
        notes: `Failed: ${errorType}. Details: ${detailsString}`, // Using Notes for details
        sourceEmailId: message.getId(), 
        processedTimestamp: new Date() 
    };
    writeJobDataToSheet(sheet, errorData);
}
