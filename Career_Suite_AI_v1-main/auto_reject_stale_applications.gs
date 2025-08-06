/**
 * Script: Auto-Reject Stale Applications
 * Version: 1.3 (Standalone - Template Version)
 * Date: April 29, 2024 (Update as needed)
 *
 * Purpose: This script runs independently to review a specified Google Sheet ('Applications').
 *          It identifies rows where the 'Last Update Date' is older than a defined
 *          threshold (e.g., 7 weeks) and the status is considered 'active' (not final),
 *          automatically changing the status to 'Rejected' (or your configured status).
 *
 * Setup:   1. **CRITICAL:** Update the `SPREADSHEET_ID` constant below with YOUR sheet's ID.
 *          2. Verify `SHEET_NAME`, column numbers (`STATUS_COL`, `LAST_UPDATE_DATE_COL`), and status strings
 *             (`REJECTED_STATUS`, `FINAL_STATUSES`) match YOUR specific Google Sheet setup and preferences.
 *             These MUST align with the sheet structure and any primary email processing script you use.
 *          3. Adjust `WEEKS_THRESHOLD` if desired.
 *          4. Manually create a time-driven trigger (e.g., run weekly) in the
 *             Google Apps Script editor targeting the 'markStaleApplicationsAsRejected' function.
 *             Go to Script Editor -> Triggers (Clock Icon) -> Add Trigger -> Choose function 'markStaleApplicationsAsRejected',
 *             Select event source 'Time-driven', Choose trigger type (e.g., 'Weekly timer').
 *
 * NOTE:    This script is intended to run separately from any main email processing script.
 *          Ensure configuration constants here match your actual sheet layout.
 */

// ============================================================================
// --- USER CONFIGURATION SECTION ---
// --- You MUST modify the values below to match YOUR setup ---
// ============================================================================

// --- Core Setup ---

// !!! REQUIRED: Replace placeholder with the ID of YOUR Google Sheet !!!
// Find the ID in your sheet's URL: docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // <<<--- EDIT THIS LINE

// Name of the sheet tab within your spreadsheet where application data resides
// <<< MUST MATCH the actual tab name in your Google Sheet >>>
const SHEET_NAME = "Applications";                 // <<<--- EDIT THIS if your sheet tab has a different name

// --- Timing & Status ---

// Number of weeks after the 'Last Update Date' before an application is considered stale
// Adjust this value based on your preference (e.g., 6, 8, 10)
const WEEKS_THRESHOLD = 7;                         // <<<--- EDIT THIS if you want a different timeout period

// The status text to assign to applications identified as stale
// <<< MUST MATCH the status string you want to use for automatic rejection >>>
const REJECTED_STATUS = "Rejected"; // <<<--- EDIT THIS if you use a different term (e.g., "Closed - No Response")

// Define statuses that should NOT be automatically overwritten by this script.
// Add any other statuses you consider "final" or want to protect from being changed.
// <<< Ensure these EXACTLY match the status text used in your sheet and potentially your email parser script >>>
const FINAL_STATUSES = new Set([
  REJECTED_STATUS,          // Include the status you set above
  "Offer/Accepted",         // Example final status
  "Withdrawn",              // Example final status
  "Hired",                  // Example final status
  "Offer Declined"          // Example final status
  // Add more strings here if needed, matching your sheet exactly
]);                         // <<<--- EDIT THIS SET to include all your final/protected statuses

// --- Column Configuration (IMPORTANT!) ---

// Column numbers (1 = A, 2 = B, 3 = C, etc.) in YOUR sheet
// <<< MUST MATCH your actual column positions in the Google Sheet >>>

const STATUS_COL = 6;              // Example: Column F contains the Application Status
                                   // <<<--- EDIT THIS NUMBER if your Status column is different

const LAST_UPDATE_DATE_COL = 7;    // Example: Column G contains the 'Last Update Date' timestamp
                                   // <<<--- EDIT THIS NUMBER if your Last Update Date column is different

// ============================================================================
// --- END OF USER CONFIGURATION SECTION ---
// --- Do not modify below unless you understand the script's logic ---
// ============================================================================


/**
 * Main function to find and mark stale applications as rejected.
 * This function should be targeted by a manually configured time-driven trigger.
 */
function markStaleApplicationsAsRejected() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`==== AUTO-REJECT STALE SCRIPT STARTING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  // --- Input Validation ---
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || !SPREADSHEET_ID) {
    Logger.log("[FATAL ERROR] SPREADSHEET_ID has not been set in the script configuration. Please edit the script and replace 'YOUR_SPREADSHEET_ID_HERE' with your actual Google Sheet ID. Aborting.");
    // Optional: Could add SpreadsheetApp.getUi().alert(...) here if running manually, but less useful for triggers.
    return;
  }

  let ss; // Declare Spreadsheet object variable
  try {
    // Use openById for standalone scripts
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
       // This case might be redundant given openById throws on invalid ID, but kept for clarity
       throw new Error(`Could not open Spreadsheet. Check SPREADSHEET_ID constant: "${SPREADSHEET_ID}"`);
    }
     Logger.log(`[INFO] Attempting to open Spreadsheet with ID: "${SPREADSHEET_ID}"`);
  } catch (e) {
    Logger.log(`[FATAL ERROR] Failed to open spreadsheet with ID "${SPREADSHEET_ID}". Please verify the ID is correct, the spreadsheet exists, and the script has permission. Error: ${e}`);
    return; // Stop execution
  }

  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`[FATAL ERROR] Sheet "${SHEET_NAME}" not found within spreadsheet ID "${SPREADSHEET_ID}". Please verify the SHEET_NAME constant in the script matches the actual tab name in your sheet. Aborting.`);
    return;
  }
   Logger.log(`[INFO] Successfully opened sheet "${SHEET_NAME}".`);

  const range = sheet.getDataRange();
  const values = range.getValues(); // Get all data as a 2D array

  if (values.length <= 1) {
    Logger.log("[INFO] Sheet has no data beyond the header row. Nothing to process. Exiting.");
    return; // No data rows to check
  }

  // Calculate the date threshold
  const thresholdDate = new Date();
  thresholdDate.setDate(thresholdDate.getDate() - (WEEKS_THRESHOLD * 7));
  Logger.log(`[INFO] Threshold Date calculated: Applications last updated before ${thresholdDate.toLocaleDateString()} (and not in a final state) will be marked as stale.`);
  Logger.log(`[INFO] Configured Stale Status: "${REJECTED_STATUS}"`);
  Logger.log(`[INFO] Configured Final/Protected Statuses: ${JSON.stringify(Array.from(FINAL_STATUSES))}`);


  let updatedRowCount = 0;
  let rowsToUpdate = []; // Store row indices (0-based for array) and new status

  // Iterate through rows, starting from the second row (index 1) to skip header
  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    // --- Basic Column Index Validation ---
    if (STATUS_COL <= 0 || STATUS_COL > row.length || LAST_UPDATE_DATE_COL <= 0 || LAST_UPDATE_DATE_COL > row.length) {
        Logger.log(`[WARN] Skipping Row ${i + 1} due to invalid column index configuration (Status Col: ${STATUS_COL}, Last Update Col: ${LAST_UPDATE_DATE_COL}, Row Length: ${row.length}). Check column constants.`);
        continue; // Skip this row if configured columns don't exist
    }

    const currentStatus = row[STATUS_COL - 1]?.toString().trim() ?? ""; // Adjust for 0-based index, handle potential null/undefined
    const lastUpdateDateValue = row[LAST_UPDATE_DATE_COL - 1]; // Adjust for 0-based index

    // --- Validation Checks for this row ---
    // 1. Is status already final/protected?
    if (FINAL_STATUSES.has(currentStatus)) {
      // Logger.log(`[DEBUG] Skipping Row ${i + 1}: Status "${currentStatus}" is in the FINAL_STATUSES set.`);
      continue;
    }

    // 2. Is the Last Update Date valid?
    if (!lastUpdateDateValue || !(lastUpdateDateValue instanceof Date) || isNaN(lastUpdateDateValue.getTime())) {
      // Logger.log(`[DEBUG] Skipping Row ${i + 1}: Invalid or missing Last Update Date (Value: ${lastUpdateDateValue}).`);
      continue; // Skip if date is missing or not a valid Date object
    }

    // --- Check Date Threshold ---
    if (lastUpdateDateValue < thresholdDate) {
      // Only update if the current status is different from the target rejected status
      if (currentStatus !== REJECTED_STATUS) {
        Logger.log(`[INFO] Marking Row ${i + 1} for update: Last Updated (${lastUpdateDateValue.toLocaleDateString()}) < Threshold (${thresholdDate.toLocaleDateString()}). Current status: "${currentStatus}". New Status: "${REJECTED_STATUS}"`);
        // Modify the status in the original 'values' array directly
        values[i][STATUS_COL - 1] = REJECTED_STATUS;
        updatedRowCount++;
        // OPTIONAL: If you add a 'Notes' column, you could update it here too:
        // const notesColIndex = NOTES_COL - 1; // Define NOTES_COL in config if needed
        // if (notesColIndex >= 0 && notesColIndex < row.length) {
        //   values[i][notesColIndex] = (values[i][notesColIndex] ? values[i][notesColIndex] + "; " : "") + `Auto-updated stale on ${new Date().toLocaleDateString()}`;
        // }
      } else {
        // Logger.log(`[DEBUG] Skipping Row ${i + 1}: Already has status "${REJECTED_STATUS}".`);
      }
    }
  } // End row loop

  // --- Batch Write back to Sheet ---
  if (updatedRowCount > 0) {
     Logger.log(`[INFO] Found ${updatedRowCount} stale application(s) to update. Attempting to write changes back to the sheet...`);
    try {
      // Write the entire (potentially modified) data range back
      range.setValues(values);
      Logger.log(`[SUCCESS] Successfully updated ${updatedRowCount} stale application statuses to "${REJECTED_STATUS}" in the sheet.`);
    } catch (e) {
      Logger.log(`[ERROR] Failed to write updates back to the sheet: ${e}\nStack: ${e.stack}`);
      // Consider sending an email notification on error? Requires MailApp scope and setup.
      // MailApp.sendEmail("your-email@example.com", "Auto-Reject Script Failed", `Error writing to sheet: ${e.message}\nSheet ID: ${SPREADSHEET_ID}`);
    }
  } else {
    Logger.log("[INFO] No stale applications found needing updates during this run.");
  }

  const SCRIPT_END_TIME = new Date();
  Logger.log(`==== AUTO-REJECT STALE SCRIPT FINISHED (${SCRIPT_END_TIME.toLocaleString()}) ==== Time Elapsed: ${(SCRIPT_END_TIME - SCRIPT_START_TIME)/1000}s ====`);
}

// --- Trigger Function and Menu Function Removed ---
// Manually set up a time-driven trigger in the Apps Script Editor
// targeting the 'markStaleApplicationsAsRejected' function as described in the header comments.
