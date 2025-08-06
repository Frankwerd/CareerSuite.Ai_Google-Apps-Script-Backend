// SheetSetup.gs
// This file contains functions for setting up and formatting the master resume data sheet.
// It defines the structure of the resume sections and applies styling to the sheet.
// Relies on global constants defined in 'Constants.gs' for styling, sheet names, and structure.

// Constants like HEADER_BACKGROUND_COLOR, RESUME_STRUCTURE, MASTER_RESUME_DATA_SHEET_NAME,
// NUM_DEDICATED_BULLET_COLUMNS, and MASTER_RESUME_SPREADSHEET_ID are expected to be defined in 'Constants.gs'.

// Example in Constants.gs:
// const HEADER_BACKGROUND_COLOR = "#F8BBD0";
// const MASTER_RESUME_DATA_SHEET_NAME = "ResumeData";
// const NUM_DEDICATED_BULLET_COLUMNS = 3;
// const RESUME_STRUCTURE = [ { title: "PERSONAL INFO", headers: ["Key", "Value"] }, ... ];
// const MASTER_RESUME_SPREADSHEET_ID = "YOUR_ID_OR_PLACEHOLDER";


/**
 * Sets up or reformats the master resume data Google Sheet.
 * It creates the sheet if it doesn't exist (or if a new ID is generated),
 * defines section headers, applies formatting, and adds placeholder rows for data entry.
 * Uses the RESUME_STRUCTURE and styling constants defined in 'Constants.gs'.
 */
function setupMasterResumeSheet() {
  let ui;
  try {
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
      ui = SpreadsheetApp.getUi();
    }
    Logger.log(`Attempting sheet setup. Max dedicated bullet columns: ${NUM_DEDICATED_BULLET_COLUMNS}. Target sheet name: "${MASTER_RESUME_DATA_SHEET_NAME}"`);

    let targetSpreadsheetId = MASTER_RESUME_SPREADSHEET_ID; 
    let spreadsheet;
    let newSheetCreatedThisRun = false; // Flag to track if we created the SS

    if (!targetSpreadsheetId || String(targetSpreadsheetId).toUpperCase().includes("YOUR_ACTUAL") || String(targetSpreadsheetId).trim() === "") {
      const newSpreadsheetTitle = `Master Resume Data (${new Date().toLocaleDateString()})`;
      spreadsheet = SpreadsheetApp.create(newSpreadsheetTitle);
      targetSpreadsheetId = spreadsheet.getId();
      newSheetCreatedThisRun = true;
      Logger.log(`NEW Spreadsheet created: "${newSpreadsheetTitle}" (ID: ${targetSpreadsheetId})`);
      if (ui) {
        ui.alert("New Sheet Created", `A new spreadsheet: "${newSpreadsheetTitle}", ID: ${targetSpreadsheetId}. Please update MASTER_RESUME_SPREADSHEET_ID in Constants.gs to this ID if you want to reuse it for future runs.`);
      }
    } else {
      try {
        spreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
        Logger.log(`Opened existing spreadsheet: "${spreadsheet.getName()}" (ID: ${targetSpreadsheetId})`);
      } catch (e) {
        Logger.log(`ERROR: Could not open spreadsheet ID "${targetSpreadsheetId}". ${e.message}. Creating new one.`);
        const newSpreadsheetTitleOnError = `Master Resume Data (New - ${new Date().toLocaleDateString()})`;
        spreadsheet = SpreadsheetApp.create(newSpreadsheetTitleOnError);
        targetSpreadsheetId = spreadsheet.getId();
        newSheetCreatedThisRun = true;
        if (ui) {
          ui.alert("Error Opening Sheet", `Old ID fail. New sheet: "${newSpreadsheetTitleOnError}", ID: ${targetSpreadsheetId}. Update MASTER_RESUME_SPREADSHEET_ID in Constants.gs.`);
        }
      }
    }

    if (!spreadsheet) {
      Logger.log("CRITICAL ERROR: Failed to open or create spreadsheet.");
      if (ui) ui.alert("Error", "Spreadsheet could not be accessed.");
      return;
    }

    let sheetToFormat = spreadsheet.getSheetByName(MASTER_RESUME_DATA_SHEET_NAME);
    let defaultSheet1 = spreadsheet.getSheetByName("Sheet1");

    if (sheetToFormat) { // Desired sheet name already exists
      Logger.log(`Sheet "${MASTER_RESUME_DATA_SHEET_NAME}" found.`);
      // If "Sheet1" also exists AND it's not the same sheet as our target sheet, delete "Sheet1".
      if (defaultSheet1 && defaultSheet1.getSheetId() !== sheetToFormat.getSheetId()) {
        Logger.log(`Also found "Sheet1" (different from target sheet). Deleting "Sheet1".`);
        spreadsheet.deleteSheet(defaultSheet1);
      }
      // Prompt to clear the existing target sheet if it wasn't just created this run
      if (!newSheetCreatedThisRun) { // Only prompt if we didn't just create the spreadsheet (new SS comes with a clear Sheet1 which we handle)
        const confirmClearMsg = `Sheet "${MASTER_RESUME_DATA_SHEET_NAME}" exists. Clear its contents and reformat? This will ERASE existing data.`;
        const userChoice = ui ? ui.alert("Confirm Clear", confirmClearMsg, ui.ButtonSet.YES_NO) : 'YES';
        if (userChoice === ui.Button.YES || userChoice === 'YES') {
          sheetToFormat.clear(); // Clears content and formatting
          Logger.log(`Cleared existing sheet: "${MASTER_RESUME_DATA_SHEET_NAME}".`);
        } else {
          Logger.log("Setup aborted by user (did not clear existing sheet).");
          if (ui) ui.alert("Setup Aborted", "Sheet setup cancelled.");
          return;
        }
      } else {
         // If the whole SS was new this run, and MASTER_RESUME_DATA_SHEET_NAME was "Sheet1" initially (newly created SS).
         // Then sheetToFormat would be "Sheet1". We rename and clear it.
         if (sheetToFormat.getName() === "Sheet1" && MASTER_RESUME_DATA_SHEET_NAME !== "Sheet1") {
             sheetToFormat.setName(MASTER_RESUME_DATA_SHEET_NAME);
             Logger.log(`Renamed initial "Sheet1" to "${MASTER_RESUME_DATA_SHEET_NAME}".`);
         }
         sheetToFormat.clear(); // Always clear if we are dealing with a sheet in a newly created SS
         Logger.log(`Ensured sheet "${MASTER_RESUME_DATA_SHEET_NAME}" is clear in newly created spreadsheet.`);
      }
    } else { // Desired sheet name does NOT exist
      Logger.log(`Sheet "${MASTER_RESUME_DATA_SHEET_NAME}" not found.`);
      if (defaultSheet1) { // "Sheet1" exists
        // If this is the *only* sheet, rename it to our target name.
        // Or if the whole spreadsheet was newly created this run (it will only have "Sheet1").
        if (spreadsheet.getSheets().length === 1 || newSheetCreatedThisRun) {
          Logger.log(`Renaming existing "Sheet1" to "${MASTER_RESUME_DATA_SHEET_NAME}".`);
          defaultSheet1.setName(MASTER_RESUME_DATA_SHEET_NAME);
          sheetToFormat = defaultSheet1;
          sheetToFormat.clear(); // Clear the renamed sheet
        } else {
          // "Sheet1" exists, but other sheets also exist. So, insert our new target sheet
          // and explicitly delete the original "Sheet1" IF it's not desired by MASTER_RESUME_DATA_SHEET_NAME.
          sheetToFormat = spreadsheet.insertSheet(MASTER_RESUME_DATA_SHEET_NAME);
          Logger.log(`Inserted new sheet "${MASTER_RESUME_DATA_SHEET_NAME}".`);
          if (MASTER_RESUME_DATA_SHEET_NAME !== "Sheet1") { // Only delete if "Sheet1" isn't our target name
             Logger.log(`Also found "Sheet1". Deleting it as a new target sheet was inserted.`);
             spreadsheet.deleteSheet(defaultSheet1);
          }
        }
      } else { // Neither desired sheet nor "Sheet1" exists
        sheetToFormat = spreadsheet.insertSheet(MASTER_RESUME_DATA_SHEET_NAME);
        Logger.log(`Inserted new sheet "${MASTER_RESUME_DATA_SHEET_NAME}" (no "Sheet1" was present).`);
      }
    }

    if (!sheetToFormat) {
        Logger.log("CRITICAL ERROR: Could not obtain a valid sheet object to format.");
        if (ui) ui.alert("Error", "Could not prepare sheet for formatting.");
        return;
    }
    
    // --- Now sheetToFormat is the one we work on. Ensure it's clear before applying format ---
    // The logic above should have cleared it, but one final clear here is okay if paranoid,
    // or remove this if you trust the specific clear() calls above.
    // Let's assume previous logic handled clearing and only activate.
    sheetToFormat.activate();
    // If not cleared by specific logic above (e.g. user said NO to clearing existing), we should respect it.
    // The logic in the "if (sheetToFormat)" block handles clearing when necessary.

    // ... (rest of your setupMasterResumeSheet function to apply formatting to sheetToFormat) ...
    // Replace all `sheet.` with `sheetToFormat.`
    // For example:
    // sheetToFormat.setColumnWidth(1, 250); ...
    // RESUME_STRUCTURE.forEach(section => { ... sheetToFormat.getRange(...); ... });

    let maxHeaderLength = 0;
    RESUME_STRUCTURE.forEach(section => { // RESUME_STRUCTURE from Constants.gs
      let currentSectionHeaders = section.headers ? [...section.headers] : [];
      if (["EXPERIENCE", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "PROJECTS"].includes(section.title)) {
        const bulletBaseName = (section.title === "PROJECTS") ? "DescriptionBullet" : "Responsibility";
        for (let i = 1; i <= NUM_DEDICATED_BULLET_COLUMNS; i++) { // NUM_DEDICATED_BULLET_COLUMNS from Constants.gs
          const bulletHeader = `${bulletBaseName}${i}`;
          if (!currentSectionHeaders.find(h => h.toUpperCase() === bulletHeader.toUpperCase())) { 
            currentSectionHeaders.push(bulletHeader);
          }
        }
      }
      if (currentSectionHeaders.length > maxHeaderLength) maxHeaderLength = currentSectionHeaders.length;
    });
    const numColumnsToFormat = Math.max(2, maxHeaderLength);

    sheetToFormat.setColumnWidth(1, 250); 
    if (numColumnsToFormat >= 2) sheetToFormat.setColumnWidth(2, 300); 
    for (let c = 3; c <= numColumnsToFormat; c++) {
      sheetToFormat.setColumnWidth(c, 180);
    }
    if (sheetToFormat.getMaxColumns() < numColumnsToFormat) {
      sheetToFormat.insertColumns(sheetToFormat.getMaxColumns() + 1, numColumnsToFormat - sheetToFormat.getMaxColumns());
    }
    if (sheetToFormat.getMaxRows() < 100) { 
        sheetToFormat.insertRows(sheetToFormat.getMaxRows() + 1, 100 - sheetToFormat.getMaxRows());
    }

    let currentRow = 1;
    RESUME_STRUCTURE.forEach(section => { 
      const mainHeaderRange = sheetToFormat.getRange(currentRow, 1, 1, numColumnsToFormat);
      mainHeaderRange.setValue(section.title.toUpperCase())
        .setBackground(HEADER_BACKGROUND_COLOR) 
        .setFontColor(HEADER_FONT_COLOR)       
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("left").setVerticalAlignment("middle")
        .setBorder(true, true, true, true, null, null, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM) 
        .mergeAcross();
      currentRow++;

      let sectionSpecificHeaders = section.headers ? [...section.headers] : [];
       if (["EXPERIENCE", "LEADERSHIP & UNIVERSITY INVOLVEMENT"].includes(section.title)) {
            for (let i = 1; i <= NUM_DEDICATED_BULLET_COLUMNS; i++) {
                const headerName = `Responsibility${i}`;
                if (!sectionSpecificHeaders.find(h => h.toUpperCase() === headerName.toUpperCase())) sectionSpecificHeaders.push(headerName);
            }
        } else if (section.title === "PROJECTS") {
            for (let i = 1; i <= NUM_DEDICATED_BULLET_COLUMNS; i++) {
                const headerName = `DescriptionBullet${i}`;
                 if (!sectionSpecificHeaders.find(h => h.toUpperCase() === headerName.toUpperCase())) sectionSpecificHeaders.push(headerName);
            }
        }
      
      if (section.title === "PERSONAL INFO") {
        const personalInfoKeys = (section.headers && section.headers[0] && section.headers[0].toUpperCase() === "KEY") 
                               ? ["Full Name", "Location", "Phone", "Email", "LinkedIn", "Portfolio", "GitHub"] 
                               : ["fullName", "location", "phone", "email", "linkedin", "portfolio", "github"]; 
        personalInfoKeys.forEach(key => {
          sheetToFormat.getRange(currentRow, 1).setValue(key).setFontWeight("normal").setBackground(SUB_HEADER_BACKGROUND_COLOR);
          sheetToFormat.getRange(currentRow++, 2).setValue("").setBackground(null); 
        });
      } else if (section.title === "SUMMARY") {
        const summaryPlaceholderRow = sheetToFormat.getRange(currentRow, 1, 1, numColumnsToFormat);
        summaryPlaceholderRow.setValue("(Enter Summary content here - typically in Column B, or Column A if columns are merged for wider input)");
        currentRow++; 
      } else if (sectionSpecificHeaders.length > 0) {
        const subHeaderRange = sheetToFormat.getRange(currentRow, 1, 1, sectionSpecificHeaders.length);
        subHeaderRange.setValues([sectionSpecificHeaders])
          .setBackground(SUB_HEADER_BACKGROUND_COLOR) 
          .setFontWeight("bold").setFontSize(10)
          .setBorder(null, null, true, true, null, null, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK); 
        currentRow++;
      }
      
      const blankDataRowsToInsert = (section.title !== "PERSONAL INFO" && section.title !== "SUMMARY") ? 2 : 1;
      if (currentRow + blankDataRowsToInsert <= sheetToFormat.getMaxRows()) {
        sheetToFormat.getRange(currentRow, 1, blankDataRowsToInsert, numColumnsToFormat).clearContent(); 
      }
      currentRow += blankDataRowsToInsert + 1; 
    });
    
    Logger.log(`Sheet structure for "${MASTER_RESUME_DATA_SHEET_NAME}" created/updated in spreadsheet ID: ${targetSpreadsheetId}.`);
    if (ui) {
      ui.alert("Sheet Setup Complete!", `Sheet "${MASTER_RESUME_DATA_SHEET_NAME}" in spreadsheet ID: ${targetSpreadsheetId} structured.`);
    }
  } catch (e) {
    Logger.log(`CRITICAL ERROR during sheet setup: ${e.message}\nStack: ${e.stack}`);
    if (ui) ui.alert("Error During Setup", `An error occurred: ${e.message}`);
  }
}


// ... (your onOpen function can remain the same)
/**
 * Adds a custom menu to the Google Spreadsheet UI when the spreadsheet is opened.
 * This function should ideally be defined only once across all .gs files in the project.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Resume Toolbelt')
      .addItem('SETUP: Create/Format Resume Data Sheet', 'setupMasterResumeSheet')
      .addToUi();
  } catch (e) {
    Logger.log(`Error creating custom menu: ${e.message}. This script might not be bound to a spreadsheet.`);
  }
}

/**
 * Adds a custom menu to the Google Spreadsheet UI when the spreadsheet is opened.
 * This function should ideally be defined only once across all .gs files in the project.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Resume Toolbelt')
      .addItem('SETUP: Create/Format Resume Data Sheet', 'setupMasterResumeSheet')
      // Add other top-level menu items relevant to this utility here
      .addToUi();
  } catch (e) {
    Logger.log(`Error creating custom menu: ${e.message}. This script might not be bound to a spreadsheet.`);
  }
}
