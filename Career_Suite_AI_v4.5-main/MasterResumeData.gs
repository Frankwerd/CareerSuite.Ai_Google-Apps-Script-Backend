// MasterResumeData.gs
// This file is responsible for fetching and parsing master resume data from a designated Google Sheet.
// It transforms the tabular sheet data into a structured JavaScript object.
// It relies on global constants defined in 'Constants.gs' for spreadsheet ID, sheet name, etc.

// Global constants like MASTER_RESUME_SPREADSHEET_ID, MASTER_RESUME_DATA_SHEET_NAME, 
// and NUM_DEDICATED_BULLET_COLUMNS are now expected to be defined in 'Constants.gs'.
// Example from Constants.gs:
// const MASTER_RESUME_SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
// const MASTER_RESUME_DATA_SHEET_NAME = "ResumeData";
// const NUM_DEDICATED_BULLET_COLUMNS = 3; // Renamed from NUM_BULLET_COLUMNS_TO_PARSE for clarity

// --- HELPER FUNCTIONS for Data Parsing ---

/**
 * Formats a date value from a Google Sheet cell.
 * Handles Date objects, date strings, and "Present".
 * @param {*} dateCellValue The value from the sheet cell.
 * @param {boolean} [isPotentialEndDate=false] If true, an empty value defaults to "Present".
 * @return {string} The formatted date string (e.g., "Month Year", "Present", or original if unparseable).
 */
function formatDateFromSheet(dateCellValue, isPotentialEndDate = false) {
  if (dateCellValue === null || dateCellValue === undefined || String(dateCellValue).trim() === "") {
    return isPotentialEndDate ? "Present" : "";
  }
  const dateString = String(dateCellValue).trim();
  if (dateString.toUpperCase() === "PRESENT") {
    return "Present";
  }

  let dateObj;
  if (dateCellValue instanceof Date) { // Check if it's already a Date object from Sheets
    dateObj = dateCellValue;
  } else { // If it's a string, try to parse it
    // Attempt common date parsing; robust library might be needed for very diverse formats
    dateObj = new Date(dateString);
  }

  if (isNaN(dateObj.getTime())) {
    // Logger.log_DEBUG(`formatDateFromSheet: Could not parse date from "${dateString}", returning as is.`);
    return dateString; // Return original string if it's not a reliably processable date
  }

  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  // If the original string was just a year (YYYY), or already "Month YYYY", return it to preserve user's specific formatting.
  if (/^\d{4}$/.test(dateString)) { return dateString; }
  if (/^[A-Za-z]{3,9}\s\d{4}$/.test(dateString)) { return dateString; }
  
  return `${monthNames[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
}

/**
 * Splits multi-line text from a sheet cell (entered with Alt+Enter or containing literal \n)
 * into an array of strings, trimming whitespace and removing empty lines.
 * @param {string} textCellContent The string content from the sheet cell.
 * @return {string[]} An array of trimmed, non-empty lines.
 */
function smartSplitMultiLine(textCellContent) {
  if (!textCellContent || typeof textCellContent !== 'string' || textCellContent.trim() === "") {
    return [];
  }
  // Logger.log_DEBUG(`smartSplitMultiLine - INPUT: "${textCellContent}"`);
  // Step 1: Replace any literal "\\n" (two characters: backslash followed by n) with an actual newline character '\n'.
  const stringWithActualNewlines = textCellContent.replace(/\\n/g, '\n');
  // Step 2: Split the string by any standard newline sequence (\r\n, \r, or \n).
  const linesArray = stringWithActualNewlines.split(/\r\n|\r|\n/g);
  const cleanedLines = linesArray.map(s => s.trim()).filter(s => s && s.length > 0);
  // Logger.log_DEBUG(`smartSplitMultiLine - OUTPUT Count: ${cleanedLines.length}, Lines: ${JSON.stringify(cleanedLines)}`);
  return cleanedLines;
}

/**
 * Gets a canonical (standardized) section title from raw sheet header text.
 * Handles variations in capitalization and minor phrasing.
 * @param {string} text The raw text from the sheet (typically first cell of a section header row).
 * @return {string|null} The canonical section title (e.g., "EXPERIENCE") or null if not recognized.
 */
function getCanonicalSectionTitle(text) {
  if (!text || typeof text !== 'string') return null;
  const upperText = text.trim().toUpperCase();
  if (upperText.startsWith("PERSONAL INFO")) return "PERSONAL INFO";
  if (upperText.startsWith("SUMMARY")) return "SUMMARY";
  if (upperText.startsWith("EXPERIENCE")) return "EXPERIENCE";
  if (upperText.startsWith("EDUCATION")) return "EDUCATION";
  if (upperText.startsWith("PROJECTS")) return "PROJECTS";
  if (upperText.startsWith("TECHNICAL SKILLS") || upperText.includes("CERTIFICATES")) return "TECHNICAL SKILLS & CERTIFICATES";
  if (upperText.startsWith("LEADERSHIP") || upperText.includes("INVOLVEMENT")) return "LEADERSHIP & UNIVERSITY INVOLVEMENT";
  if (upperText.startsWith("HONORS") || upperText.startsWith("AWARDS")) return "HONORS & AWARDS";
  return null; // Not a recognized section start
}

/**
 * Checks if a given canonical section title typically contains tabular data with column headers.
 * @param {string} sectionTitleKey The canonical section title.
 * @return {boolean} True if the section is expected to be tabular, false otherwise.
 */
function isTabularSection(sectionTitleKey) {
  return ["EXPERIENCE", "EDUCATION", "PROJECTS", "TECHNICAL SKILLS & CERTIFICATES",
          "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"].includes(sectionTitleKey);
}

/**
 * Checks if a row array (from sheet.getValues()) is effectively blank
 * (all cells are null, undefined, or empty strings after trimming).
 * @param {Array<*>} rowArray An array representing a row of cells.
 * @return {boolean} True if the row is effectively blank, false otherwise.
 */
function rowIsEffectivelyBlank(rowArray) {
  if (!rowArray || !Array.isArray(rowArray)) return true;
  return rowArray.every(cell => (cell === null || cell === undefined || String(cell).trim() === ""));
}


// --- MAIN DATA FETCHING AND PARSING FUNCTION ---

/**
 * Fetches and parses master resume data from a specified Google Sheet.
 * The sheet is expected to have sections like "PERSONAL INFO", "SUMMARY", "EXPERIENCE", etc.,
 * with specific formatting (section titles in column A, then headers for tabular data).
 *
 * @param {string} [spreadsheetId=MASTER_RESUME_SPREADSHEET_ID] The ID of the Google Spreadsheet.
 *        Defaults to `MASTER_RESUME_SPREADSHEET_ID` from Constants.gs.
 * @param {string} [dataSheetName=MASTER_RESUME_DATA_SHEET_NAME] The name of the sheet tab.
 *        Defaults to `MASTER_RESUME_DATA_SHEET_NAME` from Constants.gs.
 * @return {Object|null} The structured resume data object, or null on error.
 *         The object structure includes: { resumeSchemaVersion, personalInfo, summary, sections: [...] }.
 */
function getMasterResumeData(
    spreadsheetId = MASTER_RESUME_SPREADSHEET_ID, // From Constants.gs
    dataSheetName = MASTER_RESUME_DATA_SHEET_NAME   // From Constants.gs
  ) {
  Logger.log(`--- Loading Master Resume from SSID: "${spreadsheetId}", Sheet: "${dataSheetName}" ---`);
  
  const resumeOutput = {
    // Consider using CURRENT_RESUME_DATA_SCHEMA_VERSION from Constants.gs
    resumeSchemaVersion: "1.2.8_final_parser_with_all_logs", // Or use: CURRENT_RESUME_DATA_SCHEMA_VERSION
    personalInfo: {},
    summary: "",
    sections: []
  };

  try {
    // Step 1: Validate and Open Spreadsheet
    Logger.log(`Parser - Step 1: Validating and opening Spreadsheet ID: "${spreadsheetId}"`);
    if (!spreadsheetId || String(spreadsheetId).toUpperCase().includes("YOUR_") || String(spreadsheetId).trim() === "") {
      Logger.log("CRITICAL ERROR (getMasterResumeData): Spreadsheet ID is invalid or a placeholder. Please update MASTER_RESUME_SPREADSHEET_ID in Constants.gs or pass a valid ID.");
      try { SpreadsheetApp.getUi().alert("Spreadsheet ID Error", "Master Resume Spreadsheet ID is missing or invalid. Please configure it in the script (Constants.gs)."); } catch (uiError) { /* UI not available */ }
      return null;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log(`Parser - Step 2: Spreadsheet ${ss ? `successfully opened (Name: ${ss.getName()})` : 'FAILED TO OPEN'}.`);
    if (!ss) return null;

    // Step 2: Get Sheet
    Logger.log(`Parser - Step 3: Getting sheet by name: "${dataSheetName}"`);
    const sheet = ss.getSheetByName(dataSheetName);
    Logger.log(`Parser - Step 4: Sheet ${sheet ? `found (Name: ${sheet.getName()})` : 'NOT FOUND'}.`);
    if (!sheet) {
        Logger.log(`ERROR: Sheet with name "${dataSheetName}" not found in spreadsheet ID "${spreadsheetId}".`);
        return null;
    }

    // Step 3: Get Data from Sheet
    Logger.log(`Parser - Step 5: Getting data range from sheet "${sheet.getName()}". Max rows: ${sheet.getMaxRows()}, Last row with content: ${sheet.getLastRow()}`);
    if (sheet.getLastRow() === 0) {
      Logger.log(`ERROR: The sheet "${dataSheetName}" appears to be completely empty.`);
      return null;
    }
    const allData = sheet.getDataRange().getValues();
    Logger.log(`Parser - Step 6: Retrieved ${allData ? allData.length : 'N/A'} rows. First row columns: ${allData && allData[0] ? allData[0].length : 'N/A'}.`);
    if (!allData || allData.length === 0) {
      Logger.log("ERROR: No data retrieved from the sheet range. The sheet might be empty or inaccessible.");
      return null;
    }
    
    // Step 4: Iterate Through Rows and Process Sections
    Logger.log("Parser - Step 7: Starting iteration through sheet rows to identify and process sections...");
    let currentSectionTitleKey = null;
    let currentSectionRawData = [];
    let currentSectionHeaders = [];

    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const firstCellRawText = (row[0] || "").toString().trim();
      let potentialNewSectionKey = getCanonicalSectionTitle(firstCellRawText);

      if (potentialNewSectionKey) { // Start of a new section
        // Process any previously accumulated section data
        if (currentSectionTitleKey && (currentSectionRawData.length > 0 || (currentSectionTitleKey === "SUMMARY" && resumeOutput.summary.length > 0))) {
          processAccumulatedSection(resumeOutput, currentSectionTitleKey, currentSectionHeaders, currentSectionRawData);
        }
        // Reset for the new section
        currentSectionTitleKey = potentialNewSectionKey;
        Logger.log(`Found section: ${currentSectionTitleKey} (from sheet text: "${firstCellRawText}") at sheet row ${i + 1}`);
        currentSectionRawData = [];
        currentSectionHeaders = [];

        // If it's a tabular section, try to find its headers in the next non-blank row
        if (isTabularSection(currentSectionTitleKey) && (i + 1) < allData.length) {
          let headerRowIndex = i + 1;
          // Skip any blank rows between section title and its headers
          while (headerRowIndex < allData.length && rowIsEffectivelyBlank(allData[headerRowIndex])) {
            headerRowIndex++;
          }
          // Check if the identified header row isn't another section title itself
          if (headerRowIndex < allData.length && !getCanonicalSectionTitle((allData[headerRowIndex][0] || "").toString().trim())) {
            currentSectionHeaders = (allData[headerRowIndex] || []).map(h => (h || "").toString().trim());
            Logger.log(`  Headers for ${currentSectionTitleKey}: ${currentSectionHeaders.filter(h => h).join(" | ")} (from sheet row ${headerRowIndex + 1})`);
            i = headerRowIndex; // Advance main loop past this header row
          } else {
            currentSectionHeaders = []; // No valid headers found
            Logger.log(`WARNING: No data headers found for tabular section ${currentSectionTitleKey} immediately after row ${i + 1}. Next non-blank row content: "${allData[headerRowIndex] ? (allData[headerRowIndex][0] || "") : "END OF DATA"}"`);
          }
        }
        continue; // Move to next row after processing section header
      }

      // If currently within a section, accumulate its data
      if (currentSectionTitleKey) {
        if (rowIsEffectivelyBlank(row)) {
          // Logger.log_DEBUG(`  Skipping blank row ${i + 1} within section ${currentSectionTitleKey}`);
          continue;
        }

        if (currentSectionTitleKey === "SUMMARY") {
          let summaryTextInRow = "";
          const cellBText = (row[1] || "").toString().trim(); // Potential summary in Col B
          const cellAText = (row[0] || "").toString().trim(); // Potential summary in Col A (if B is instructional)
          
          // Heuristic: Prefer text in Col B unless it looks like an instruction or is very short
          if (cellBText && !cellBText.startsWith("(") && cellBText.length > 5) {
            summaryTextInRow = cellBText;
          } else if (cellAText && !cellAText.startsWith("(") && !getCanonicalSectionTitle(cellAText) && cellAText.length > 5) {
            summaryTextInRow = cellAText;
          }
          
          if (summaryTextInRow) {
            currentSectionRawData.push([summaryTextInRow]); // Store as array for consistent processing
          } else if ((cellAText && cellAText.startsWith("(")) || (cellBText && cellBText.startsWith("("))) {
            Logger.log(`  Skipping potential instruction line in SUMMARY section: "${cellAText || cellBText}"`);
          }
          continue;
        }

        // Sanity check for tabular sections: if we are accumulating data but have no headers, something is wrong.
        if (isTabularSection(currentSectionTitleKey) && currentSectionHeaders.length === 0 && currentSectionRawData.length > 1) {
          Logger.log(`ERROR: Data row ${i + 1} encountered for tabular section "${currentSectionTitleKey}" but no headers were identified. Aborting accumulation for this misformatted section.`);
          currentSectionTitleKey = null; // Stop processing this malformed "section"
          continue;
        }
        currentSectionRawData.push(row.map(c => (c === null || c === undefined) ? "" : String(c))); // Keep raw strings, toString for safety
      }
    }

    // Process the last accumulated section after loop finishes
    if (currentSectionTitleKey && (currentSectionRawData.length > 0 || (currentSectionTitleKey === "SUMMARY" && resumeOutput.summary.length > 0))) {
      processAccumulatedSection(resumeOutput, currentSectionTitleKey, currentSectionHeaders, currentSectionRawData);
    }
    
    // --- Final Parsed Data Verification Log Block (Valuable for Debugging) ---
    Logger.log("----- DETAILED PARSED DATA VERIFICATION (from getMasterResumeData) -----");
    Logger.log(`PersonalInfo: ${JSON.stringify(resumeOutput.personalInfo, null, 2)}`);
    Logger.log(`Summary (length ${resumeOutput.summary.length}): "${resumeOutput.summary.substring(0, 100)}..."`);
    Logger.log(`Total Sections Parsed: ${resumeOutput.sections.length}`);
    
    const expectedSectionTitles = ["TECHNICAL SKILLS & CERTIFICATES", "EXPERIENCE", "PROJECTS", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS", "EDUCATION"];
    expectedSectionTitles.forEach(title => {
      const section = resumeOutput.sections.find(s => s.title === title);
      if (section) {
        Logger.log(`--- VERIFY Section: ${title} ---`);
        if (section.items) {
          Logger.log(`  Found ${section.items.length} items.`);
          if (section.items.length > 0) {
            Logger.log(`    First item sample: ${JSON.stringify(section.items[0]).substring(0, 250)}...`);
            if (section.items[0].responsibilities && Array.isArray(section.items[0].responsibilities)) {
              Logger.log(`      First item, Responsibilities count: ${section.items[0].responsibilities.length}. First: "${(section.items[0].responsibilities[0]||"").substring(0,50)}..."`);
            }
            if (title === "EDUCATION" && section.items[0].relevantCoursework && Array.isArray(section.items[0].relevantCoursework)) {
              Logger.log(`      First Edu item, Courses count: ${section.items[0].relevantCoursework.length}. First: "${(section.items[0].relevantCoursework[0]||"").substring(0,50)}..."`);
            }
          }
        } else if (section.subsections) {
          Logger.log(`  Found ${section.subsections.length} subsections.`);
          if (section.subsections.length > 0 && section.subsections[0].items) {
            Logger.log(`    First subsection ("${section.subsections[0].name}") has ${section.subsections[0].items.length} items.`);
            if (section.subsections[0].items.length > 0) {
              Logger.log(`      First item in first subsection sample: ${JSON.stringify(section.subsections[0].items[0]).substring(0, 250)}...`);
              if (section.title === "PROJECTS" && section.subsections[0].items[0].descriptionBullets) {
                Logger.log(`        First Proj item, Description Bullets count: ${section.subsections[0].items[0].descriptionBullets.length}. First: "${(section.subsections[0].items[0].descriptionBullets[0]||"").substring(0,50)}..."`);
              }
              // Add more specific logging for other subsection types if needed
            }
          }
        }
      } else {
        Logger.log(`WARNING (Verification): Expected section "${title}" was NOT FOUND in parsed data.`);
      }
    });
    Logger.log("----- END OF DETAILED PARSED DATA VERIFICATION -----");
    // --- End Verification Log Block ---

    Logger.log(`SUCCESS (getMasterResumeData): Finished loading from Sheet. Candidate: ${resumeOutput.personalInfo.fullName || "N/A"}`);
    return resumeOutput;

  } catch (e) {
    Logger.log(`CRITICAL EXCEPTION in getMasterResumeData: ${e.message}\nStack: ${e.stack}`);
    return null; // Return null on any critical error
  }
}

/**
 * Processes accumulated raw data for a specific section and maps it to the
 * structured `resumeOutput` object. This is a helper for `getMasterResumeData`.
 *
 * @param {Object} resumeOutput The main resume object being built.
 * @param {string} sectionTitleKey The canonical title of the section being processed.
 * @param {string[]} headers An array of header strings for tabular data (empty for non-tabular).
 * @param {Array<Array<string>>} dataRows An array of rows, where each row is an array of cell strings.
 */
function processAccumulatedSection(resumeOutput, sectionTitleKey, headers, dataRows) {
  // Logger.log_DEBUG(`  processAccumulatedSection: Processing "${sectionTitleKey}", Headers: [${headers.join(", ")}], Data Rows: ${dataRows.length}`);

  // --- PERSONAL INFO Section ---
  if (sectionTitleKey === "PERSONAL INFO") {
    dataRows.forEach(row => {
      let keyRaw = (row[0] || "").toString().trim();
      const value = (row[1] || "").toString().trim();
      if (!keyRaw || !value) return; // Skip if key or value is missing

      let keyNormalized = keyRaw.toLowerCase().replace(/\s+/g, '');
      if (keyNormalized.includes("fullname") || keyNormalized.includes("name")) resumeOutput.personalInfo.fullName = value;
      else if (keyNormalized.includes("linkedin")) resumeOutput.personalInfo.linkedin = value;
      else if (keyNormalized.includes("github")) resumeOutput.personalInfo.github = value;
      else if (keyNormalized.includes("portfolio")) resumeOutput.personalInfo.portfolio = value;
      else if (keyNormalized.includes("phone")) resumeOutput.personalInfo.phone = value;
      else if (keyNormalized.includes("email")) resumeOutput.personalInfo.email = value;
      else if (keyNormalized.includes("location")) resumeOutput.personalInfo.location = value;
      else { // Fallback for any other keys not explicitly mapped
        let camelCaseKey = keyRaw.replace(/\s+(.)/g, (_match, char) => char.toUpperCase());
        camelCaseKey = camelCaseKey.charAt(0).toLowerCase() + camelCaseKey.slice(1);
        resumeOutput.personalInfo[camelCaseKey] = value;
      }
    });
    // Ensure standard 'fullName' key if 'fullname' was used
    if (!resumeOutput.personalInfo.fullName && resumeOutput.personalInfo.fullname) {
      resumeOutput.personalInfo.fullName = resumeOutput.personalInfo.fullname;
      delete resumeOutput.personalInfo.fullname;
    }
    return;
  }

  // --- SUMMARY Section ---
  if (sectionTitleKey === "SUMMARY") {
    resumeOutput.summary = dataRows.map(rowArray => (rowArray[0] || "").trim()).join("\n").trim();
    return;
  }

  // --- Handling for Tabular Sections ---
  if (!isTabularSection(sectionTitleKey)) { // Should not happen if logic in getMasterResumeData is correct
      Logger.log(`WARN (processAccumulatedSection): "${sectionTitleKey}" is not a recognized tabular section type for detailed processing here. Skipping.`);
      return;
  }

  if (!headers || headers.length === 0 || dataRows.length === 0) {
    Logger.log(`INFO (processAccumulatedSection): No headers or data rows for tabular section "${sectionTitleKey}". Ensuring empty section structure in output.`);
    // Ensure section exists in output even if empty, with appropriate structure (items or subsections)
    if (!resumeOutput.sections.find(s => s.title === sectionTitleKey)) {
      if (sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES" || sectionTitleKey === "PROJECTS") {
        resumeOutput.sections.push({ title: sectionTitleKey, subsections: (sectionTitleKey === "PROJECTS" ? [{ name: "General Projects", items: []}] : []) });
      } else {
        resumeOutput.sections.push({ title: sectionTitleKey, items: [] });
      }
    }
    return;
  }

  // Step 1: Convert sheet rows to preliminary objects using headers from the sheet
  const itemsFromSheet = dataRows.map(row => {
    let item = {};
    headers.forEach((header, index) => {
      if (header && header.trim() !== "") { // Only process if header is not empty
        item[header.trim()] = (row[index] === null || row[index] === undefined) ? "" : String(row[index]); // Convert to string
      }
    });
    return item;
  });

  // Step 2: Map sheet headers to desired camelCase JS object keys
  const mappedItemsIntermediate = itemsFromSheet.map(itemFromSheet => {
    const newItem = {};
    for (const headerKeyInSheet in itemFromSheet) {
      let jsKey = headerKeyInSheet.replace(/\s+/g, ''); // Remove spaces
      jsKey = jsKey.charAt(0).toLowerCase() + jsKey.slice(1); // Basic camelCase

      // --- Explicit remappings from common Sheet headers to desired JS camelCase keys ---
      const keyMappings = {
        "JobTitle": "jobTitle", "Company": "company", /* "Location" is fine */
        "StartDate": "startDate", "EndDate": "endDate",
        "Institution": "institution", "Degree": "degree", /* "Location" */ "RelevantCoursework": "relevantCoursework",
        "CategoryName": "categoryName", "SkillItem": "skillItem", /* "Details" */
        "Issuer": "issuer", "IssueDate": "issueDate",
        "ProjectName": "projectName", "Organization": "organization", /* "Role" */ /* "StartDate", "EndDate" */
        /* "Technologies" */ "Impact": "impact", "FutureDevelopment": "futureDevelopment",
        "GitHubName1": "githubName1", "GitHubURL1": "githubUrl1", // Standardized to github...
        "AwardName": "awardName",
        // Common variations (if any) for the numbered bullet columns (e.g. "Responsibility 1")
        // will be handled by the generic camelCasing turning them into "responsibility1" etc.
      };
      
      // Apply specific mappings or fall back to generic camelCasing
      if (keyMappings[headerKeyInSheet]) jsKey = keyMappings[headerKeyInSheet];
      else if (headerKeyInSheet === "Date" && sectionTitleKey === "HONORS & AWARDS") jsKey = "date";
      else if (headerKeyInSheet === "Description" && sectionTitleKey === "LEADERSHIP & UNIVERSITY INVOLVEMENT") jsKey = "description";
      else if (headerKeyInSheet === "GPA" && sectionTitleKey === "EDUCATION") jsKey = "gpa";
      // Other direct camelCases like 'location', 'details', 'role', 'technologies' are fine as is after generic camelCasing.
      
      newItem[jsKey] = (itemFromSheet[headerKeyInSheet] || "").toString().trim(); // Trim string values here
    }
    return newItem;
  });

  // Step 3: Process specific fields (dates, collecting numbered bullets, splitting multi-line cells)
  // Assumes NUM_DEDICATED_BULLET_COLUMNS is available from Constants.gs
  const processedItems = mappedItemsIntermediate.map(item => {
    // Date formatting for standard date fields
    if (item.startDate) item.startDate = formatDateFromSheet(item.startDate);
    if (item.endDate) item.endDate = formatDateFromSheet(item.endDate, true); // true indicates it's a potential end date
    if (item.date && sectionTitleKey === "HONORS & AWARDS") item.date = formatDateFromSheet(item.date);
    if (item.issueDate && sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES") item.issueDate = formatDateFromSheet(item.issueDate);
    
    // Collect numbered bullet columns (e.g., responsibility1, responsibility2) into an array
    if (sectionTitleKey === "EXPERIENCE" || sectionTitleKey === "LEADERSHIP & UNIVERSITY INVOLVEMENT") {
      let bulletsArray = [];
      for (let k = 1; k <= NUM_DEDICATED_BULLET_COLUMNS; k++) {
        const bulletKey = `responsibility${k}`; // Expects e.g., "responsibility1"
        if (item[bulletKey] && item[bulletKey].trim() !== "") {
          bulletsArray.push(item[bulletKey].trim());
        }
        delete item[bulletKey]; // Clean up original numbered key
      }
      item.responsibilities = bulletsArray; // Add the collected array
    } else if (sectionTitleKey === "PROJECTS") {
      let descBulletsArray = [];
      for (let k = 1; k <= NUM_DEDICATED_BULLET_COLUMNS; k++) {
        const bulletKey = `descriptionBullet${k}`; // Expects e.g., "descriptionBullet1"
        if (item[bulletKey] && item[bulletKey].trim() !== "") {
          descBulletsArray.push(item[bulletKey].trim());
        }
        delete item[bulletKey];
      }
      item.descriptionBullets = descBulletsArray;

      // Handle 'technologies' if it's a single cell meant for splitting
      if (typeof item.technologies === 'string') item.technologies = smartSplitMultiLine(item.technologies);
      else if (!Array.isArray(item.technologies)) item.technologies = []; // Ensure it's an array if not a string
    }
    
    // Handle other common multi-line fields (e.g., 'relevantCoursework' in EDUCATION)
    if (sectionTitleKey === "EDUCATION" && typeof item.relevantCoursework === 'string') {
      item.relevantCoursework = smartSplitMultiLine(item.relevantCoursework);
    } else if (sectionTitleKey === "EDUCATION" && !Array.isArray(item.relevantCoursework)) {
      item.relevantCoursework = []; // Ensure array if not string
    }
    
    // Final trim for any remaining top-level string properties
    for (const key in item) {
      if (typeof item[key] === 'string' && !Array.isArray(item[key])) { // Don't trim arrays by mistake
        item[key] = item[key].trim();
      }
    }
    return item;
  });

  // Step 4: Push processed items into the correct section structure in resumeOutput object
  let sectionObject = resumeOutput.sections.find(s => s.title === sectionTitleKey);
  if (!sectionObject) { // If section doesn't exist yet, create it
      if (sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES" || sectionTitleKey === "PROJECTS") {
          sectionObject = { title: sectionTitleKey, subsections: (sectionTitleKey === "PROJECTS" ? [{ name: "General Projects", items: []}] : []) };
      } else {
          sectionObject = { title: sectionTitleKey, items: [] };
      }
      resumeOutput.sections.push(sectionObject);
  }

  if (["EXPERIENCE", "EDUCATION", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"].includes(sectionTitleKey)) {
    if (processedItems.length > 0) sectionObject.items.push(...processedItems);
  } else if (sectionTitleKey === "TECHNICAL SKILLS & CERTIFICATES") {
    const subsectionsMap = new Map();
    processedItems.forEach(item => {
      const category = item.categoryName || "Uncategorized";
      if (!subsectionsMap.has(category)) subsectionsMap.set(category, { name: category, items: [] });
      const categoryItemsArray = subsectionsMap.get(category).items;
      // Differentiate between skills and certifications based on presence of issuer/issueDate
      if (item.issuer || item.issueDate) {
        categoryItemsArray.push({ name: item.skillItem || "", issuer: item.issuer || "", issueDate: item.issueDate || "" });
      } else {
        categoryItemsArray.push({ skill: item.skillItem || "", details: item.details || "" });
      }
    });
    if (subsectionsMap.size > 0) sectionObject.subsections = Array.from(subsectionsMap.values());
  } else if (sectionTitleKey === "PROJECTS") {
    const projectItemsStructured = processedItems.map(item => ({
      projectName: item.projectName || "", organization: item.organization || "", role: item.role || "",
      startDate: item.startDate || "", endDate: item.endDate || "",
      descriptionBullets: item.descriptionBullets || [],
      technologies: item.technologies || [],
      impact: item.impact || "",
      futureDevelopment: item.futureDevelopment || "",
      githubLinks: (item.githubName1 && item.githubUrl1) ? [{ name: item.githubName1, url: item.githubUrl1 }] : [],
      // Handle other GitHub links if sheet structure supports GitHubName2/URL2 etc.
      subProjectsOrRoles: [] // Kept simple as per assumption of flat structure from single sheet row
    }));
    // Assuming 'PROJECTS' section uses a default 'General Projects' subsection if not otherwise categorized.
    // Ensure the subsection structure is correctly managed.
    if (projectItemsStructured.length > 0) {
        if (!sectionObject.subsections || sectionObject.subsections.length === 0) {
             sectionObject.subsections = [{ name: "General Projects", items: [] }];
        }
        // Find or ensure 'General Projects' subsection exists if that's the target
        let generalProjectsSub = sectionObject.subsections.find(sub => sub.name === "General Projects");
        if (!generalProjectsSub) {
            generalProjectsSub = { name: "General Projects", items: [] };
            sectionObject.subsections.push(generalProjectsSub);
        }
        generalProjectsSub.items.push(...projectItemsStructured);
    }
  }
}
