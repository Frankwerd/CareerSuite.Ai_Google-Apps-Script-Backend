/**
 * Job Application Email Parser Simulator
 */

// Mock data storage
let mockSheet = [];
let mockLabels = {
  "JobAppToProcess": { threads: [] },
  "TestAppToProcess": { threads: [] } // Changed label name slightly as per original intent
};

// Configuration
const SHEET_NAME = "Applications";
const GMAIL_LABEL_TO_PROCESS = "JobAppToProcess";
const GMAIL_LABEL_APPLIED_AFTER_PROCESSING = "TestAppToProcess"; // Changed label name slightly as per original intent

// Column Indices (1-based)
const PROCESSED_TIMESTAMP_COL = 1; const EMAIL_DATE_COL = 2; const PLATFORM_COL = 3; const COMPANY_COL = 4;
const JOB_TITLE_COL = 5; const STATUS_COL = 6; const LAST_UPDATE_DATE_COL = 7;
const EMAIL_SUBJECT_COL = 8; const EMAIL_LINK_COL = 9; const EMAIL_ID_COL = 10;

// Status Values
const DEFAULT_STATUS = "Applied";
const REJECTED_STATUS = "Rejected";
const ACCEPTED_STATUS = "Offer/Accepted";
const INTERVIEW_STATUS = "Interview Scheduled";
const ASSESSMENT_STATUS = "Assessment/Screening";
const MANUAL_REVIEW_NEEDED = "N/A - Manual Review Needed";

// Default Platform Value
const DEFAULT_PLATFORM = "Other";

// Status Keywords
const REJECTION_KEYWORDS = ["unfortunately", "regret to inform", "not moving forward", "will not be proceeding",
                          "selected other candidate", "position has been filled", "pursuing other applicant",
                          "after careful consideration", "won't be moving forward", "application was unsuccessful"];
const ACCEPTANCE_KEYWORDS = ["pleased to offer", "offer of employment", "job offer", "congratulations",
                           "welcome aboard", "extend an offer", "next steps include your offer"];
const INTERVIEW_KEYWORDS = ["invitation to interview", "schedule an interview", "interview request",
                          "like to schedule a time", "speak with you about your application",
                          "next step is an interview", "schedule time to chat"];
const ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "next step is a test", "online test",
                           "screening task", "technical screen"];

// Initialize the mock spreadsheet with headers
function initMockSheet() {
  mockSheet = [
    ["Processed Timestamp", "Email Date", "Platform", "Company", "Job Title", "Status", "Last Update Date",
     "Email Subject", "Email Link", "Email ID"]
  ];
  console.log("Mock sheet initialized with headers");
  return mockSheet;
}

// --- Mock Google Service Objects ---

// Mock Gmail thread and message classes
class MockThread {
  constructor(id, subject, from, body, date) {
    this.id = id;
    // Start with the "to process" label
    this.labels = [mockLabels["JobAppToProcess"]];
    this.messages = [new MockMessage(id, subject, from, body, date)];
  }

  getId() { return this.id; }

  getLabels() {
    // Return mock label objects that have a getName method
    return this.labels.map(labelObj => ({
        getName: () => Object.keys(mockLabels).find(key => mockLabels[key] === labelObj)
    }));
  }

  getMessages() { return this.messages; }

  removeLabel(labelObjToRemove) {
    const labelNameToRemove = Object.keys(mockLabels).find(key => mockLabels[key] === labelObjToRemove);
    this.labels = this.labels.filter(l => l !== labelObjToRemove);
    // Remove from the global label's threads collection too
    if (labelNameToRemove && mockLabels[labelNameToRemove]) {
        mockLabels[labelNameToRemove].threads = mockLabels[labelNameToRemove].threads.filter(t => t.id !== this.id);
        console.log(`LABEL: Removed thread ${this.id} from label "${labelNameToRemove}"`);
    }
    console.log(`THREAD ${this.id}: Removed label "${labelNameToRemove}"`);
  }

  addLabel(labelObjToAdd) {
    const labelNameToAdd = Object.keys(mockLabels).find(key => mockLabels[key] === labelObjToAdd);
    if (!this.labels.includes(labelObjToAdd)) {
      this.labels.push(labelObjToAdd);
      // Add to the global label's threads collection
      if (labelNameToAdd && mockLabels[labelNameToAdd]) {
        if (!mockLabels[labelNameToAdd].threads.some(t => t.id === this.id)) {
            mockLabels[labelNameToAdd].threads.push(this);
            console.log(`LABEL: Added thread ${this.id} to label "${labelNameToAdd}"`);
        }
      }
      console.log(`THREAD ${this.id}: Added label "${labelNameToAdd}"`);
    }
  }
}

class MockMessage {
  constructor(id, subject, from, body, date) {
    this.id = id; // Use thread ID as message ID for simplicity in this mock
    this.subject = subject;
    this.from = from;
    this.body = body;
    this.date = date || new Date();
  }
  getId() { return this.id; }
  getSubject() { return this.subject; }
  getFrom() { return this.from; }
  getPlainBody() { return this.body; }
  getDate() { return this.date; }
}

// Mock SpreadsheetApp
const MockSpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
      if (name === SHEET_NAME) {
        return {
          // --- Mock Sheet Methods ---
          getLastRow: () => mockSheet.length,
          getRange: (row, col, numRows = 1, numCols = 1) => {
            // Basic implementation to get values, sufficient for pre-loading
            const values = [];
            const endRow = row + numRows - 1;
            const endCol = col + numCols - 1;

            for (let i = row; i <= endRow; i++) {
              const rowValues = [];
              const sheetRowIndex = i - 1;
              if (sheetRowIndex < mockSheet.length) {
                for (let j = col; j <= endCol; j++) {
                  const sheetColIndex = j - 1;
                  if (sheetColIndex < mockSheet[sheetRowIndex].length) {
                    rowValues.push(mockSheet[sheetRowIndex][sheetColIndex]);
                  } else {
                    rowValues.push(""); // Handle out-of-bounds columns
                  }
                }
              } else {
                 // Handle out-of-bounds rows by adding empty rows of correct width
                 rowValues.push(...Array(numCols).fill(""));
              }
              values.push(rowValues);
            }
            return {
              getValues: () => values,
              // setValue is not directly used by the script's append/update logic
              // but could be added here if needed for other tests.
              setValue: (val) => { console.log(`(MockRange.setValue called with ${val} - not implemented for sheet update)`); }
            };
          },
          appendRow: (rowData) => {
            // Ensure rowData has the same number of columns as the header
            const headerLength = mockSheet[0]?.length || 10; // Default to 10 if header missing
            const paddedRowData = Array(headerLength).fill("");
            rowData.forEach((val, index) => {
              if (index < headerLength) {
                paddedRowData[index] = val !== undefined ? val : "";
              }
            });
            mockSheet.push(paddedRowData);
            console.log(`SHEET: Appended row: ${paddedRowData.join(', ')}`);
          }
          // --- End Mock Sheet Methods ---
        };
      }
      console.error(`SHEET: Mock sheet "${name}" not found.`);
      return null;
    }
  })
};

// Mock GmailApp
const MockGmailApp = {
  getUserLabelByName: (name) => {
    console.log(`GMAIL: Getting label by name: "${name}"`);
    return mockLabels[name] || null;
  },
  getThreadById: (id) => {
    // Search across all label thread lists
    for (const labelName in mockLabels) {
      const thread = mockLabels[labelName].threads.find(t => t.id === id);
      if (thread) {
        console.log(`GMAIL: Found thread by ID: ${id} in label "${labelName}"`);
        return thread;
      }
    }
    console.log(`GMAIL: Thread ID ${id} not found.`);
    return null;
  }
};

// Helper function to get or create a label (mock implementation)
function getOrCreateLabel(labelName) {
  if (!mockLabels[labelName]) {
    console.log(`LABEL: Creating new mock label: "${labelName}"`);
    mockLabels[labelName] = { threads: [] };
  }
  return mockLabels[labelName];
}

// --- End Mock Google Service Objects ---


// Create various test email cases
function createTestCases() {
  // Test Case 1: New application email (Acme Inc)
  const newAppThread = new MockThread(
    "newapp123", // ID
    "Thank you for applying to Software Engineer at Acme Inc", // Subject
    "recruiting@acme.com", // From
    "Thank you for submitting your application for the Software Engineer position at Acme Inc. We are reviewing your qualifications and will be in touch if there's a match.", // Body
    new Date(2025, 3, 25) // Date (Month is 0-indexed, so 3 is April)
  );
  mockLabels["JobAppToProcess"].threads.push(newAppThread);

  // Test Case 2: Rejection email for existing company (TechCorp - should update existing entry)
  const rejectionThread = new MockThread(
    "reject456", // ID
    "Update on your TechCorp application", // Subject
    "hr@techcorp.com", // From
    "Dear Applicant, \n\nThank you for your interest in TechCorp. Unfortunately, we have decided not to move forward with your application for Senior Developer at this time. We wish you the best in your job search.\n\nRegards,\nTechCorp Hiring Team", // Body
    new Date(2025, 3, 26) // Date
  );
  mockLabels["JobAppToProcess"].threads.push(rejectionThread);

  // Test Case 3: Interview invitation (NewStartup)
  const interviewThread = new MockThread(
    "interview789", // ID
    "Interview Invitation - Product Manager at NewStartup", // Subject
    "talent@newstartup.io", // From
    "Hello,\n\nWe'd like to schedule an interview with you regarding your application for the Product Manager role at NewStartup. Please let us know your availability for next week.\n\nBest regards,\nRecruitment Team", // Body
    new Date(2025, 3, 27) // Date
  );
  mockLabels["JobAppToProcess"].threads.push(interviewThread);

  // Test Case 4: Assessment request (BigTech)
  const assessmentThread = new MockThread(
    "assess101", // ID
    "Next Steps: Coding Assessment for BigTech Application", // Subject
    "no-reply@bigtech.com", // From
    "Hello candidate,\n\nAs part of our hiring process at BigTech for the Cloud Engineer role, we'd like you to complete an online coding assessment. You'll receive a separate email with instructions.\n\nThe assessment should take approximately 90 minutes.", // Body
    new Date(2025, 3, 28) // Date
  );
  mockLabels["JobAppToProcess"].threads.push(assessmentThread);

  // Test Case 5: Offer letter (Dream Company)
  const offerThread = new MockThread(
    "offer202", // ID
    "Congratulations! Job Offer from Dream Company", // Subject
    "offers@dreamcompany.org", // From
    "Dear Candidate,\n\nWe are pleased to offer you the position of Senior Software Architect at Dream Company. Attached you'll find details of our offer including compensation and benefits.\n\nWe look forward to welcoming you to our team!", // Body
    new Date(2025, 3, 29) // Date
  );
  mockLabels["JobAppToProcess"].threads.push(offerThread);

  console.log(`SIM: Created ${mockLabels["JobAppToProcess"].threads.length} test email threads in label "${GMAIL_LABEL_TO_PROCESS}"`);
}

// Print simulation results
function printResults() {
  console.log("\n=== SIMULATION RESULTS ===");
  console.log(`Final spreadsheet rows: ${mockSheet.length}`);
  console.log("Spreadsheet Data (Company | Title | Status | Email ID):");
  console.log("-".repeat(100));
  mockSheet.forEach((row, index) => {
    if (index === 0) {
      console.log("HEADERS: " + row.join(" | "));
    } else {
      const company = row[COMPANY_COL - 1] || 'N/A';
      const title = row[JOB_TITLE_COL - 1] || 'N/A';
      const status = row[STATUS_COL - 1] || 'N/A';
      const emailId = row[EMAIL_ID_COL - 1] || 'N/A';
      console.log(`ROW ${index + 1}: ${company} | ${title} | ${status} | ${emailId}`);
    }
  });
  console.log("-".repeat(100));

  console.log("\nFinal Label State:");
  console.log(`- "${GMAIL_LABEL_TO_PROCESS}" Threads: ${mockLabels["JobAppToProcess"]?.threads?.length || 0}`);
  console.log(`- "${GMAIL_LABEL_APPLIED_AFTER_PROCESSING}" Threads: ${mockLabels["TestAppToProcess"]?.threads?.length || 0}`);
}

/**
 * Main function to process emails - adapted from original code
 */
function processJobApplicationEmails() {
  const ss = MockSpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log(`PROCESS: Sheet "${SHEET_NAME}" not found.`); return; }

  const processingLabel = MockGmailApp.getUserLabelByName(GMAIL_LABEL_TO_PROCESS);
  if (!processingLabel || !processingLabel.threads) { console.log(`PROCESS: Label "${GMAIL_LABEL_TO_PROCESS}" not found or has no 'threads' property.`); return; }

  const processedLabel = getOrCreateLabel(GMAIL_LABEL_APPLIED_AFTER_PROCESSING);

  // Pre-load existing data for faster lookup
  const lastRow = sheet.getLastRow();
  const existingData = {}; // Structure: {'lowercase company': [{row: R, emailId: EID, title: T}, ...], ...}
  const existingEmailIds = new Set(); // Still need this for initial skip

  if (lastRow >= 2) { // Check if there's data beyond the header row
    try {
      // Fetch data range (Company to Email ID)
      const range = sheet.getRange(2, COMPANY_COL, lastRow - 1, EMAIL_ID_COL - COMPANY_COL + 1);
      const values = range.getValues(); // Gets a 2D array

      for (let i = 0; i < values.length; i++) {
        const rowNum = i + 2; // Sheet rows are 1-based, data starts at row 2
        // Calculate correct indices within the fetched 'values' sub-array
        const companyColIndex = COMPANY_COL - COMPANY_COL; // Should be 0
        const emailIdColIndex = EMAIL_ID_COL - COMPANY_COL;
        const jobTitleColIndex = JOB_TITLE_COL - COMPANY_COL;

        const companyName = values[i][companyColIndex]?.toString().trim().toLowerCase() || "";
        const emailId = values[i][emailIdColIndex]?.toString().trim() || "";
        const jobTitleVal = values[i][jobTitleColIndex]?.toString().trim() || "";

        if (emailId) {
          existingEmailIds.add(emailId);
        }
        if (companyName && companyName !== 'n/a' && companyName !== MANUAL_REVIEW_NEEDED.toLowerCase()) {
            if (!existingData[companyName]) {
                existingData[companyName] = [];
            }
            // Store row number, email ID, and title for potential updates
            existingData[companyName].push({ row: rowNum, emailId: emailId, title: jobTitleVal });
        }
      }
    } catch (e) {
        console.error("PROCESS: Error reading existing sheet data:", e);
        return; // Stop processing if we can't read the sheet reliably
    }
  }
  console.log(`PROCESS: Pre-loaded data for ${Object.keys(existingData).length} companies. Found ${existingEmailIds.size} existing email IDs.`);

  // Get a copy of the threads array to avoid issues if modifying during iteration
  const threadsToProcess = [...processingLabel.threads];
  console.log(`PROCESS: Found ${threadsToProcess.length} threads with label "${GMAIL_LABEL_TO_PROCESS}".`);
  let processedCount = 0;
  let updatedCount = 0; // Track updates

  // Loop through Threads
  for (const thread of threadsToProcess) {
    // Double-check if the thread still has the processing label (it might have been removed by processing another message in the same thread)
    const currentThread = MockGmailApp.getThreadById(thread.getId()); // Get the latest state
    if (!currentThread || !currentThread.getLabels().some(l => l.getName() === GMAIL_LABEL_TO_PROCESS)) {
        console.log(`PROCESS: Skipping thread ${thread.getId()}, label "${GMAIL_LABEL_TO_PROCESS}" likely already removed.`);
        continue;
    }

    const messages = currentThread.getMessages();
    // Process messages, usually the latest one is most relevant, but loop allows finding *any* unlogged message
    for (const message of messages) {
      const messageId = message.getId();

      // Skip If Email ID Already Processed (based on pre-loaded sheet data)
      if (existingEmailIds.has(messageId)) {
        console.log(`PROCESS: Skipping already processed email ID: ${messageId} found in sheet.`);
        continue;
      }

      // --- Process New Email ---
      console.log(`PROCESS: ----- Processing Email ID: ${messageId} -----`);
      try {
        const emailDate = message.getDate();
        const emailSubject = message.getSubject();
        const sender = message.getFrom();
        const emailLink = `https://mail.google.com/mail/u/0/#inbox/${messageId}`; // Mock link
        const processedTimestamp = new Date();

        let platform = DEFAULT_PLATFORM;
        let company = MANUAL_REVIEW_NEEDED; // Default to manual review
        let jobTitle = MANUAL_REVIEW_NEEDED; // Default to manual review
        let status = DEFAULT_STATUS;
        let isUpdate = false;
        let targetRow = -1;

        // Basic platform detection
        if (sender.includes("linkedin")) platform = "LinkedIn";
        else if (sender.includes("indeed")) platform = "Indeed";
        else if (sender.includes("glassdoor")) platform = "Glassdoor";
        else if (sender.includes("ziprecruiter")) platform = "ZipRecruiter";
        else if (sender.includes("google")) platform = "Google"; // Example

        console.log(`PROCESS: Subject: "${emailSubject}" From: ${sender}`);

        // --- Improved Subject Parsing ---
        let preliminaryCompany = MANUAL_REVIEW_NEEDED;
        let preliminaryTitle = MANUAL_REVIEW_NEEDED;
        let match;

        // Order matters - try more specific patterns first

        // Pattern: "Thank you for applying to [Title] at [Company]" (Case 1)
        match = emailSubject.match(/applying to\s+(.+?)\s+at\s+([^-–—]+)/i);
        if (match) {
            preliminaryTitle = match[1].trim();
            preliminaryCompany = match[2].trim();
        } else {
             // Pattern: "Update on your [Company] application" (Case 2 - Simpler variation)
            match = emailSubject.match(/update on your\s+(.+?)\s+application/i);
            if (match) {
                preliminaryCompany = match[1].trim();
                // Title might be in body or previous entry for updates
            } else {
                // Pattern: "Interview Invitation - [Title] at [Company]" (Case 3)
                match = emailSubject.match(/Interview Invitation\s*[-–—]\s*(.+?)\s+at\s+([^-–—]+)/i);
                 if (match) {
                    preliminaryTitle = match[1].trim();
                    preliminaryCompany = match[2].trim();
                } else {
                    // Pattern: "Next Steps: [Whatever] for [Company] Application" (Case 4)
                    match = emailSubject.match(/for\s+(.+?)\s+Application/i);
                     if (match) {
                        preliminaryCompany = match[1].trim().replace(/^the\s/i, ''); // Attempt to capture company
                        // Title less likely in this subject format
                    } else {
                        // Pattern: "Job Offer from [Company]" (Case 5)
                         match = emailSubject.match(/Job Offer from\s+(.+)/i);
                         if (match) {
                             preliminaryCompany = match[1].trim();
                             // Title likely in body or previous entry
                         }
                         // Add more robust patterns here if needed
                    }
                }
            }
        }

        // Clean up common suffixes and leading 'the' if matched broadly
        preliminaryCompany = preliminaryCompany
            .replace(/inc\.?,?/i, '')
            .replace(/llc\.?,?/i, '')
            .replace(/limited\.?,?/i, '')
            .replace(/corporation\.?,?/i, '')
            .replace(/corp\.?,?/i, '')
            .trim();

        // If title wasn't found but company was, check common patterns like "[Company] - [Title]"
        if (preliminaryTitle === MANUAL_REVIEW_NEEDED && preliminaryCompany !== MANUAL_REVIEW_NEEDED) {
             match = emailSubject.match(new RegExp(preliminaryCompany + '\\s*[-–—]\\s*(.*)', 'i'));
             if (match && match[1]) {
                 preliminaryTitle = match[1].trim();
             }
        }
        // --- End Improved Subject Parsing ---

        console.log(`PROCESS: Parsed from Subject - Company: "${preliminaryCompany}", Title: "${preliminaryTitle}"`);

        // Core Logic: Check if this Company exists in Sheet Data
        const parsedCompanyLower = preliminaryCompany.toLowerCase();
        if (parsedCompanyLower && parsedCompanyLower !== MANUAL_REVIEW_NEEDED.toLowerCase() && existingData[parsedCompanyLower]) {
          // --- UPDATE PATH ---
          isUpdate = true;
          const companyMatches = existingData[parsedCompanyLower];
          // Find the most recent entry for this company (highest row number)
          const latestMatch = companyMatches.reduce((latest, current) =>
            (current.row > latest.row ? current : latest), companyMatches[0]);
          targetRow = latestMatch.row;
          company = preliminaryCompany; // Use the parsed company name
          // Keep the existing job title from the sheet unless the update email *clearly* specifies a different one (harder logic)
          jobTitle = latestMatch.title || MANUAL_REVIEW_NEEDED;
          console.log(`PROCESS: UPDATE identified for Company "${company}" (Found ${companyMatches.length} matches). Targeting last entry at row: ${targetRow}`);

          // Determine NEW STATUS based on email body
          let newStatus = null;
          try {
            const plainBody = message.getPlainBody();
            if (plainBody) {
              const bodyLower = plainBody.toLowerCase();
              // Status keyword check (order matters - offer is more specific than interview)
              if (ACCEPTANCE_KEYWORDS.some(k => bodyLower.includes(k))) newStatus = ACCEPTED_STATUS;
              else if (INTERVIEW_KEYWORDS.some(k => bodyLower.includes(k))) newStatus = INTERVIEW_STATUS;
              else if (ASSESSMENT_KEYWORDS.some(k => bodyLower.includes(k))) newStatus = ASSESSMENT_STATUS;
              else if (REJECTION_KEYWORDS.some(k => bodyLower.includes(k))) newStatus = REJECTED_STATUS;

              if(newStatus) console.log(`PROCESS: UPDATE determined new status from body: ${newStatus}`);
              else console.log(`PROCESS: UPDATE: No specific status keywords found in body.`);
            } else {
              console.log(`PROCESS: UPDATE: Body empty/unreadable for ${messageId}`);
            }
          } catch (bodyError) {
            console.log(`PROCESS: UPDATE: Error reading body: ${bodyError}`);
          }

          // Perform Sheet Update (Actual Modification of mockSheet)
          if (targetRow > 0 && targetRow <= mockSheet.length && newStatus) {
             // Update the row in the mockSheet array (adjust for 0-based index)
             const rowIndexToUpdate = targetRow - 1;
             mockSheet[rowIndexToUpdate][STATUS_COL - 1] = newStatus;
             mockSheet[rowIndexToUpdate][LAST_UPDATE_DATE_COL - 1] = processedTimestamp;
             mockSheet[rowIndexToUpdate][PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp; // Update processed time
             mockSheet[rowIndexToUpdate][EMAIL_SUBJECT_COL - 1] = emailSubject; // Update to latest email subject
             mockSheet[rowIndexToUpdate][EMAIL_LINK_COL - 1] = emailLink;     // Update to latest email link
             mockSheet[rowIndexToUpdate][EMAIL_ID_COL - 1] = messageId;       // Update to latest email ID

             console.log(`PROCESS: ---> UPDATED mockSheet Row ${targetRow}: Status='${newStatus}', UpdateDate set, Email Info Updated.`);
             updatedCount++;
             existingEmailIds.add(messageId); // Add ID to prevent re-processing if script runs again on same data

          } else if (targetRow > 0 && targetRow <= mockSheet.length) {
             // Even if no status change, update metadata (e.g., last contact)
             const rowIndexToUpdate = targetRow - 1;
             mockSheet[rowIndexToUpdate][LAST_UPDATE_DATE_COL - 1] = processedTimestamp;
             mockSheet[rowIndexToUpdate][PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp;
             mockSheet[rowIndexToUpdate][EMAIL_SUBJECT_COL - 1] = emailSubject;
             mockSheet[rowIndexToUpdate][EMAIL_LINK_COL - 1] = emailLink;
             mockSheet[rowIndexToUpdate][EMAIL_ID_COL - 1] = messageId;

             console.log(`PROCESS: ---> UPDATED mockSheet Row ${targetRow}: No status change detected, but updated timestamp/email info.`);
             existingEmailIds.add(messageId); // Still mark as processed

          } else {
            console.log(`PROCESS: ERROR: Update identified but target row invalid (${targetRow}), or row index out of bounds, or no new status found. Not updating sheet.`);
            // Decide if you still want to apply labels even if sheet update fails
          }

        } else {
          // --- NEW ENTRY PATH ---
          isUpdate = false;
          console.log(`PROCESS: NEW application entry identified for Company: "${preliminaryCompany}"`);
          company = preliminaryCompany;
          jobTitle = preliminaryTitle;
          status = DEFAULT_STATUS; // Default for new applications

          // Optionally try to extract more info from body for new entries (if subject parsing failed)
          if (company === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED) {
              try {
                  const plainBody = message.getPlainBody();
                  if (plainBody) {
                    // Very basic body parsing - enhance as needed
                    const bodyLower = plainBody.toLowerCase();
                    let bodyMatch;
                    if (company === MANUAL_REVIEW_NEEDED) {
                        bodyMatch = bodyLower.match(/company:\s*(.+)/i) || bodyLower.match(/applying to\s*.*\s*at\s*(.+)/i);
                        if (bodyMatch) company = bodyMatch[1].split('\n')[0].trim(); // Take first line after match
                    }
                    if (jobTitle === MANUAL_REVIEW_NEEDED) {
                        bodyMatch = bodyLower.match(/position:\s*(.+)/i) || bodyLower.match(/role:\s*(.+)/i) || bodyLower.match(/job title:\s*(.+)/i);
                         if (bodyMatch) jobTitle = bodyMatch[1].split('\n')[0].trim(); // Take first line after match
                    }
                    // Clean up again if extracted from body
                    company = company.replace(/inc\.?,?/i, '').replace(/llc\.?,?/i, '').trim();
                    console.log(`PROCESS: Parsed from Body - Company: "${company}", Title: "${jobTitle}"`);
                  }
              } catch (bodyError) {
                   console.log(`PROCESS: NEW: Error reading body for extra info: ${bodyError}`);
              }
          }

          // Final fallbacks if still N/A
          if (!company || company.trim() === '') company = MANUAL_REVIEW_NEEDED;
          if (!jobTitle || jobTitle.trim() === '') jobTitle = MANUAL_REVIEW_NEEDED;


          // Append New Row using mock sheet's appendRow
          const newRowData = [];
          newRowData[PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp;
          newRowData[EMAIL_DATE_COL - 1] = emailDate;
          newRowData[PLATFORM_COL - 1] = platform;
          newRowData[COMPANY_COL - 1] = company;
          newRowData[JOB_TITLE_COL - 1] = jobTitle;
          newRowData[STATUS_COL - 1] = status;
          newRowData[LAST_UPDATE_DATE_COL - 1] = ""; // Leave blank initially
          newRowData[EMAIL_SUBJECT_COL - 1] = emailSubject;
          newRowData[EMAIL_LINK_COL - 1] = emailLink;
          newRowData[EMAIL_ID_COL - 1] = messageId;

          sheet.appendRow(newRowData); // Use the mock appendRow method
          console.log(`PROCESS: ---> APPENDED New Row: P:${platform}, C:${company}, T:${jobTitle}, S:${status}, ID:${messageId}`);

          // Update caches for subsequent emails in the same run
          existingEmailIds.add(messageId);
          const newCompanyLower = company.toLowerCase();
          if (newCompanyLower !== MANUAL_REVIEW_NEEDED.toLowerCase()) {
              if (!existingData[newCompanyLower]) existingData[newCompanyLower] = [];
              existingData[newCompanyLower].push({ row: mockSheet.length, emailId: messageId, title: jobTitle }); // Use current length as new row number
          }
          processedCount++;
        }

        // Manage Labels (Applies to both Update and New paths if successful)
        // Only manage labels if the sheet operation (append/update) was potentially successful
        if (isUpdate && targetRow > 0 || !isUpdate) {
            console.log(`PROCESS: Applying label changes for thread ${currentThread.getId()} (Email ID: ${messageId})`);
            currentThread.removeLabel(processingLabel);
            currentThread.addLabel(processedLabel);
        } else {
            console.log(`PROCESS: Skipping label changes for thread ${currentThread.getId()} due to sheet update issue.`);
        }
        console.log(`PROCESS: ----- Finished Email ID: ${messageId} -----`);

      } catch (error) {
        console.error(`PROCESS: FATAL ERROR processing msg ID ${messageId}: ${error.stack || error}`);
        // Optionally add error handling like adding a specific "Error" label
      }

      // --- End Process New Email ---
      // If one message in a thread is processed, we typically stop processing others in that thread for this run
      // Break from message loop once an email ID is processed (new or update)
      break;
    } // End message loop
  } // End thread loop

  console.log(`PROCESS: Finished processing run. Added ${processedCount} new entries, updated ${updatedCount} existing entries.`);
}

// Add an existing entry to the mock sheet to test updates
function addExistingEntry() {
    // Ensure mockSheet is initialized with headers first
  if (mockSheet.length === 0) initMockSheet();

  const existingEntry = [];
  existingEntry[PROCESSED_TIMESTAMP_COL - 1] = new Date(2025, 3, 20); // Processed date
  existingEntry[EMAIL_DATE_COL - 1] = new Date(2025, 3, 15); // Email date
  existingEntry[PLATFORM_COL - 1] = "LinkedIn";            // Platform
  existingEntry[COMPANY_COL - 1] = "TechCorp";            // Company (Ensure case matches test case parsing)
  existingEntry[JOB_TITLE_COL - 1] = "Senior Developer";    // Job Title
  existingEntry[STATUS_COL - 1] = DEFAULT_STATUS;             // Status
  existingEntry[LAST_UPDATE_DATE_COL - 1] = "";                    // Last Update Date
  existingEntry[EMAIL_SUBJECT_COL - 1] = "Your application for Senior Developer at TechCorp"; // Subject
  existingEntry[EMAIL_LINK_COL - 1] = "https://mail.google.com/mail/u/0/#inbox/existingid1"; // Link
  existingEntry[EMAIL_ID_COL - 1] = "existingid1";          // Email ID

  // Pad the entry if necessary to match header length
  const headerLength = mockSheet[0]?.length || 10;
  while (existingEntry.length < headerLength) {
      existingEntry.push("");
  }

  mockSheet.push(existingEntry);
  console.log("SIM: Added pre-existing test entry for TechCorp to mock sheet");
}

// Main simulation function
function runSimulation() {
  console.log("============================================================");
  console.log("STARTING JOB APPLICATION EMAIL PARSER SIMULATION");
  console.log("============================================================");

  // Initialize mock environment
  // Map Logger.log to console.log if your script uses Logger.log extensively
  global.Logger = { log: console.log };
  // Provide mock SpreadsheetApp and GmailApp globally if script expects them there
  global.SpreadsheetApp = MockSpreadsheetApp;
  global.GmailApp = MockGmailApp;

  // 1. Initialize test sheet
  initMockSheet(); // Creates headers

  // 2. Add a pre-existing entry to test the update logic
  addExistingEntry();

  // 3. Create mock email threads and add them to the "JobAppToProcess" label
  createTestCases();

  console.log("\n=== RUNNING processJobApplicationEmails ===\n");
  // 4. Run the actual processing function against the mock environment
  processJobApplicationEmails();
  console.log("\n=== FINISHED processJobApplicationEmails ===\n");

  // 5. Display results (final state of mockSheet and labels)
  printResults();

  console.log("============================================================");
  console.log("SIMULATION COMPLETE");
  console.log("============================================================");
}

// Export the main function for index.js
module.exports = { runSimulation }; 