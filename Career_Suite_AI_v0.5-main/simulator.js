/**
 * Job Application Email Parser Simulator
 */

// Mock data storage
let mockSheet = [];
let mockLabels = {
  "JobAppToProcess": { threads: [] },
  "TestAppToProcess": { threads: [] }
};

// Configuration
const SHEET_NAME = "Applications";
const GMAIL_LABEL_TO_PROCESS = "JobAppToProcess";
const GMAIL_LABEL_APPLIED_AFTER_PROCESSING = "TestAppToProcess";

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
    this.labels = [mockLabels["JobAppToProcess"]]; // Start with the "to process" label
    this.messages = [new MockMessage(id, subject, from, body, date)];
  }

  getId() { return this.id; }

  getLabels() {
    return this.labels.map(labelObj => ({
        getName: () => Object.keys(mockLabels).find(key => mockLabels[key] === labelObj)
    }));
  }

  getMessages() { return this.messages; }

  removeLabel(labelObjToRemove) {
    const labelNameToRemove = Object.keys(mockLabels).find(key => mockLabels[key] === labelObjToRemove);
    this.labels = this.labels.filter(l => l !== labelObjToRemove);
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
    this.id = id;
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
            const values = [];
            const endRow = row + numRows - 1;
            const endCol = col + numCols - 1;

            for (let i = row; i <= endRow; i++) {
              const rowValues = [];
              const sheetRowIndex = i - 1;
              if (sheetRowIndex < mockSheet.length) {
                for (let j = col; j <= endCol; j++) {
                  const sheetColIndex = j - 1;
                  // Check if the column exists in the row, otherwise push empty string
                  if (sheetColIndex >= 0 && sheetColIndex < mockSheet[sheetRowIndex].length) {
                    rowValues.push(mockSheet[sheetRowIndex][sheetColIndex]);
                  } else {
                    rowValues.push(""); // Handle columns outside the current row length but within numCols request
                  }
                }
              } else {
                 // Handle rows outside the current sheet length by adding empty rows of correct width
                 rowValues.push(...Array(numCols).fill(""));
              }
              values.push(rowValues);
            }
            return {
              getValues: () => values,
              setValue: (val) => { console.log(`(MockRange.setValue called with ${val} - not used for sheet update)`); }
            };
          },
          appendRow: (rowData) => {
            const headerLength = mockSheet[0]?.length || 10;
            const paddedRowData = Array(headerLength).fill("");
            rowData.forEach((val, index) => {
              if (index < headerLength) {
                paddedRowData[index] = val !== undefined ? val : "";
              }
            });
            mockSheet.push(paddedRowData);
            console.log(`SHEET: Appended row (new length ${mockSheet.length}): ${paddedRowData.map(v => typeof v === 'string' ? v.substring(0, 30) : v).join(', ')}...`); // Log truncated data
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
    for (const labelName in mockLabels) {
      const thread = mockLabels[labelName].threads.find(t => t.id === id);
      if (thread) {
        //console.log(`GMAIL: Found thread by ID: ${id} in label "${labelName}"`); // Can be noisy, uncomment if needed
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


// --- Test Data Setup ---

// Add several pre-existing entries to the mock sheet to test updates
function addPreExistingApplications() {
    // Ensure mockSheet is initialized with headers first
    if (mockSheet.length === 0) initMockSheet();
    console.log("SIM: Adding pre-existing application entries to mock sheet...");

    const headerLength = mockSheet[0]?.length || 10; // Get expected number of columns

    const applications = [
        // Entry 1: TechCorp (This one will be targeted by the rejection email)
        {
            processed: new Date(2025, 3, 20), emailDate: new Date(2025, 3, 15), platform: "LinkedIn",
            company: "TechCorp", title: "Senior Developer", status: DEFAULT_STATUS, lastUpdate: "",
            subject: "Your application for Senior Developer at TechCorp", link: "...", emailId: "existing_techcorp_1"
        },
        // Entry 2: Beta Corp (This one will be targeted by an interview email)
        {
            processed: new Date(2025, 3, 21), emailDate: new Date(2025, 3, 16), platform: "Indeed",
            company: "Beta Corp", title: "Data Analyst", status: DEFAULT_STATUS, lastUpdate: "",
            subject: "Application Received - Data Analyst at Beta Corp", link: "...", emailId: "existing_betacorp_1"
        },
        // Entry 3: Gamma Inc (This one will be targeted by an assessment email)
        {
            processed: new Date(2025, 3, 22), emailDate: new Date(2025, 3, 17), platform: "Other",
            company: "Gamma Inc", title: "Backend Engineer", status: DEFAULT_STATUS, lastUpdate: "",
            subject: "Re: Your Gamma Inc Application", link: "...", emailId: "existing_gammainc_1"
        },
         // Entry 4: Delta Solutions (This one will be targeted by an offer email)
        {
            processed: new Date(2025, 3, 23), emailDate: new Date(2025, 3, 18), platform: "Glassdoor",
            company: "Delta Solutions", title: "Project Manager", status: INTERVIEW_STATUS, // Maybe they already interviewed
            lastUpdate: new Date(2025, 3, 22), // Simulate a previous update
            subject: "Interview Follow-up - Delta Solutions", link: "...", emailId: "existing_deltasol_1"
        }
    ];

    applications.forEach(app => {
        const entry = [];
        entry[PROCESSED_TIMESTAMP_COL - 1] = app.processed;
        entry[EMAIL_DATE_COL - 1] = app.emailDate;
        entry[PLATFORM_COL - 1] = app.platform;
        entry[COMPANY_COL - 1] = app.company;
        entry[JOB_TITLE_COL - 1] = app.title;
        entry[STATUS_COL - 1] = app.status;
        entry[LAST_UPDATE_DATE_COL - 1] = app.lastUpdate;
        entry[EMAIL_SUBJECT_COL - 1] = app.subject;
        entry[EMAIL_LINK_COL - 1] = app.link.replace("...", `https://mail.google.com/mail/u/0/#inbox/${app.emailId}`);
        entry[EMAIL_ID_COL - 1] = app.emailId;

        // Pad the entry if necessary
        while (entry.length < headerLength) {
            entry.push("");
        }
        mockSheet.push(entry);
        console.log(`SIM: Added pre-existing entry for ${app.company} (${app.title})`);
    });
     console.log(`SIM: Mock sheet now has ${mockSheet.length} rows (including header).`);
}

// Create various test email cases, including updates for pre-existing apps
function createTestCases() {
    console.log("SIM: Creating test email threads...");
    // --- Emails for NEW applications ---

    // Test Case 1: New application email (Acme Inc) - Should create a new entry
    const newAppThread = new MockThread(
        "newapp_acme_1", // ID
        "Thank you for applying to Software Engineer at Acme Inc", // Subject
        "recruiting@acme.com", // From
        "Thank you for submitting your application for the Software Engineer position at Acme Inc...", // Body
        new Date(2025, 3, 25) // Date
    );
    mockLabels["JobAppToProcess"].threads.push(newAppThread);

    // --- Emails designed to UPDATE pre-existing applications ---

    // Test Case 2: Rejection email for TechCorp (matches pre-existing entry)
    const rejectionThread = new MockThread(
        "update_techcorp_reject_1", // New ID for this specific email
        "Update on your TechCorp application", // Subject (Parsable to TechCorp)
        "hr@techcorp.com", // From
        "Dear Applicant, \n\nThank you for your interest in TechCorp. Unfortunately, after careful consideration, we have decided not to move forward with your application for Senior Developer at this time. We wish you the best...\n\nRegards,\nTechCorp Hiring Team", // Body with rejection keywords
        new Date(2025, 3, 26) // Date
    );
    mockLabels["JobAppToProcess"].threads.push(rejectionThread);

    // Test Case 3: Interview invitation for Beta Corp (matches pre-existing entry)
    const interviewThread = new MockThread(
        "update_betacorp_interview_1", // New ID
        "Interview Invitation - Data Analyst at Beta Corp", // Subject (Parsable to Beta Corp & Title)
        "talent@betacorp.io", // From
        "Hello,\n\nWe were impressed with your profile and would like to schedule an interview with you regarding your application for the Data Analyst role at Beta Corp. Please let us know your availability...\n\nBest regards,\nRecruitment Team", // Body with interview keywords
        new Date(2025, 3, 27) // Date
    );
    mockLabels["JobAppToProcess"].threads.push(interviewThread);

    // Test Case 4: Assessment request for Gamma Inc (matches pre-existing entry)
    const assessmentThread = new MockThread(
        "update_gammainc_assess_1", // New ID
        "Next Steps: Online Assessment for Gamma Inc Application", // Subject (Parsable to Gamma Inc)
        "assessments@gamma-inc.net", // From
        "Hello candidate,\n\nAs the next step in our hiring process at Gamma Inc for the Backend Engineer role, we'd like you to complete an online coding assessment. Instructions will follow...\n\nThe assessment is timed.", // Body with assessment keywords
        new Date(2025, 3, 28) // Date
    );
    mockLabels["JobAppToProcess"].threads.push(assessmentThread);

    // Test Case 5: Offer letter for Delta Solutions (matches pre-existing entry)
    const offerThread = new MockThread(
        "update_deltasol_offer_1", // New ID
        "Congratulations! Job Offer from Delta Solutions", // Subject (Parsable to Delta Solutions)
        "offers@deltasolutions.com", // From
        "Dear Candidate,\n\nFollowing your recent interviews, we are pleased to offer you the position of Project Manager at Delta Solutions! Attached you'll find details of our offer of employment...\n\nWe look forward to welcoming you!", // Body with offer keywords
        new Date(2025, 3, 29) // Date
    );
    mockLabels["JobAppToProcess"].threads.push(offerThread);

    console.log(`SIM: Created ${mockLabels["JobAppToProcess"].threads.length} test email threads in label "${GMAIL_LABEL_TO_PROCESS}"`);
}


// Print simulation results
function printResults() {
  console.log("\n=== SIMULATION RESULTS ===");
  console.log(`Final spreadsheet rows: ${mockSheet.length}`);
  console.log("Spreadsheet Data (Company | Title | Status | Email ID | Last Update Date):");
  console.log("-".repeat(120));
  mockSheet.forEach((row, index) => {
    if (index === 0) {
      console.log("HEADERS: " + row.join(" | "));
    } else {
      const company = row[COMPANY_COL - 1] || 'N/A';
      const title = row[JOB_TITLE_COL - 1] || 'N/A';
      const status = row[STATUS_COL - 1] || 'N/A';
      const emailId = row[EMAIL_ID_COL - 1] || 'N/A';
      const lastUpdate = row[LAST_UPDATE_DATE_COL -1] ? new Date(row[LAST_UPDATE_DATE_COL -1]).toLocaleDateString() : ''; // Format date nicely
      console.log(`ROW ${index + 1}: ${company.padEnd(15)} | ${title.padEnd(20)} | ${status.padEnd(20)} | ${emailId.padEnd(25)} | ${lastUpdate}`);
    }
  });
  console.log("-".repeat(120));

  console.log("\nFinal Label State:");
  console.log(`- "${GMAIL_LABEL_TO_PROCESS}" Threads: ${mockLabels["JobAppToProcess"]?.threads?.length || 0}`);
  console.log(`- "${GMAIL_LABEL_APPLIED_AFTER_PROCESSING}" Threads: ${mockLabels["TestAppToProcess"]?.threads?.length || 0}`);
}


// --- Main Processing Logic ---

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
      const range = sheet.getRange(2, COMPANY_COL, lastRow - 1, EMAIL_ID_COL - COMPANY_COL + 1);
      const values = range.getValues();

      for (let i = 0; i < values.length; i++) {
        const rowNum = i + 2; // Sheet rows are 1-based, data starts at row 2
        const companyColIndex = COMPANY_COL - COMPANY_COL; // 0
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
            existingData[companyName].push({ row: rowNum, emailId: emailId, title: jobTitleVal });
        }
      }
    } catch (e) {
        console.error("PROCESS: Error reading existing sheet data:", e);
        return;
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
    const currentThread = MockGmailApp.getThreadById(thread.getId());
    if (!currentThread || !currentThread.getLabels().some(l => l.getName() === GMAIL_LABEL_TO_PROCESS)) {
        console.log(`PROCESS: Skipping thread ${thread.getId()}, label "${GMAIL_LABEL_TO_PROCESS}" likely already removed.`);
        continue;
    }

    const messages = currentThread.getMessages();
    // Process only the first/latest message in the thread for simplicity in this simulation
    const message = messages[0];
    if (!message) continue; // Should not happen with current mock setup

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
      const emailLink = `https://mail.google.com/mail/u/0/#inbox/${messageId}`;
      const processedTimestamp = new Date();

      let platform = DEFAULT_PLATFORM;
      let company = MANUAL_REVIEW_NEEDED;
      let jobTitle = MANUAL_REVIEW_NEEDED;
      let status = DEFAULT_STATUS;
      let isUpdate = false;
      let targetRow = -1;

      // Basic platform detection
      if (sender.includes("linkedin")) platform = "LinkedIn";
      else if (sender.includes("indeed")) platform = "Indeed";
      else if (sender.includes("glassdoor")) platform = "Glassdoor";
      else if (sender.includes("ziprecruiter")) platform = "ZipRecruiter";
      else if (sender.includes("google")) platform = "Google";

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
                   // Make this more specific to avoid overly broad matches
                   match = emailSubject.match(/Next Steps:.*for\s+(.+?)\s+Application/i);
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
       // --- Temporarily disabling suffix removal for better simulation matching ---
       /*
       preliminaryCompany = preliminaryCompany
           .replace(/inc\.?,?/i, '')
           .replace(/llc\.?,?/i, '')
           .replace(/limited\.?,?/i, '')
           .replace(/corporation\.?,?/i, '')
           .replace(/corp\.?,?/i, '')
           .trim();
       */
        preliminaryCompany = preliminaryCompany.trim(); // Still trim whitespace
 
 
       // If title wasn't found but company was, check common patterns like "[Company] - [Title]"
       if (preliminaryTitle === MANUAL_REVIEW_NEEDED && preliminaryCompany !== MANUAL_REVIEW_NEEDED) {
            // Ensure company name is safe for regex (escape special characters if necessary)
            const escapedCompany = preliminaryCompany.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            match = emailSubject.match(new RegExp(escapedCompany + '\\s*[-–—]\\s*(.*)', 'i'));
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
        // Keep the existing job title from the sheet unless the update email *clearly* specifies a different one
        jobTitle = latestMatch.title || preliminaryTitle || MANUAL_REVIEW_NEEDED; // Use existing, then parsed, then fallback
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
           const rowIndexToUpdate = targetRow - 1;
           mockSheet[rowIndexToUpdate][STATUS_COL - 1] = newStatus;
           mockSheet[rowIndexToUpdate][LAST_UPDATE_DATE_COL - 1] = processedTimestamp;
           mockSheet[rowIndexToUpdate][PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp;
           mockSheet[rowIndexToUpdate][EMAIL_SUBJECT_COL - 1] = emailSubject;
           mockSheet[rowIndexToUpdate][EMAIL_LINK_COL - 1] = emailLink;
           mockSheet[rowIndexToUpdate][EMAIL_ID_COL - 1] = messageId;
           // Optionally update Title if the new email provided one and the old one was Manual Review
           if (jobTitle !== MANUAL_REVIEW_NEEDED && mockSheet[rowIndexToUpdate][JOB_TITLE_COL - 1] === MANUAL_REVIEW_NEEDED) {
               mockSheet[rowIndexToUpdate][JOB_TITLE_COL - 1] = jobTitle;
               console.log(`PROCESS: ---> Also updated Job Title on row ${targetRow} to "${jobTitle}"`);
           }

           console.log(`PROCESS: ---> UPDATED mockSheet Row ${targetRow}: Status='${newStatus}', UpdateDate set, Email Info Updated.`);
           updatedCount++;
           existingEmailIds.add(messageId); // Add ID to prevent re-processing

        } else if (targetRow > 0 && targetRow <= mockSheet.length) {
           // Even if no status change, update metadata
           const rowIndexToUpdate = targetRow - 1;
           mockSheet[rowIndexToUpdate][LAST_UPDATE_DATE_COL - 1] = processedTimestamp;
           mockSheet[rowIndexToUpdate][PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp;
           mockSheet[rowIndexToUpdate][EMAIL_SUBJECT_COL - 1] = emailSubject;
           mockSheet[rowIndexToUpdate][EMAIL_LINK_COL - 1] = emailLink;
           mockSheet[rowIndexToUpdate][EMAIL_ID_COL - 1] = messageId;
            if (jobTitle !== MANUAL_REVIEW_NEEDED && mockSheet[rowIndexToUpdate][JOB_TITLE_COL - 1] === MANUAL_REVIEW_NEEDED) {
               mockSheet[rowIndexToUpdate][JOB_TITLE_COL - 1] = jobTitle;
               console.log(`PROCESS: ---> Also updated Job Title on row ${targetRow} to "${jobTitle}"`);
           }

           console.log(`PROCESS: ---> UPDATED mockSheet Row ${targetRow}: No status change detected, but updated timestamp/email/title info.`);
           existingEmailIds.add(messageId); // Still mark as processed

        } else {
          console.log(`PROCESS: ERROR: Update identified but target row invalid (${targetRow}), or row index out of bounds, or no new status found. Not updating sheet.`);
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
                  const bodyLower = plainBody.toLowerCase();
                  let bodyMatch;
                  if (company === MANUAL_REVIEW_NEEDED) {
                      bodyMatch = bodyLower.match(/company:\s*(.+)/i) || bodyLower.match(/applying to\s*.*\s*at\s*(.+)/i);
                      if (bodyMatch) company = bodyMatch[1].split('\n')[0].trim();
                  }
                  if (jobTitle === MANUAL_REVIEW_NEEDED) {
                      bodyMatch = bodyLower.match(/position:\s*(.+)/i) || bodyLower.match(/role:\s*(.+)/i) || bodyLower.match(/job title:\s*(.+)/i);
                       if (bodyMatch) jobTitle = bodyMatch[1].split('\n')[0].trim();
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
        if (!company || company.trim() === '' || company.trim().toLowerCase() === 'n/a') company = MANUAL_REVIEW_NEEDED;
        if (!jobTitle || jobTitle.trim() === '' || jobTitle.trim().toLowerCase() === 'n/a') jobTitle = MANUAL_REVIEW_NEEDED;


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

      // Manage Labels (Applies to both Update and New paths if successful or potentially handled)
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
    }
    // --- End Process Email ---
    // Break from message loop - only process one message per thread per run in this simulation
    // break; // Commented out - let it process all messages if needed, though mock only has one per thread

  } // End thread loop

  console.log(`PROCESS: Finished processing run. Added ${processedCount} new entries, updated ${updatedCount} existing entries.`);
}


// --- Simulation Execution ---

// Main simulation function
function runSimulation() {
  console.log("============================================================");
  console.log("STARTING JOB APPLICATION EMAIL PARSER SIMULATION");
  console.log("============================================================");

  // Initialize mock environment
  global.Logger = { log: console.log };
  global.SpreadsheetApp = MockSpreadsheetApp;
  global.GmailApp = MockGmailApp;

  // 1. Initialize test sheet (Headers only)
  initMockSheet();

  // 2. Add pre-existing applications to simulate prior runs
  addPreExistingApplications(); // Use the new function name here

  // 3. Create mock email threads (including updates for pre-existing ones)
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