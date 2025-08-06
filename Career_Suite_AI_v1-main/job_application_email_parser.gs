/**
 * Script: Job Application Email Parser & Sheet Logger (Template)
 * Version: 8.1 - Template Version (Optimized Body Fetch, Refined Logging)
 * Date: April 29, 2024 (Modify as needed)
 *
 * Purpose: This script processes job application emails fetched from Gmail based on a specific label,
 *          parses them for details like Company, Job Title, and Status using keyword and pattern matching,
 *          and then updates or appends corresponding rows in the Google Sheet it is attached to.
 *
 * ============================================================================
 * --- HOW TO USE THIS SCRIPT ---
 * ============================================================================
 *
 * 1.  **Open Your Google Sheet:** Go to the Google Sheet you want to use for tracking job applications.
 *     (It's recommended to use the template provided alongside this script, as the column numbers below MUST match).
 *
 * 2.  **Open Script Editor:** In the Google Sheet menu, go to "Extensions" > "Apps Script".
 *
 * 3.  **Paste the Code:** Delete any default code in the editor (like `function myFunction() {}`)
 *     and paste this entire script's content.
 *
 * 4.  **Save the Project:** Click the floppy disk icon (Save project). Give it a descriptive name
 *     (e.g., "Job Application Email Parser").
 *
 * 5.  **CONFIGURE THE SCRIPT:** Carefully review and adjust the settings in the
 *     "USER CONFIGURATION SECTION" below. This is CRITICAL for the script to work correctly
 *     with your sheet and Gmail setup. Pay close attention to:
 *     *   `SHEET_NAME`: Must match your sheet tab name exactly.
 *     *   `GMAIL_LABEL_...`: Must match the Gmail labels you will use. Create these labels in Gmail if they don't exist.
 *     *   `Column Indices (...)`: MUST match the column letters in YOUR sheet (A=1, B=2, etc.).
 *
 * 6.  **Set Up Gmail Labels:** In your Gmail account, create the labels you defined in the configuration
 *     (e.g., "AppToProcess", "AppDONEProcess", "ManualReviewNeeded").
 *
 * 7.  **Run Manually & Authorize:**
 *     *   Refresh your Google Sheet page. A new menu item "Job Processor (v8.1)" should appear.
 *     *   Click "Job Processor (v8.1)" > "1. Process Emails Now".
 *     *   The first time you run it, Google will ask for authorization. Review the permissions
 *         (it needs access to your Gmail and Spreadsheets) and click "Allow".
 *     *   Label a test email in Gmail with your "AppToProcess" label and run again to verify it works.
 *
 * 8.  **Set Up Automatic Trigger (Optional but Recommended):**
 *     *   In the Script Editor window, click the "Triggers" icon (looks like a clock) on the left sidebar.
 *     *   Click "+ Add Trigger".
 *     *   Configure the trigger settings:
 *         *   Choose which function to run: `processJobApplicationEmails`
 *         *   Choose which deployment should run: `Head`
 *         *   Select event source: `Time-driven`
 *         *   Select type of time based trigger: `Hourly timer` (or your preference)
 *         *   Error notification settings: Choose how you want to be notified if it fails.
 *     *   Click "Save". You might need to authorize again.
 *     *   (Alternatively, use the "Job Processor (v8.1)" > "2. Setup/Verify Hourly Trigger" menu item in the Sheet,
 *        which tries to create a basic hourly trigger).
 *
 * 9.  **Start Labeling:** As relevant emails come in (or for past emails), apply your "AppToProcess"
 *     label in Gmail. The script will pick them up on its next run (manual or triggered).
 *
 * ============================================================================
 * --- USER CONFIGURATION SECTION ---
 * --- Review and adjust these settings carefully! ---
 * ============================================================================
 */

// --- General Settings ---
const DEBUG_MODE = false; // SET TO true FOR DETAILED LOGGING (useful for troubleshooting, noisy otherwise), false FOR NORMAL OPERATION

// --- Google Sheet Configuration ---
// <<< CRITICAL: Must match the exact name of the tab in YOUR Google Sheet where data is stored >>>
const SHEET_NAME = "Applications";                 // <<<--- EDIT THIS if your sheet tab is named differently

// --- Gmail Label Configuration ---
// <<< CRITICAL: These labels MUST exist in your Gmail. Create them if needed. Script will try to create if missing, but creation can fail. >>>
const GMAIL_LABEL_TO_PROCESS = "AppToProcess";         // Label for emails the script should scan and process
const GMAIL_LABEL_APPLIED_AFTER_PROCESSING = "AppDONEProcess"; // Label applied to emails successfully processed by the script
const GMAIL_LABEL_MANUAL_REVIEW = "ManualReviewNeeded";   // Label applied if the script couldn't parse details or encountered an error

// --- Column Index Configuration (IMPORTANT!) ---
// <<< CRITICAL: These numbers MUST match the actual column positions in YOUR Google Sheet. (Column A = 1, B = 2, C = 3, etc.) >>>
// <<< If you use the provided template sheet, these defaults should be correct. If you change the sheet layout, UPDATE THESE NUMBERS! >>>
const PROCESSED_TIMESTAMP_COL = 1; // A: Timestamp when the script processed the email (Script Execution Time)
const EMAIL_DATE_COL = 2;          // B: Date the email was received (Email's Date)
const PLATFORM_COL = 3;            // C: Platform/ATS (e.g., LinkedIn, Greenhouse, Other)
const COMPANY_COL = 4;             // D: Company Name
const JOB_TITLE_COL = 5;           // E: Job Title
const STATUS_COL = 6;              // F: Application Status (Applied, Rejected, Interview, etc.)
const LAST_UPDATE_DATE_COL = 7;    // G: Timestamp of the *email* being processed (reflects last application update)
const EMAIL_SUBJECT_COL = 8;       // H: Subject line of the processed email
const EMAIL_LINK_COL = 9;          // I: Direct link to the processed email in Gmail
const EMAIL_ID_COL = 10;           // J: Unique ID of the processed email message
const TOTAL_COLUMNS_IN_SHEET = 10; // Total number of columns the script interacts with. MUST be >= the highest column index used above.
                                   // <<<--- UPDATE THIS if you add/remove columns AND adjust the indices above >>>

// --- Status Values & Keywords ---
// <<< OPTIONAL: You can customize these status names and the keywords used to detect them >>>
const DEFAULT_STATUS = "Applied";                 // Status assigned when a new application email is processed without other status keywords
const REJECTED_STATUS = "Rejected";               // Status assigned when rejection keywords are found
const ACCEPTED_STATUS = "Offer/Accepted";         // Status assigned when acceptance keywords are found
const INTERVIEW_STATUS = "Interview Scheduled";   // Status assigned when interview keywords are found
const ASSESSMENT_STATUS = "Assessment/Screening"; // Status assigned when assessment keywords are found
const MANUAL_REVIEW_NEEDED = "N/A - Manual Review Needed"; // Value used if Company/Title extraction fails critically
const DEFAULT_PLATFORM = "Other";                 // Default platform if detection fails

// Keywords for Status Detection (Lowercase)
// <<< OPTIONAL: Expand these lists with more keywords/phrases for better accuracy based on emails you receive >>>
const REJECTION_KEYWORDS = [ "unfortunately", "regret to inform", "not moving forward", "will not be proceeding", "won't be moving forward","application was unsuccessful", "you have not been selected", "did not select you", "weren't selected","decline to move forward", "unable to offer you the position", "will not be advancing","decision has been made to not proceed", "will not be moving your application forward","application will not be given further consideration", "selected other candidate", "pursuing other applicant","other candidates", "better aligned candidates", "decided to proceed with other applicants", "have chosen another candidate","move forward with another candidate", "not select you for further consideration", "qualifications more closely match","other applicants were a better match", "focused on candidates whose experiences more closely align","position has been filled", "filled the role", "position is no longer available", "position is now closed", "role has been filled","after careful consideration", "thank you for your interest, however", "wish you all the best","wish you success in your job search", "encourage you to apply for future roles", "keep your resume on file","keep your profile on file", "best of luck", "unable"];
const ACCEPTANCE_KEYWORDS = [ "pleased to offer", "offer of employment", "job offer", "extend an offer", "thrilled to offer", "happy to extend an offer","extend a formal offer", "offer details", "letter of offer", "official offer letter", "pleased to extend","formal offer of employment", "conditional offer", "details on your offer", "congratulations", "welcome aboard","excited to have you join", "welcome to the team", "next steps include your offer", "finalizing the offer details","please find your offer attached", "was accepted", "next steps in the offer process", "compensation package","offer includes details on compensation", "background check initiated", "onboarding documents", "start date"];
const INTERVIEW_KEYWORDS = [ "invitation to interview", "schedule an interview", "interview request", "invited to interview", "like to schedule a time","schedule time to chat", "schedule a call", "schedule your interview", "set up an interview", "book an interview","interview coordination", "interview availability", "confirm a time", "suggest some times", "calendar invitation","booking link", "use this link to schedule", "speak with you about your application", "speak further","discussion regarding your application", "chat about your background", "follow-up conversation", "meet with","team would like to meet you", "invited to meet the team", "availability for a conversation", "preliminary interview","speak with the team", "meet the hiring manager", "next step is an interview", "next round", "proceed with interview process","interview loop", "proceeding to the next stage", "moving to the interview phase", "advancing your application","introductory call", "phone screen", "video interview", "virtual interview", "technical interview","behavioral interview", "panel interview", "hiring manager interview", "onsite interview"];
const ASSESSMENT_KEYWORDS = [ "assessment", "coding challenge", "online test", "screening task", "technical screen", "assignment", "coding exercise","technical evaluation", "complete an assessment", "take-home task", "problem-solving exercise", "practical test","technical assignment", "skills assessment", "skills test", "candidate assessment test", "online assessment","coding assessment", "technical task", "complete this assignment", "take-home assignment", "case study","required to complete", "invite you to take", "hackerrank", "codility", "next step is a test", "homework assignment","part of our process involves a"];

// --- Platform Keywords for Domain Part Matching (lowercase) ---
// <<< OPTIONAL: Add more domain keywords or adjust mappings for better platform detection >>>
// Format: "keyword_in_domain": "Platform Name To Display"
const PLATFORM_DOMAIN_KEYWORDS = { "linkedin": "LinkedIn", "indeed": "Indeed", "wellfound": "Wellfound", "angel": "Wellfound", "greenhouse": "Greenhouse", "lever": "Lever", "workday": "Workday", "icims": "iCIMS", "smartrecruiters": "SmartRecruiters", "ashby": "Ashby", "ashbyhq": "Ashby", "bamboohr": "BambooHR", "oraclecloud": "Oracle HCM", "taleo": "Taleo", "tal": "Taleo", "ziprecruiter": "ZipRecruiter" };

// --- Ignored Domains ---
// <<< OPTIONAL: Add email domains here (lowercase) that should generally NOT be considered the 'Company' name >>>
// E.g., Add domains of generic notification services or common email providers if they cause issues.
const IGNORED_DOMAINS = new Set([ 'greenhouse.io', 'lever.co', 'myworkday.com', 'myworkdayjobs.com', 'icims.com', 'oraclecloud.com', 'smartrecruiters.com', 'tal.net', 'taleo.net','bamboohr.com', 'ashbyhq.com','wellfound.com', 'angel.co','linkedin.com', 'indeed.com', 'ziprecruiter.com', 'glassdoor.com', 'hired.com','google.com', 'googlemail.com', 'gmail.com', 'outlook.com', 'hotmail.com', 'live.com', 'yahoo.com', 'aol.com', 'msn.com','fastmail.com', 'protonmail.com', 'mail.com','amazonses.com', 'sendgrid.net', 'sparkpostmail.com', 'mailgun.org' ]);

/**
 * ============================================================================
 * --- END OF USER CONFIGURATION SECTION ---
 * --- Generally, no need to modify code below this line unless extending functionality ---
 * ============================================================================
 */


// --- Helper Function: Get or Create Gmail Label ---
function getOrCreateLabel(labelName) {
  // Input validation for label name
  if (!labelName || typeof labelName !== 'string' || labelName.trim() === "") {
    Logger.log(`[ERROR] Invalid labelName received for getOrCreateLabel: "${labelName}" (Type: ${typeof labelName}). Check configuration.`);
    SpreadsheetApp.getUi().alert(`Script Configuration Error: Invalid Gmail label name provided: "${labelName}". Please check the script's configuration section.`);
    return null; // Return null to indicate failure
   }

  let label = null;
  try {
      label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
      Logger.log(`[WARN] Error checking for Gmail label "${labelName}": ${e}. This might happen if label access is restricted.`);
      // Don't attempt to create if checking failed for reasons other than 'not found'.
      // We'll proceed, and later checks for the label object will handle the failure.
  }

  if (!label) {
    if (DEBUG_MODE) Logger.log(`[DEBUG] Label "${labelName}" not found. Attempting to create...`);
    try {
        label = GmailApp.createLabel(labelName);
        Logger.log(`[INFO] Successfully created Gmail label: "${labelName}"`);
        // Optional: Alert user only on first creation?
        // SpreadsheetApp.getUi().alert(`Created missing Gmail label: "${labelName}". Please ensure your emails are labeled correctly.`);
    } catch (e) {
        Logger.log(`[ERROR] Failed to create Gmail label "${labelName}": ${e}\n${e.stack}. Please create this label manually in Gmail and ensure the script has permissions.`);
        SpreadsheetApp.getUi().alert(`Script Error: Failed to create required Gmail label "${labelName}". Please create it manually in Gmail and check script permissions.`);
        return null; // Return null on creation failure
    }
  }
  return label;
}


// --- Helper: Parse Company from Sender Domain ---
function parseCompanyFromDomain(sender) {
    const emailMatch = sender.match(/<([^>]+)>/);
    if (!emailMatch || !emailMatch[1]) return null; // No email address found in sender string

    const emailAddress = emailMatch[1];
    const domainParts = emailAddress.split('@');
    if (domainParts.length !== 2) return null; // Invalid email format

    let domain = domainParts[1].toLowerCase();

    // Check if the domain is in the explicitly ignored list (but allow specific subdomains if needed)
    if (IGNORED_DOMAINS.has(domain) && domain !== 'hi.wellfound.com') { // Example exception for wellfound notification subdomain
         if (DEBUG_MODE) Logger.log(`[DEBUG] Domain Parse: Domain "${domain}" is in IGNORED_DOMAINS.`);
         return null;
    }

    // Attempt to remove common prefixes like 'careers.', 'jobs.', etc.
    // Be slightly less aggressive: Match start of string, allow hyphen/dot after prefix
    domain = domain.replace(/^(?:careers|jobs|recruiting|apply|hr|talent|notification|notifications)[.-]/i, '');
    // Remove common TLDs and preceding punctuation/symbols (more robust)
    domain = domain.replace(/\.[a-z]{2,}(\.[a-z]{2})?$/i, ''); // Remove .com, .co.uk, .org etc.
    // Replace remaining non-alphanumeric characters (like hyphens) with spaces
    domain = domain.replace(/[^a-z0-9]/gi, ' ');
    // Capitalize words and join
    domain = domain.split(' ')
                   .filter(word => word.length > 0) // Remove empty strings resulting from multiple separators
                   .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                   .join(' ');

    const finalDomain = domain.trim();
    if (DEBUG_MODE) Logger.log(`[DEBUG] Domain Parse: Original: "${domainParts[1]}", Processed: "${finalDomain}"`);
    return finalDomain || null; // Return null if processing resulted in an empty string
}


// --- Helper: Parse Company from Sender Display Name ---
function parseCompanyFromSenderName(sender) {
    // Extract name part before the <email>
    const nameMatch = sender.match(/^"?(.*?)"?\s*</);
    let name = nameMatch ? nameMatch[1].trim() : sender.split('<')[0].trim();

    if (!name || name.includes('@')) return null; // Basic validation

    // Remove common platform indicators / noise words more carefully
    name = name.replace(/\|\s*(?:Greenhouse|Lever|Wellfound|Workday|Ashby|iCIMS|SmartRecruiters|Taleo|BambooHR)\b/gi, ''); // Case-insensitive, word boundary
    name = name.replace(/\s*(?:via Wellfound|via LinkedIn|via Indeed|from Greenhouse|from Lever)\b/gi, ''); // Case-insensitive, word boundary
    name = name.replace(/\s*(?:Careers|Recruiting|Recruitment|Hiring Team|Hiring|Talent Acquisition|Talent|HR|Team|Notifications?|Jobs?|Updates?|Apply)\b/gi, ''); // Case-insensitive, word boundary

    // Remove trailing punctuation or common noise words often left after replacements
    name = name.replace(/[|,-_.\s]+$/, '').trim(); // Remove trailing symbols/whitespace

    // Basic sanity check: Avoid very short names or common non-company names
    if (name.length > 2 && !/^(?:noreply|no-reply|jobs|careers|support|info|admin|hr|talent|recruiting)$/i.test(name)) {
       if (DEBUG_MODE) Logger.log(`[DEBUG] Sender Name Parse: Original: "${sender}", Processed: "${name}"`);
       return name;
    }

    if (DEBUG_MODE && name) Logger.log(`[DEBUG] Sender Name Parse: Original: "${sender}", Discarded (too short or generic): "${name}"`);
    return null;
}


// --- Helper: Clean Text Extracted from HTML --- // (Currently unused but potentially useful)
function cleanHtmlText(htmlSnippet) {
  if (!htmlSnippet) return "";
  // Basic HTML entity decoding (add more if needed)
  let cleaned = htmlSnippet.replace(/ /g, ' ')
                           .replace(/&/g, '&')
                           .replace(/</g, '<')
                           .replace(/>/g, '>')
                           .replace(/"/g, '"')
                           .replace(/'/g, "'");
  // Strip HTML tags
  cleaned = cleaned.replace(/<style([\s\S]*?)<\/style>/gi, '');   // Remove style blocks
  cleaned = cleaned.replace(/<script([\s\S]*?)<\/script>/gi, '');  // Remove script blocks
  cleaned = cleaned.replace(/<[^>]*>/g, ' '); // Remove remaining tags
  // Normalize whitespace
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  return cleaned;
}


// --- Parsing Orchestrator ---
// Accepts pre-fetched plainBody to avoid redundant API calls.
function extractCompanyAndTitle(message, platform, emailSubject, plainBody) { // Added plainBody parameter
    // --- INITIALIZE ---
    let company = MANUAL_REVIEW_NEEDED;
    let title = MANUAL_REVIEW_NEEDED;
    const sender = message.getFrom() || ""; // Ensure sender is never null
    // Use the passed-in plainBody parameter directly (fetched once in the main loop)

    if (DEBUG_MODE) Logger.log(`[DEBUG] PARSE_DETAIL: Starting Company/Title Extraction for Subject: "${emailSubject}", Sender: "${sender}"`);

    // --- Pre-parse Sender Domain & Name (for fallback) ---
    let tempCompanyFromDomain = null;
    let tempCompanyFromName = null;
    try {
        tempCompanyFromDomain = parseCompanyFromDomain(sender);
        if (DEBUG_MODE && tempCompanyFromDomain) Logger.log(`[DEBUG]  -> Pre-parsed Domain Company: "${tempCompanyFromDomain}"`);
    } catch (e) { Logger.log(`[WARN] Error during initial company domain parsing: ${e}`); }
    try {
        tempCompanyFromName = parseCompanyFromSenderName(sender);
         if (DEBUG_MODE && tempCompanyFromName) Logger.log(`[DEBUG]  -> Pre-parsed Sender Name Company: "${tempCompanyFromName}"`);
    } catch (e) { Logger.log(`[WARN] Error during initial company sender name parsing: ${e}`); }


    // --- Tiered Parsing Logic ---

    // 1. Platform Specific Logic (Example: Wellfound - refined)
    if (platform === "Wellfound") {
        if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Running Wellfound specific parsing logic...");
        try {
            if (plainBody) {
                // --- Wellfound Subject/Body Company Parsing ---
                // Prioritize subject matches
                let wfCoSub = emailSubject.match(/update from (.*?)(?: \| |$)/i) // More specific separator
                             || emailSubject.match(/application to (.*?)(?: successfully|$)/i);
                if (wfCoSub && wfCoSub[1]) {
                    company = wfCoSub[1].trim();
                    if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Wellfound Subject Parse Success (Company): "${company}"`);
                } else {
                     // Fallback to body header if subject fails
                     const bodyLinesWF = plainBody.substring(0, 600).split('\n'); // Look near top
                     // Regex for potential company name format (starts uppercase, allows spaces/common symbols, avoids sentences)
                     const companyHeaderRegexWF = /^[A-Z][A-Za-z\s.&'-]+$/;
                     // Find the first line that looks like a company name header
                     const companyHeaderIndexWF = bodyLinesWF.findIndex(l => {
                         const trimmedLine = l.trim();
                         return companyHeaderRegexWF.test(trimmedLine) && trimmedLine.length > 1 && trimmedLine.split(' ').length < 6; // Keep it shortish
                     });

                     if (companyHeaderIndexWF !== -1) {
                         const potentialCompany = bodyLinesWF[companyHeaderIndexWF].trim();
                         // Avoid grabbing lines that are likely job titles or locations instead
                         if (!/(?:engineer|developer|manager|analyst|designer|remote|hybrid|onsite)/i.test(potentialCompany)) {
                            company = potentialCompany;
                            if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Wellfound Body Header Parse Success (Company): "${company}"`);
                         } else {
                             if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Wellfound Body Header candidate "${potentialCompany}" skipped (looks like title/location).`);
                         }
                     }
                }

                // --- Wellfound Title Parsing (if company found and title needed) ---
                if (company !== MANUAL_REVIEW_NEEDED && title === MANUAL_REVIEW_NEEDED) {
                    const bodyLinesWF = plainBody.substring(0, 800).split('\n'); // Search slightly wider
                    const companyHeaderIndexWF = bodyLinesWF.findIndex(l => l.trim() === company); // Find the company line again
                    if (companyHeaderIndexWF !== -1) {
                         // Look for title on the *next few lines* after the company header
                         for (let j = companyHeaderIndexWF + 1; j < Math.min(companyHeaderIndexWF + 4, bodyLinesWF.length); j++) {
                             const potentialTitleLine = bodyLinesWF[j].trim();
                             // Check if it looks like a title: contains letters, not just symbols/numbers, not obviously location/salary
                             if (potentialTitleLine && /[a-zA-Z]{3,}/.test(potentialTitleLine) &&
                                 !/[$]|\d{5,}|(remote|hybrid|onsite|location:|new york|san francisco|london)/i.test(potentialTitleLine) &&
                                  potentialTitleLine.split(' ').length < 8 // Avoid long sentences
                                 ) {
                                 title = potentialTitleLine;
                                 if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Wellfound Body Parse Success (Title near header): "${title}"`);
                                 break; // Found title, stop looking
                             }
                         }
                     }
                 }

                 // --- Specific Wellfound Plain Text '* Title' Pattern Logic (if title still needed) ---
                 if (title === MANUAL_REVIEW_NEEDED && sender.toLowerCase().includes("team@hi.wellfound.com")) {
                      if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Wellfound title still needed, trying PLAIN TEXT '* pattern' scan...");
                     try {
                        const lines = plainBody.split('\n');
                        // Find the marker phrase reliably, ignoring potential encoding quirks like =XX
                        const markerPhrase = "if there's a match, we will make an email introduction.";
                        const markerIndex = lines.findIndex(line => line.replace(/=\d{2}/g, '').trim().toLowerCase().includes(markerPhrase.toLowerCase()));

                        if (markerIndex !== -1) {
                            if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found marker phrase at line index ${markerIndex}. Scanning subsequent lines...`);
                            for (let i = markerIndex + 1; i < lines.length; i++) {
                                const currentLine = lines[i].replace(/=\d{2}/g, '').trim(); // Clean line
                                if (currentLine.startsWith('* ')) {
                                     const potentialTitle = currentLine.substring(2).trim();
                                    // Basic validation for the extracted title
                                    if (potentialTitle && potentialTitle.length > 2 && /[a-zA-Z]/.test(potentialTitle)) {
                                        title = potentialTitle;
                                        Logger.log(`[INFO] PARSE_DETAIL: Extracted Wellfound title from plain text '* ' pattern: "${title}"`);
                                        break; // Found it
                                    } else {
                                         if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found line starting with '* ' but content invalid: "${currentLine}"`);
                                    }
                                } else if (currentLine.length > 5 && !currentLine.match(/^[\s*\-=–—<>]+$/)) {
                                    // Stop if we hit a significant line of text that isn't the title pattern
                                    if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found significant non-empty/non-marker line before finding '* ' pattern. Stopping '*' scan. Line: "${currentLine}"`);
                                    break;
                                }
                             }
                            if (title === MANUAL_REVIEW_NEEDED && DEBUG_MODE) { Logger.log("[DEBUG]  -> Scanned lines after marker phrase, did not find valid '* ' title pattern."); }
                         } else if (DEBUG_MODE) { Logger.log("[DEBUG]  -> Marker phrase for '*' scan not found in plain body."); }
                     } catch (e) { Logger.log(`[ERROR] During Wellfound plain text '* ' title parsing: ${e}`); }
                 }

             } else { Logger.log("[WARN] PARSE_DETAIL: Plain body was null/unavailable for Wellfound parsing."); }
        } catch (e) { Logger.log(`[ERROR] During Wellfound specific platform parsing block: ${e}`); }
    } // --- End Wellfound Specific ---


    // --- General Parsing (applied if platform-specific logic didn't find everything) ---

    // 2. Simple Subject Regex (If Company still needed) - More specific patterns
    if (company === MANUAL_REVIEW_NEEDED) {
        if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Trying Simple Subject Regex for Company...");
        let simpleSubjectMatch = emailSubject.match(/^(?:Re:|Fwd:)?\s*(?:Application|Update|Interest).*? for (.+?) at (.+)/i) // "for TITLE at COMPANY"
                                || emailSubject.match(/^(?:Re:|Fwd:)?\s*Update from ([^-:|]+)/i) // "Update from COMPANY"
                                || emailSubject.match(/^(?:Re:|Fwd:)?\s*([A-Z][\w\s.&'-]+?)\s*[|–—:-]\s*Application/i); // "COMPANY | Application..."
         if (simpleSubjectMatch) {
             // Determine which capture group holds the company based on the matched regex
             let potentialCompany = (simpleSubjectMatch.length > 2 && simpleSubjectMatch[2]) ? simpleSubjectMatch[2] : simpleSubjectMatch[1];
             potentialCompany = potentialCompany.trim();
             // Basic sanity check - avoid grabbing something that looks like a job title
             if (!/(?:engineer|developer|manager|analyst|designer|lead|specialist|intern)\b/i.test(potentialCompany) || potentialCompany.split(' ').length > 4) {
                 company = potentialCompany;
                 if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found Company using Simple Subject Regex: "${company}"`);
             } else {
                  if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Simple Subject Regex matched "${potentialCompany}" but skipped (looks like title).`);
             }
         }
    }

    // 3. Complex Subject Regex (If Company or Title still needed) - Prioritize structure
     if (company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED) {
         if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Trying Complex Subject Regex...");
         // Patterns ordered from more specific/reliable to less specific
         // Added non-capturing groups (?:...) for prefixes like Re:, Fwd:
         // Made separators more explicit [-–—:|]
         // Added word boundaries \b where appropriate
         const complexPatterns = [
             // Structure: Invite/Interview... for TITLE at COMPANY
             { r: /^(?:Re:|Fwd:)?\s*Invite.*?interview.*? for the? (.+?) role at ([^-:|]+)/i, titleIdx: 1, companyIdx: 2 },
             { r: /^(?:Re:|Fwd:)?\s*Invite.*?interview.*? for (.+?) at ([^-:|]+)/i, titleIdx: 1, companyIdx: 2 },
             // Structure: Invite/Interview... : TITLE at COMPANY
             { r: /^(?:Re:|Fwd:)?\s*Interview.*? [-–—:|]\s*(.+?) at (.+)/i, titleIdx: 1, companyIdx: 2 },
             // Structure: Application for TITLE at COMPANY
             { r: /^(?:Re:|Fwd:)?\s*Application for (?:the )?(.+?)(?:\s*role)?\s+at\s+([^-:|]+)/i, titleIdx: 1, companyIdx: 2 },
             // Structure: Your Application for TITLE - COMPANY or COMPANY - TITLE
             { r: /^(?:Re:|Fwd:)?\s*(?:Your )?Application(?: for)?\s+(.+?)\s*[-–—|]\s*([\w\s.&'-]+)/i, titleIdx: 1, companyIdx: 2 }, // Title - Company more likely
             { r: /^(?:Re:|Fwd:)?\s*([\w\s.&'-]+?)\s*[-–—|]\s*(?:Your )?Application(?: for)?\s+(.+)/i, titleIdx: 2, companyIdx: 1 }, // Company - Title
             // Structure: Interest in the TITLE role at COMPANY
             { r: /^(?:Re:|Fwd:)?\s*Interest in the\s+(.+?)\s+role(?:\s+at\s+([^-:|]+))?/i, titleIdx: 1, companyIdx: 2 },
             // Structure: Update on your TITLE application at COMPANY
             { r: /^(?:Re:|Fwd:)?\s*Update on your\s+(.+?)\s+app(?:lication)?(?:\s+at\s+(.+))?/i, titleIdx: 1, companyIdx: 2 },
             // Structure: Next Steps: TITLE Application or similar headers
             { r: /^(?:Re:|Fwd:)?\s*(?:Next Steps|Regarding Your Application)(?: for)?\s*[-–—:|]\s*(.+?)(?:\s+at\s+([^-:|]+))?/i, titleIdx: 1, companyIdx: 2 },
             // Structure: COMPANY | TITLE (common in automated systems)
             { r: /^(?:Re:|Fwd:)?\s*([A-Z][\w\s.&'-]+?)\s*[|]\s*(.+)/i, titleIdx: 2, companyIdx: 1 } // Company | Title
         ];
         try {
            for (const patternInfo of complexPatterns) {
                let match = emailSubject.match(patternInfo.r);
                if (match) {
                    if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Complex Subject Regex matched: ${patternInfo.r}`);
                    // Extract Title if needed and found
                    if (title === MANUAL_REVIEW_NEEDED && patternInfo.titleIdx > 0 && match[patternInfo.titleIdx]) {
                        const potentialTitle = match[patternInfo.titleIdx].trim();
                        // Avoid extracting company name as title if structure is ambiguous (e.g., "Company | Application Update")
                        if (!/application|update|assessment|interview|offer/i.test(potentialTitle) || potentialTitle.split(' ').length > 1) {
                             title = potentialTitle;
                             if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Extracted Title: "${title}"`);
                        } else {
                            if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Potential Title "${potentialTitle}" skipped (looks like status/generic).`);
                        }
                    }
                    // Extract Company if needed and found
                    if (company === MANUAL_REVIEW_NEEDED && patternInfo.companyIdx > 0 && match[patternInfo.companyIdx]) {
                        const potentialCompany = match[patternInfo.companyIdx].trim();
                        // Avoid extracting title or generic terms as company
                         if (potentialCompany.length > 1 && !/application|update|assessment|interview|offer|role|position/i.test(potentialCompany)) {
                            company = potentialCompany;
                            if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Extracted Company: "${company}"`);
                         } else {
                             if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Potential Company "${potentialCompany}" skipped (looks like title/generic).`);
                         }
                    }
                    // Stop if both are found
                    if (company !== MANUAL_REVIEW_NEEDED && title !== MANUAL_REVIEW_NEEDED) {
                        if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found both Company and Title via Complex Subject Regex. Stopping pattern search.`);
                        break;
                    }
                }
            }
        } catch (e) { Logger.log(`[ERROR] During Complex Subject Regex processing: ${e}`); }
    }

    // 4. Body Scan Fallback (If Company or Title still needed) - Be cautious
     if (plainBody && (company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED)) {
        if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Attempting Body Scan Fallback (using first ~750 chars)...");
        try {
            const bodyStartLower = plainBody.substring(0, 750).toLowerCase(); // Use lowercase for matching

            // Company from body (look for patterns like "applying to X at Y", "application with Y")
            if (company === MANUAL_REVIEW_NEEDED) {
                 // Regex: Look for phrases indicating company context, capture the likely company name (starts uppercase, reasonable length)
                 let bodyCompanyMatch = plainBody.substring(0, 750).match(/(?:applying to|application with|interest in|interview with)\s+([A-Z][\w\s.&'-]{2,})/); // Simpler, looks for capitalized word/phrase after keywords
                 if (bodyCompanyMatch && bodyCompanyMatch[1]) {
                     const potentialCompany = bodyCompanyMatch[1].trim();
                      if (!/(?:engineer|developer|manager|analyst|designer|lead|specialist|intern|role|position)\b/i.test(potentialCompany)) {
                         company = potentialCompany;
                         if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found Company via Body Scan: "${company}"`);
                      } else {
                          if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Body Scan company candidate "${potentialCompany}" skipped (looks like title).`);
                      }
                 }
            }

            // Title from body (look for "application for the X position/role", "role of X")
            if (title === MANUAL_REVIEW_NEEDED) {
                // Regex: Capture text after common title phrases, stop at punctuation or words like 'at', 'with', 'in'
                 let bodyTitleMatch = bodyStartLower.match(/(?:position of|role of|application for|regarding the|interest in the)\s+([\w\s\-()&',./]+?)\s*(?:position|role|opening|opportunity|at|with|in|[\n.,(])/i);
                 if (bodyTitleMatch && bodyTitleMatch[1]) {
                     let potentialTitle = bodyTitleMatch[1].trim();
                     // Clean common contaminants like (Remote) or salary info often following title in body
                     potentialTitle = potentialTitle.replace(/\s*\(.*?(?:remote|hybrid|onsite|contract|salary|usd|eur|gbp)[\s\S]*?\)/gi, '');
                     potentialTitle = potentialTitle.replace(/[-–—]\s*(?:remote|hybrid|onsite|contract)\s*$/gi, ''); // Trailing status
                     potentialTitle = potentialTitle.replace(/\b(?:at|with|in|for)\b.*$/i, ''); // Remove trailing prepositional phrases
                     potentialTitle = potentialTitle.trim();

                     if (potentialTitle.length > 2 && potentialTitle.split(' ').length < 8) { // Basic sanity check
                         title = potentialTitle;
                         // Capitalize title properly
                         title = title.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
                         if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found Title via Body Scan: "${title}"`);
                     } else {
                          if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Body Scan title candidate "${bodyTitleMatch[1]}" skipped (failed validation). Final cleaned: "${potentialTitle}"`);
                     }
                 }
            }
        } catch (e) { Logger.log(`[ERROR] During Plain Body Scan fallback: ${e}`); }
    } else if (!plainBody && (company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED)) {
         Logger.log("[WARN] PARSE_DETAIL: Plain body was null/unavailable for Body Scan fallback.");
    }

    // 5. Use Pre-Parsed Sender Info as Final Fallback for Company
     if (company === MANUAL_REVIEW_NEEDED) {
         if (DEBUG_MODE) Logger.log("[DEBUG] PARSE_DETAIL: Company still not found, attempting final fallback using sender info...");
        if (tempCompanyFromName) { // Prefer sender name if available and reasonable
            company = tempCompanyFromName;
            if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Using Fallback Company from Sender Name: "${company}"`);
        } else if (tempCompanyFromDomain) { // Use domain parse as second fallback
            company = tempCompanyFromDomain;
            if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Using Fallback Company from Sender Domain: "${company}"`);
        } else {
            if (DEBUG_MODE) Logger.log("[DEBUG]  -> Fallback failed: No valid pre-parsed company name or domain found.");
        }
    }

    // 6. Final Cleaning and Validation
     try {
        if (company !== MANUAL_REVIEW_NEEDED) {
            // Remove common legal suffixes, trailing junk, ensure not empty
            company = company.split(/[\n\r#(]/)[0] // Take first line if multiline found
                             .replace(/\s+(?:inc|llc|ltd|corp|gmbh|incorporated|limited|corporation)[.,]?$/i, '') // Remove suffixes
                             .replace(/[,"']?$/, '') // Remove trailing quote/comma
                             .replace(/&/g, '&') // Decode ampersand if missed
                             .trim();
            if (!company || company.length < 2) { // If cleaning results in empty/too short, revert
                if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Final Company cleaning resulted in invalid name ("${company}"). Reverting to Manual Review.`);
                company = MANUAL_REVIEW_NEEDED;
            }
        }
     } catch (e) { Logger.log(`[WARN] Error during final Company name cleaning: ${e}`); company = MANUAL_REVIEW_NEEDED; }

     try {
        if (title !== MANUAL_REVIEW_NEEDED) {
            // Remove common junk like job codes, location specifiers, ensure not empty
             title = title.replace(/^JR\d+\s*[-–—]?\s*/i, '') // Job Req IDs at start
                       .replace(/\(.*?(?:remote|hybrid|onsite|contract|location|based|id:\s*\d+).*?\)/gi, '') // Parenthetical info
                       .replace(/[-–—]\s*(?:remote|hybrid|onsite|contract)\s*$/gi, '') // Trailing status
                       .replace(/[\n\r]/g, ' ') // Replace newlines with space
                       .replace(/\s+/g, ' ') // Collapse multiple spaces
                       .replace(/[,\.]*$/, '') // Remove trailing commas/periods
                       .replace(/^["'-–—]+|["'-–—]+$/g, '') // Remove leading/trailing quotes/hyphens
                       .trim();
             if (!title || title.length < 3 || /^\d+$/.test(title)) { // If cleaning results in empty/too short/only numbers, revert
                if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Final Title cleaning resulted in invalid name ("${title}"). Reverting to Manual Review.`);
                title = MANUAL_REVIEW_NEEDED;
             } else {
                // Attempt proper title case (simple version)
                title = title.split(' ').map(word => {
                    if (word.length > 0) {
                        // Handle common initialisms like API, UI, UX, QA, IT, HR etc. - keep uppercase
                        if (/^(?:api|ui|ux|qa|it|hr|vp|sr|jr|ceo|cto|cfo|ios|id)$/i.test(word)) {
                           return word.toUpperCase();
                        }
                        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
                    }
                    return '';
                }).join(' ');
             }
        }
     } catch (e) { Logger.log(`[WARN] Error during final Title cleaning: ${e}`); title = MANUAL_REVIEW_NEEDED; }

    // Log final outcome
    Logger.log(`[INFO] PARSE_DETAIL: --- Final Extraction Result --- Company: "${company}", Title: "${title}"`);
    return { company: company, title: title };
} // --- End extractCompanyAndTitle ---


// --- Helper: Parse Email Body for Status Keywords ---
function parseBodyForStatus(plainBody) { // Accepts plainBody
    if (!plainBody || plainBody.length < 10) { // Basic check if body is usable
        if (DEBUG_MODE) Logger.log("[DEBUG] STATUS_PARSE: Skipping status check - plain body too short or missing.");
        return null;
    }
    // Prepare body text: lowercase, remove punctuation, normalize whitespace
    // Consider removing common email boilerplate first? (e.g., signatures) - potentially complex
    let bodyLower = plainBody.toLowerCase();
    // More comprehensive punctuation/symbol removal
    bodyLower = bodyLower.replace(/[.,!?;:()\[\]{}'"“”‘’\-–—@#$%^&*+=<>]/g, ' ');
    bodyLower = bodyLower.replace(/\s+/g, ' ').trim(); // Normalize whitespace

    // Check keyword arrays (defined at top) - order matters (e.g., check acceptance before interview)
    if (ACCEPTANCE_KEYWORDS.some(keyword => bodyLower.includes(` ${keyword} `) || bodyLower.startsWith(`${keyword} `) || bodyLower.endsWith(` ${keyword}`))) {
        if (DEBUG_MODE) Logger.log(`[DEBUG] STATUS_PARSE: Found ACCEPTANCE keyword match.`);
        return ACCEPTED_STATUS;
    }
    if (INTERVIEW_KEYWORDS.some(keyword => bodyLower.includes(` ${keyword} `) || bodyLower.startsWith(`${keyword} `) || bodyLower.endsWith(` ${keyword}`))) {
        if (DEBUG_MODE) Logger.log(`[DEBUG] STATUS_PARSE: Found INTERVIEW keyword match.`);
        return INTERVIEW_STATUS;
    }
    if (ASSESSMENT_KEYWORDS.some(keyword => bodyLower.includes(` ${keyword} `) || bodyLower.startsWith(`${keyword} `) || bodyLower.endsWith(` ${keyword}`))) {
        if (DEBUG_MODE) Logger.log(`[DEBUG] STATUS_PARSE: Found ASSESSMENT keyword match.`);
        return ASSESSMENT_STATUS;
    }
    // Check rejection last as its keywords can sometimes appear in other contexts
    if (REJECTION_KEYWORDS.some(keyword => bodyLower.includes(` ${keyword} `) || bodyLower.startsWith(`${keyword} `) || bodyLower.endsWith(` ${keyword}`))) {
        if (DEBUG_MODE) Logger.log(`[DEBUG] STATUS_PARSE: Found REJECTION keyword match.`);
        return REJECTED_STATUS;
    }

    if (DEBUG_MODE) Logger.log("[DEBUG] STATUS_PARSE: No specific status keywords found in body.");
    return null; // No definitive status found
} // --- End parseBodyForStatus ---


// --- Helper: Apply Final Gmail Labels ---
function applyFinalLabels(processedThreadOutcomes, processingLabel, processedLabel, manualReviewLabel) {
    const threadIdsToLabel = Object.keys(processedThreadOutcomes);
    if (threadIdsToLabel.length === 0) {
        Logger.log("[INFO] LABEL_MGMT: No thread outcomes require label updates.");
        return;
    }
    Logger.log(`[INFO] LABEL_MGMT: Starting final label application based on ${threadIdsToLabel.length} thread outcomes...`);
    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;

    // Validate label objects before starting the loop
    if (!processingLabel || typeof processingLabel.getName !== 'function') { Logger.log("[ERROR] LABEL_MGMT: Invalid 'processingLabel' object passed. Aborting labeling."); return; }
    if (!processedLabel || typeof processedLabel.getName !== 'function') { Logger.log("[ERROR] LABEL_MGMT: Invalid 'processedLabel' object passed. Aborting labeling."); return; }
    if (!manualReviewLabel || typeof manualReviewLabel.getName !== 'function') { Logger.log("[ERROR] LABEL_MGMT: Invalid 'manualReviewLabel' object passed. Aborting labeling."); return; }

    const processingLabelName = processingLabel.getName();
    const processedLabelName = processedLabel.getName();
    const manualReviewLabelName = manualReviewLabel.getName();

    for (const threadId of threadIdsToLabel) {
        const outcome = processedThreadOutcomes[threadId]; // Should be 'done' or 'manual'
        const targetLabelObject = (outcome === 'manual') ? manualReviewLabel : processedLabel;
        const targetLabelName = targetLabelObject.getName(); // Get name directly from valid object

        try {
            const thread = GmailApp.getThreadById(threadId);
            if (!thread) {
                Logger.log(`[WARN] LABEL_MGMT: Thread ${threadId} not found (might have been deleted or archived?). Skipping label update.`);
                skippedCount++;
                continue;
            }

            const currentLabels = thread.getLabels().map(l => l.getName());
            let labelChanged = false;

            // 1. Remove Processing Label (if it exists)
            if (currentLabels.includes(processingLabelName)) {
                try {
                    thread.removeLabel(processingLabel);
                    if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Removed label "${processingLabelName}" from Thread ${threadId}`);
                    labelChanged = true;
                    Utilities.sleep(150); // Pause after modification
                 } catch (e) {
                     Logger.log(`[WARN] LABEL_MGMT: Failed to remove label "${processingLabelName}" from Thread ${threadId}: ${e}`);
                     // Continue trying to add the target label even if removal fails
                 }
            } else {
                 if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Thread ${threadId} did not have label "${processingLabelName}". No removal needed.`);
            }

            // 2. Add Target Label (if it doesn't exist)
            if (!currentLabels.includes(targetLabelName)) {
                 try {
                    thread.addLabel(targetLabelObject);
                    Logger.log(`[INFO] LABEL_MGMT: Added label "${targetLabelName}" to Thread ${threadId}`);
                    labelChanged = true;
                    Utilities.sleep(150); // Pause after modification
                 } catch(e) {
                     Logger.log(`[ERROR] LABEL_MGMT: Failed to add target label "${targetLabelName}" to Thread ${threadId}: ${e}`);
                     errorCount++; // Count as an error if adding fails
                     continue; // Skip to next thread on critical add error
                 }
            } else {
                if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Thread ${threadId} already has target label "${targetLabelName}". No add needed.`);
                 // If processing label was removed but target already exists, still counts as success for this thread.
            }

            if (labelChanged) {
                 successCount++;
            } else {
                 skippedCount++; // Increment if no changes were made (labels already correct)
            }

        } catch (e) {
            Logger.log(`[ERROR] LABEL_MGMT: Unhandled error processing labels for Thread ${threadId}: ${e}`);
            errorCount++;
        }
        Utilities.sleep(200); // Slightly longer pause between threads to be safe
    }
    Logger.log(`[INFO] LABEL_MGMT: Finished applying labels. Label changes applied: ${successCount}. Skipped (no changes needed/thread missing): ${skippedCount}. Errors: ${errorCount}.`);
} // --- End applyFinalLabels ---


/**
 * Main Processing Function: processJobApplicationEmails (v8.1 - Template)
 * This is the primary function called by the trigger or manual execution.
 */
function processJobApplicationEmails() {
    const SCRIPT_START_TIME = new Date();
    Logger.log(`\n\n==== SCRIPT EXECUTION STARTING (${SCRIPT_START_TIME.toLocaleString()}) ====\n(Version: 8.1 - Template)`);

    // --- Initial Setup & Validation ---
    let ss;
    try {
        ss = SpreadsheetApp.getActiveSpreadsheet(); // This script MUST be bound to a sheet
    } catch (e) {
        Logger.log(`[FATAL ERROR] Could not get active spreadsheet. Is the script correctly bound to a Google Sheet? Error: ${e}`);
        // Cannot use SpreadsheetApp.getUi() if ss is not available. Log is the best we can do.
        return;
    }

    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        Logger.log(`[FATAL ERROR] Sheet "${SHEET_NAME}" not found in the active spreadsheet. Check the SHEET_NAME configuration in the script. Aborting.`);
        try { SpreadsheetApp.getUi().alert(`Script Error: Sheet "${SHEET_NAME}" not found. Please check the SHEET_NAME configuration in the script.`); } catch (uiError) {/* ignore if UI unavailable */}
        return;
    }
     if (DEBUG_MODE) Logger.log(`[DEBUG] Target sheet "${SHEET_NAME}" found.`);

    // Get or create labels, checking for validity
    const procLbl = getOrCreateLabel(GMAIL_LABEL_TO_PROCESS);
    const doneLbl = getOrCreateLabel(GMAIL_LABEL_APPLIED_AFTER_PROCESSING);
    const manualLbl = getOrCreateLabel(GMAIL_LABEL_MANUAL_REVIEW);

    if (!procLbl || !doneLbl || !manualLbl) {
        Logger.log(`[FATAL ERROR] One or more required Gmail labels ("${GMAIL_LABEL_TO_PROCESS}", "${GMAIL_LABEL_APPLIED_AFTER_PROCESSING}", "${GMAIL_LABEL_MANUAL_REVIEW}") could not be found or created. Please ensure the labels exist in Gmail or check script permissions/configuration. Aborting.`);
        // Alert is handled within getOrCreateLabel on failure now
        return;
    }
     if (DEBUG_MODE) Logger.log(`[DEBUG] Required Gmail Labels OK.`);

    // Validate column configuration numbers
    const requiredCols = [PROCESSED_TIMESTAMP_COL, EMAIL_DATE_COL, PLATFORM_COL, COMPANY_COL, JOB_TITLE_COL, STATUS_COL, LAST_UPDATE_DATE_COL, EMAIL_SUBJECT_COL, EMAIL_LINK_COL, EMAIL_ID_COL];
    if (requiredCols.some(col => col <= 0) || TOTAL_COLUMNS_IN_SHEET < Math.max(...requiredCols)) {
         Logger.log(`[FATAL ERROR] Invalid column configuration. Check column index numbers (must be > 0) and TOTAL_COLUMNS_IN_SHEET (must be >= highest index) in the script. Aborting.`);
         try { SpreadsheetApp.getUi().alert(`Script Configuration Error: Invalid column numbers defined. Please check the column index constants at the top of the script.`); } catch (uiError) {/* ignore */}
         return;
    }


    // --- 1. Pre-load Existing Data ---
    const lastR = sheet.getLastRow();
    const existD = {}; // Cache for Company -> [{row, emailId, company, title}]
    const existIds = new Set(); // Cache for processed Email IDs
    if (lastR >= 2) { // Only preload if there's data beyond the header
        Logger.log(`[INFO] PRELOAD: Loading existing data from Sheet Row 2 to ${lastR}...`);
        try {
            // Determine the actual range needed based on configured columns
            const colsToPreload = [COMPANY_COL, JOB_TITLE_COL, EMAIL_ID_COL]; // Columns needed for matching/deduping
            const startCol = Math.min(...colsToPreload);
            const endCol = Math.max(...colsToPreload);
            const numColsToRead = endCol - startCol + 1;

            if (numColsToRead < 1 || startCol < 1 || endCol > sheet.getMaxColumns()) {
                 throw new Error(`Invalid column calculation for preloading. MinCol: ${startCol}, MaxCol: ${endCol}, SheetMaxCol: ${sheet.getMaxColumns()}`);
            }

            const preloadRange = sheet.getRange(2, startCol, lastR - 1, numColsToRead);
            const preloadValues = preloadRange.getValues();

            // Calculate the index within the *read array* for each needed column
            const companyIndexInPreload = COMPANY_COL - startCol;
            const titleIndexInPreload = JOB_TITLE_COL - startCol;
            const emailIdIndexInPreload = EMAIL_ID_COL - startCol;

            for (let i = 0; i < preloadValues.length; i++) {
                const rowNum = i + 2; // Actual sheet row number
                const rowData = preloadValues[i];
                const emailId = rowData[emailIdIndexInPreload]?.toString().trim() || "";
                const originalCompany = rowData[companyIndexInPreload]?.toString().trim() || "";
                const originalTitle = rowData[titleIndexInPreload]?.toString().trim() || "";

                // Add email ID to the set for quick deduplication check
                if (emailId) {
                    existIds.add(emailId);
                }

                // Add company info to the lookup dictionary if company name is valid
                const companyLower = originalCompany.toLowerCase();
                if (companyLower && companyLower !== 'n/a' && companyLower !== MANUAL_REVIEW_NEEDED.toLowerCase()) {
                    if (!existD[companyLower]) {
                        existD[companyLower] = [];
                    }
                    existD[companyLower].push({
                        row: rowNum,
                        emailId: emailId, // Store the ID associated with this row
                        company: originalCompany, // Store original case company
                        title: originalTitle   // Store original case title
                    });
                }
            }
            // Sort company matches by row number descending so the latest is first
            for (const compKey in existD) {
                 existD[compKey].sort((a, b) => b.row - a.row);
            }

            Logger.log(`[INFO] PRELOAD: Complete. Cached lookups for ${Object.keys(existD).length} companies. Known processed email IDs: ${existIds.size}.`);
        } catch (e) {
            Logger.log(`[FATAL ERROR] During Data Preload: ${e}\nStack: ${e.stack}\nPlease ensure column numbers in the script configuration are correct and the sheet is accessible. Aborting.`);
            try { SpreadsheetApp.getUi().alert(`Script Error: Failed to preload data from the sheet. Check Logs. Error: ${e.message}`); } catch (uiError) {/* ignore */}
            return;
        }
    } else {
        Logger.log(`[INFO] PRELOAD: Sheet has no data beyond header row. Skipping preload.`);
    }


    // --- 2. Gather & Sort New Messages ---
    let threads;
    try {
        if (DEBUG_MODE) Logger.log(`[DEBUG] GATHER: Fetching email threads with label "${procLbl.getName()}"...`);
        threads = procLbl.getThreads();
        if (DEBUG_MODE) Logger.log(`[DEBUG] GATHER: Found ${threads.length} threads with the label.`);
    } catch (e) {
         Logger.log(`[FATAL ERROR] Failed to get threads for label "${procLbl.getName()}": ${e}. Check Gmail permissions or label validity. Aborting.`);
         try { SpreadsheetApp.getUi().alert(`Script Error: Failed to access Gmail threads for label "${procLbl.getName()}". Check permissions/label. Error: ${e.message}`); } catch (uiError) {/* ignore */}
         return;
    }

    const msgToSort = [];
    let skippedCount = 0;
    let fetchErrorCount = 0;
    if (DEBUG_MODE) Logger.log(`[DEBUG] GATHER: Iterating threads, fetching messages, and checking against ${existIds.size} known processed IDs...`);

    for (const thread of threads) {
        const threadId = thread.getId();
        try {
            const messages = thread.getMessages();
            for (const msg of messages) {
                const msgId = msg.getId();
                if (!existIds.has(msgId)) {
                    // Only add message if its ID hasn't been processed before
                    msgToSort.push({
                        message: msg,
                        date: msg.getDate(), // Get date for sorting
                        threadId: threadId
                    });
                     if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found NEW message: ID ${msgId} (Thread ${threadId})`);
                } else {
                     if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Skipping OLD message: ID ${msgId} (Thread ${threadId})`);
                    skippedCount++;
                }
            }
        } catch (e) {
            Logger.log(`[ERROR] GATHER: Failed to fetch messages for Thread ${threadId}: ${e}`);
            fetchErrorCount++;
            // Decide if thread should be marked for manual review immediately? Potentially noisy.
        }
        Utilities.sleep(50); // Small pause between threads during fetch
    } // End gather loop

    Logger.log(`[INFO] GATHER: Identification complete. Found ${msgToSort.length} NEW messages to process. Skipped ${skippedCount} already processed messages. Encountered fetch errors on ${fetchErrorCount} threads.`);

    // --- Early Exit if No New Messages ---
    if (msgToSort.length === 0) {
        Logger.log("[INFO] PROCESS: No new messages found with the processing label requiring action.");
        let staleLabelOutcomes = {};
        // Check if any threads *still* have the processing label but contain no new messages
        // This helps clean up labels if all messages in a thread were already processed individually before.
        threads.forEach(th => {
            try {
                const thId = th.getId();
                 if (th.getLabels().some(l => l.getName() === procLbl.getName())) {
                     // Assume 'done' for cleanup unless a fetch error occurred for this thread earlier? Too complex maybe.
                     staleLabelOutcomes[thId] = 'done';
                     if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Thread ${thId} still has "${procLbl.getName()}" label but no new messages found. Marking for cleanup.`);
                 }
             } catch (e) {Logger.log(`[WARN] Error checking labels on thread ${th.getId()} during cleanup check: ${e}`);}
        });
        if (Object.keys(staleLabelOutcomes).length > 0) {
             Logger.log(`[INFO] FINISH: Found ${Object.keys(staleLabelOutcomes).length} threads with stale "${procLbl.getName()}" labels. Applying final labels...`);
            applyFinalLabels(staleLabelOutcomes, procLbl, doneLbl, manualLbl);
             Logger.log("[INFO] FINISH: Label cleanup finished.");
        } else {
            Logger.log("[INFO] FINISH: No new messages and no label cleanup needed.");
        }
        const SCRIPT_END_TIME_NO_MSG = new Date();
        Logger.log(`==== SCRIPT EXECUTION FINISHED (${SCRIPT_END_TIME_NO_MSG.toLocaleString()}) ==== Time Elapsed: ${(SCRIPT_END_TIME_NO_MSG - SCRIPT_START_TIME)/1000}s ====`);
        return;
    }

    // Sort new messages chronologically by email date (oldest first)
    msgToSort.sort((a, b) => a.date - b.date);
    Logger.log(`[INFO] PROCESS: Sorted ${msgToSort.length} new messages by date. Starting main processing loop...`);


    // --- 3. Process Messages Loop ---
    let threadOutcomes = {}; // Stores the final outcome ('done' or 'manual') for each thread ID processed
    let procCount = 0;       // Messages successfully written to sheet
    let updateCount = 0;     // Rows updated
    let newCount = 0;        // Rows newly appended
    let errorCount = 0;      // Messages that failed processing

    for (let i = 0; i < msgToSort.length; i++) {
        const entry = msgToSort[i];
        const { message, date: emailDate, threadId } = entry; // Get the email's received date
        const msgId = message.getId();
        const processingStartTime = new Date(); // Time this specific message processing started
        if (DEBUG_MODE) Logger.log(`\n--- Processing Message ${i + 1}/${msgToSort.length} (Msg ID: ${msgId}, Thread: ${threadId}, Email Date: ${emailDate.toLocaleString()}) ---`);

        let needsManualReviewDueToParsing = false; // Flag if company/title extraction failed
        let finalStatusFromBody = null;            // Status detected from keywords
        let sheetWriteSuccessful = false;          // Flag if write/update to sheet succeeded
        let plainBody = null;                      // Store fetched body
        let currentMsgOutcome = 'manual';          // Default outcome to 'manual' unless successful

        try {
            // --- A. Extract Basic Info ---
            const emailSubj = message.getSubject() || ""; // Default to empty string if null
            const sender = message.getFrom() || "";
            const emailLnk = `https://mail.google.com/mail/u/0/#inbox/${msgId}`; // Direct link
            const processedTimestamp = new Date(); // Timestamp for *this script action*

            // --- B. Platform Detection ---
            let plat = DEFAULT_PLATFORM; // Start with default
            try {
                const emailMatch = sender.match(/<([^>]+)>/); // Extract email address
                if (emailMatch && emailMatch[1]) {
                    const emailAddress = emailMatch[1];
                    const domainParts = emailAddress.split('@');
                    if (domainParts.length === 2) {
                        const domain = domainParts[1].toLowerCase();
                        // Match against keywords defined in PLATFORM_DOMAIN_KEYWORDS
                        for (const keyword in PLATFORM_DOMAIN_KEYWORDS) {
                            // Check if domain *contains* the keyword (more flexible than exact match)
                            if (domain.includes(keyword)) {
                                plat = PLATFORM_DOMAIN_KEYWORDS[keyword];
                                if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Platform detected from domain keyword "${keyword}" in "${domain}": ${plat}`);
                                break; // Stop after first match
                            }
                        }
                    }
                }
            } catch (e) {
                Logger.log(`[WARN] Error during platform domain keyword detection for sender "${sender}": ${e}`);
                // Platform remains DEFAULT_PLATFORM
            }
             if (DEBUG_MODE && plat === DEFAULT_PLATFORM) Logger.log(`[DEBUG]  -> Platform detection complete. No specific keyword matched. Using default: ${plat}`);


            // --- C. Fetch Body ONCE ---
             try {
                 plainBody = message.getPlainBody(); // Fetch body text
                 if (DEBUG_MODE) Logger.log(`[DEBUG] Fetched Plain Body (Length: ${plainBody ? plainBody.length : 'null'}).`);
             } catch (e) {
                 // Log warning but continue, parsing might still work with subject/sender
                 Logger.log(`[WARN] Failed to get plain body for Msg ${msgId}: ${e}. Parsing will rely on subject/sender.`);
                 plainBody = null; // Ensure body is null if fetch failed
             }

            // --- D. Extract Company & Title (Pass potentially null body) ---
            const extracted = extractCompanyAndTitle(message, plat, emailSubj, plainBody); // Pass plainBody
            let company = extracted.company;
            let title = extracted.title;
            needsManualReviewDueToParsing = (company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED);
            if (needsManualReviewDueToParsing && DEBUG_MODE) {
                 Logger.log(`[DEBUG]  -> Parsing Result: Company or Title extraction failed or required manual review.`);
            }

            // --- E. Detect Status from Body (Pass potentially null body) ---
            finalStatusFromBody = parseBodyForStatus(plainBody); // Pass plainBody
            if (DEBUG_MODE) Logger.log(`[DEBUG] Body Status Detection complete. Status Found: ${finalStatusFromBody || 'None'}`);

            // --- F. Find Existing Row based on Company ---
            const lookupKey = company.toLowerCase();
            let existingRowData = null; // Will store { row, emailId, company, title } of the latest match
            let targetRow = -1; // Target row number for update, -1 if new

            if (company !== MANUAL_REVIEW_NEEDED && existD[lookupKey]) {
                // We have potential matches for this company. Use the most recent one (first in sorted list).
                existingRowData = existD[lookupKey][0];
                targetRow = existingRowData.row;
                if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Found existing entry for company "${company}" (lookup key "${lookupKey}"). Latest is Row ${targetRow}.`);
            } else {
                if (DEBUG_MODE && company === MANUAL_REVIEW_NEEDED) Logger.log(`[DEBUG]  -> Cannot search for existing row because company parsing failed.`);
                else if (DEBUG_MODE) Logger.log(`[DEBUG]  -> No existing row found for company "${company}" (lookup key "${lookupKey}"). Will create NEW row.`);
            }

            // --- G. Write to Sheet (Update or Append) ---
            if (targetRow !== -1 && existingRowData) {
                // ##### UPDATE EXISTING ROW #####
                if (DEBUG_MODE) Logger.log(`[DEBUG]    -> Preparing UPDATE for Row ${targetRow}...`);
                try {
                    const rangeToUpdate = sheet.getRange(targetRow, 1, 1, TOTAL_COLUMNS_IN_SHEET);
                    const currentValues = rangeToUpdate.getValues()[0]; // Get current data for comparison
                    let updated = false; // Flag to track if any significant change occurred
                    let newValues = [...currentValues]; // Create a copy to modify

                    // --- Logic to decide whether to update fields ---
                    // Use parsed Company/Title ONLY IF they are valid (not MANUAL_REVIEW_NEEDED). Otherwise, keep existing.
                    let finalCompany = (company !== MANUAL_REVIEW_NEEDED) ? company : existingRowData.company;
                    let finalTitle = (title !== MANUAL_REVIEW_NEEDED) ? title : existingRowData.title;

                    // Update Status if a new status was detected from the body AND it's different
                    if (STATUS_COL > 0 && finalStatusFromBody && finalStatusFromBody !== newValues[STATUS_COL - 1]) {
                        newValues[STATUS_COL - 1] = finalStatusFromBody;
                        updated = true;
                        if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Updating Status: "${currentValues[STATUS_COL - 1]}" -> "${finalStatusFromBody}"`);
                    }
                    // Update Company if the newly parsed (valid) company differs from existing
                    if (COMPANY_COL > 0 && finalCompany !== newValues[COMPANY_COL - 1]) {
                        newValues[COMPANY_COL - 1] = finalCompany;
                        updated = true;
                         if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Updating Company: "${currentValues[COMPANY_COL - 1]}" -> "${finalCompany}"`);
                    }
                     // Update Title if the newly parsed (valid) title differs from existing
                    if (JOB_TITLE_COL > 0 && finalTitle !== newValues[JOB_TITLE_COL - 1]) {
                        newValues[JOB_TITLE_COL - 1] = finalTitle;
                        updated = true;
                        if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Updating Title: "${currentValues[JOB_TITLE_COL - 1]}" -> "${finalTitle}"`);
                    }
                    // Update Platform if it differs
                     if (PLATFORM_COL > 0 && plat !== newValues[PLATFORM_COL - 1]) {
                        newValues[PLATFORM_COL - 1] = plat;
                        updated = true;
                        if (DEBUG_MODE) Logger.log(`[DEBUG]     -> Updating Platform: "${currentValues[PLATFORM_COL - 1]}" -> "${plat}"`);
                    }

                    // --- Always update these administrative fields for an update email ---
                     if (EMAIL_DATE_COL > 0)         newValues[EMAIL_DATE_COL - 1]         = emailDate;           // Update email date to the date of THIS email
                     if (LAST_UPDATE_DATE_COL > 0)   newValues[LAST_UPDATE_DATE_COL - 1]   = emailDate;           // Use THIS email's date as the last update trigger
                     if (EMAIL_SUBJECT_COL > 0)      newValues[EMAIL_SUBJECT_COL - 1]      = emailSubj;           // Update subject to THIS email's subject
                     if (EMAIL_LINK_COL > 0)         newValues[EMAIL_LINK_COL - 1]         = emailLnk;            // Update link to THIS email
                     if (EMAIL_ID_COL > 0)           newValues[EMAIL_ID_COL - 1]           = msgId;               // Update email ID to THIS email's ID
                     if (PROCESSED_TIMESTAMP_COL > 0) newValues[PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp; // Update script process time
                     updated = true; // Always consider an update email as causing a change needing write

                     // Perform the write only if changes were detected/required
                     if (updated) {
                        rangeToUpdate.setValues([newValues]);
                        Logger.log(`[INFO]    -> SHEET WRITE: Successfully updated Row ${targetRow}.`);
                        updateCount++;
                        sheetWriteSuccessful = true;
                        // Update the cache entry for this row with potentially new details
                        existingRowData.emailId = msgId; // Update the ID associated with this row
                        existingRowData.company = finalCompany; // Update cached company
                        existingRowData.title = finalTitle; // Update cached title
                     } else {
                        // This case is less likely now since admin fields always update
                        if (DEBUG_MODE) Logger.log(`[DEBUG]    -> SHEET WRITE: No significant changes detected for Row ${targetRow}. Skipping write.`);
                        sheetWriteSuccessful = true; // Still considered success in terms of processing flow
                     }
                } catch (e) {
                    Logger.log(`[ERROR] SHEET WRITE: Failed updating Sheet Row ${targetRow}: ${e}\nStack: ${e.stack}`);
                    sheetWriteSuccessful = false; // Mark as failed
                }

            } else {
                // ##### APPEND NEW ROW #####
                const initialStatus = finalStatusFromBody || DEFAULT_STATUS; // Use detected status or default
                if (DEBUG_MODE) Logger.log(`[DEBUG]    -> Preparing APPEND new row. Initial Status: ${initialStatus}`);
                try {
                     // Create array matching the total number of columns configured
                     const newRowData = new Array(TOTAL_COLUMNS_IN_SHEET).fill(""); // Initialize with empty strings

                     // Populate array using configured column indices (check > 0)
                     if (PROCESSED_TIMESTAMP_COL > 0) newRowData[PROCESSED_TIMESTAMP_COL - 1] = processedTimestamp;
                     if (EMAIL_DATE_COL > 0)         newRowData[EMAIL_DATE_COL - 1]         = emailDate; // Date email was received
                     if (PLATFORM_COL > 0)           newRowData[PLATFORM_COL - 1]           = plat;
                     if (COMPANY_COL > 0)            newRowData[COMPANY_COL - 1]            = company; // Use parsed company (could be MANUAL_REVIEW_NEEDED)
                     if (JOB_TITLE_COL > 0)          newRowData[JOB_TITLE_COL - 1]          = title;   // Use parsed title (could be MANUAL_REVIEW_NEEDED)
                     if (STATUS_COL > 0)             newRowData[STATUS_COL - 1]             = initialStatus;
                     if (LAST_UPDATE_DATE_COL > 0)   newRowData[LAST_UPDATE_DATE_COL - 1]   = emailDate; // Initial last update is the email's date
                     if (EMAIL_SUBJECT_COL > 0)      newRowData[EMAIL_SUBJECT_COL - 1]      = emailSubj;
                     if (EMAIL_LINK_COL > 0)         newRowData[EMAIL_LINK_COL - 1]         = emailLnk;
                     if (EMAIL_ID_COL > 0)           newRowData[EMAIL_ID_COL - 1]           = msgId;

                     sheet.appendRow(newRowData);
                     const newRowNum = sheet.getLastRow(); // Get the row number of the appended row
                     Logger.log(`[INFO]    -> SHEET WRITE: Successfully appended new Row ${newRowNum}.`);
                     newCount++;
                     sheetWriteSuccessful = true;

                     // Add this new entry to the cache if company was valid
                     const newCacheKey = company.toLowerCase();
                     if (company !== MANUAL_REVIEW_NEEDED && newCacheKey !== 'n/a') {
                         const newCacheEntry = { row: newRowNum, emailId: msgId, company: company, title: title };
                         if (!existD[newCacheKey]) {
                             existD[newCacheKey] = [];
                         }
                         // Add to front of cache list (most recent)
                         existD[newCacheKey].unshift(newCacheEntry);
                     }

                } catch (e) {
                    Logger.log(`[ERROR] SHEET WRITE: Failed appending new row to sheet: ${e}\nStack: ${e.stack}`);
                    sheetWriteSuccessful = false; // Mark as failed
                }
            } // --- End Update/Append Logic ---

            // --- H. Determine Outcome & Update Caches ---
            if (sheetWriteSuccessful) {
                procCount++; // Increment successful processing count
                existIds.add(msgId); // Add this message ID to processed set immediately
                // Determine final outcome for the thread based on this message
                currentMsgOutcome = (needsManualReviewDueToParsing) ? 'manual' : 'done';
                if (DEBUG_MODE) Logger.log(`[DEBUG]  -> Message Outcome: SUCCESS. Setting thread outcome contribution to: ${currentMsgOutcome}.`);
            } else {
                errorCount++; // Increment error count
                currentMsgOutcome = 'manual'; // Ensure failure leads to manual review
                Logger.log(`[ERROR]  -> Message Outcome: FAILED (Sheet Write Error). Setting thread outcome contribution to: ${currentMsgOutcome}.`);
            }

             // Update the overall thread outcome: If any message in the thread results in 'manual', the whole thread is 'manual'.
             // Otherwise, if all messages are 'done', the thread is 'done'.
            if (!threadOutcomes[threadId] || threadOutcomes[threadId] === 'done') {
                 threadOutcomes[threadId] = currentMsgOutcome;
                 if (DEBUG_MODE && currentMsgOutcome === 'manual') Logger.log(`[DEBUG]   -> Thread ${threadId} outcome updated to 'manual' due to this message.`);
            } else {
                 // Thread outcome is already 'manual', no change needed.
                 if (DEBUG_MODE) Logger.log(`[DEBUG]   -> Thread ${threadId} outcome remains 'manual'.`);
            }


        } catch (e) {
            // Catch any unexpected errors during the processing of a single message
            Logger.log(`[FATAL ERROR] Unexpected error processing Message ${msgId} (Thread ${threadId}): ${e}\nStack: ${e.stack}`);
            threadOutcomes[threadId] = 'manual'; // Mark thread for manual review on fatal error
            errorCount++;
        } finally {
            // Log time taken for this message if debugging
            if (DEBUG_MODE) {
                const processingEndTime = new Date();
                Logger.log(`--- Finished Message ${i + 1}/${msgToSort.length} --- Time: ${(processingEndTime-processingStartTime)/1000}s ---`);
            }
            // Optional: Add a small sleep even after successful message processing?
            Utilities.sleep(100); // Short pause between messages
        }
    } // --- End Message Processing Loop ---

    Logger.log(`\n[INFO] PROCESS: Finished main processing loop.`);
    Logger.log(`[INFO] SUMMARY: Messages Processed & Written: ${procCount}, Updates: ${updateCount}, New Rows: ${newCount}, Message Errors: ${errorCount}.`);
    if (DEBUG_MODE) Logger.log(`[DEBUG] Final Thread Labeling Outcomes Determined: ${JSON.stringify(threadOutcomes)}`);

    // --- 4. Apply Final Gmail Labels based on aggregated thread outcomes ---
    applyFinalLabels(threadOutcomes, procLbl, doneLbl, manualLbl);

    // --- Completion ---
    const SCRIPT_END_TIME = new Date();
    Logger.log(`==== SCRIPT EXECUTION FINISHED (${SCRIPT_END_TIME.toLocaleString()}) === Time Elapsed: ${(SCRIPT_END_TIME - SCRIPT_START_TIME)/1000}s ===`);

} // --- End processJobApplicationEmails ---


// --- Trigger Setup Function ---
// This function is called by the menu item "2. Setup/Verify Hourly Trigger"
function createTimeDrivenTrigger(){
    const triggers = ScriptApp.getProjectTriggers();
    let triggerExists = false;
    const handlerFunction = 'processJobApplicationEmails'; // The function the trigger should run

    for(const trigger of triggers){
        if(trigger.getHandlerFunction() === handlerFunction){
            triggerExists = true;
            Logger.log(`[INFO] Time-driven trigger for function "${handlerFunction}" already exists.`);
            break;
        }
    }

    if(!triggerExists){
        Logger.log(`[INFO] No existing trigger found for "${handlerFunction}". Attempting to create hourly trigger...`);
        try{
            ScriptApp.newTrigger(handlerFunction)
                .timeBased()
                .everyHours(1) // Run every hour
                .create();
            Logger.log(`[INFO] Successfully created time-driven trigger for "${handlerFunction}" to run every hour.`);
            SpreadsheetApp.getUi().alert(`Hourly trigger for "${handlerFunction}" created successfully.`);
        } catch(e) {
            Logger.log(`[ERROR] Failed to create time-driven trigger: ${e}`);
            SpreadsheetApp.getUi().alert(`Failed to create trigger. You may need to set it up manually in the Script Editor > Triggers menu. Error: ${e.message}. Check script logs for details.`);
        }
    } else {
        SpreadsheetApp.getUi().alert(`An hourly trigger for "${handlerFunction}" already exists.`);
    }
}

// --- Menu Function ---
// Creates the custom menu in the Google Sheet when the sheet is opened.
function onOpen() {
    // Use getActiveSpreadsheet to ensure it works correctly when bound
    SpreadsheetApp.getUi()
        .createMenu('Job Processor (v8.1)') // Updated menu name to reflect version/template
        .addItem('1. Process Emails Now', 'processJobApplicationEmails')
        .addItem('2. Setup/Verify Hourly Trigger', 'createTimeDrivenTrigger')
        .addToUi();
     if (DEBUG_MODE) Logger.log("[DEBUG] Custom menu 'Job Processor (v8.1)' added to the Sheet UI.");
}
