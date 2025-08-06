# Career Suite AI v1 **REVISED PROJECT NAME**
# Automated-Job-Application-Tracker-Pipeline-Manager
Automate job tracking w/ Google Apps Script! Parses Gmail emails (company, title, status keywords) via regex &amp; logs to Sheets. Companion script auto-rejects stale applications after configurable inactivity period. Keep your pipeline clean!

# Automated Job Application Tracker & Pipeline Manager (Google Apps Script)

## Overview

This project provides a powerful, automated system for tracking job applications directly within your Google Workspace (Gmail + Google Sheets). It addresses the common pain points of manually logging application details, keeping track of statuses across numerous emails, and managing a cluttered pipeline often filled with unresponsive applications ("ghosting").

The system consists of two core Google Apps Scripts:

1.  **`job_application_email_parser.gs`**: This script connects to your Google Sheet. It monitors a specific label in your Gmail (e.g., "AppToProcess"), intelligently parses incoming emails for key details (Company, Job Title, Status, Platform) using configurable keywords and patterns, and automatically logs this information to your Google Sheet. It updates existing entries or creates new ones, managing Gmail labels (`AppDONEProcess`, `ManualReviewNeeded`) to reflect the processing status.
2.  **`auto_reject_stale_applications.gs`**: This standalone script runs periodically (e.g., daily) to review your application tracking Sheet. It identifies applications that haven't received an update email within a configurable timeframe (e.g., 7 weeks) and automatically changes their status to "Rejected" (or your configured term), helping to keep your active pipeline relevant.

Together, these scripts aim to save significant time, reduce manual errors, provide a clear, centralized dashboard, and empower you to manage your job search more effectively.

## Features ✨

*   **Automated Logging:** Extracts Company, Job Title, Status, Platform, and more from emails.
*   **Centralized Tracking:** Consolidates all application data into a single Google Sheet.
*   **Intelligent Parsing:** Uses keywords and regex to determine application status (Applied, Interview, Rejected, etc.).
*   **Platform Detection:** Attempts to identify the source platform (LinkedIn, Greenhouse, Lever, etc.) from sender details.
*   **Update Handling:** Recognizes follow-up emails for existing applications and updates the relevant row.
*   **Stale Application Management:** Automatically marks inactive applications as rejected after a set period.
*   **Gmail Workflow:** Uses labels (`AppToProcess`, `AppDONEProcess`, `ManualReviewNeeded`) for clear processing status in your inbox.
*   **Customizable:** Configure labels, sheet names, column positions, status keywords, rejection timelines, and more.
*   **Gmail Filter Integration:** Works seamlessly with Gmail filters for automatic labeling.

## How it Works ⚙️

1.  **(Optional but Recommended):** A Gmail filter automatically applies the `AppToProcess` label to incoming emails from common job boards/ATS systems (or you manually label relevant emails).
2.  **(Hourly):** The `job_application_email_parser.gs` script runs via a time-driven trigger.
3.  It finds all emails with the `AppToProcess` label in your Gmail.
4.  For each new email, it parses the sender, subject, and body to extract application details.
5.  It checks the Google Sheet to see if an application for that company already exists.
6.  It either **updates** the existing row (updating status, email details, last update date) or **appends** a new row for a new application.
7.  It removes the `AppToProcess` label and adds either `AppDONEProcess` (if successful) or `ManualReviewNeeded` (if parsing failed or an error occurred).
8.  **(Daily):** The `auto_reject_stale_applications.gs` script runs via a time-driven trigger.
9.  It reads the Google Sheet, checking the `Last Update Date` column for rows with "active" statuses.
10. If an application's `Last Update Date` is older than the configured threshold (e.g., 7 weeks) and its status isn't considered final (e.g., "Offer/Accepted", "Withdrawn"), it changes the status to "Rejected".
11. **You** view the up-to-date Google Sheet to manage your job search!

## Requirements 📋

*   A Google Account (Gmail + Google Drive/Sheets).
*   The two script files (`job_application_email_parser.gs`, `auto_reject_stale_applications.gs`).
*   A Google Sheet set up with the correct columns (see Setup).
*   (Optional but helpful) Basic familiarity with Google Sheets and Gmail filters.

## Setup Instructions 🚀

Follow these steps carefully to get the system running:

**Step 0: Preparation**

*   **Get the Scripts:** Download or copy the code for both `job_application_email_parser.gs` and `auto_reject_stale_applications.gs` from this repository.
*   **Prepare Google Sheet Structure:** You need a Google Sheet with specific columns in the correct order. The required headers are:
    `Processed Timestamp,Email Date,Platform,Company,Job Title,Status,Last Update Date,Email Subject,Email Link,Email ID`
    *   **Action:** Create a new Google Sheet.
    *   **Action:** In the first row (Row 1), enter these exact headers, one per column, starting from Column A.
    *   **(Highly Recommended):** *If a template sheet link is provided with this repository, make a copy of that template instead of creating a blank sheet.* ([**TODO:** Add link to your template sheet here if you create one]).
    *   **What this does:** Ensures the scripts read from and write to the correct columns based on their configured numbers (e.g., `COMPANY_COL = 4` expects 'Company' to be in Column D).

**Step 1: Setup Script 1 (`job_application_email_parser.gs`)**

*   **Action:** Open the Google Sheet you prepared in Step 0.
*   **Action:** Go to the menu: "Extensions" > "Apps Script".
    *   **What this does:** Opens the Google Apps Script editor *bound* to your spreadsheet. This allows the script to directly interact with this specific sheet using functions like `getActiveSpreadsheet()`.
*   **Action:** If there's any default code (like `function myFunction() {}`), delete it.
*   **Action:** Paste the entire code from `job_application_email_parser.gs` into the editor.
*   **Action:** **Configure Script 1 Constants:** Near the top of the script, find the `USER CONFIGURATION SECTION`. Carefully review and **edit** the constants if needed:
    *   `SHEET_NAME`: Ensure this matches the name of the tab in your Google Sheet (default is "Applications").
    *   `GMAIL_LABEL_...`: Verify these (`AppToProcess`, `AppDONEProcess`, `ManualReviewNeeded`) are the labels you want to use. You'll create these in Gmail later.
    *   `Column Indices (...)`: **CRITICAL:** If you did *not* use the exact column order from Step 0, you MUST update these numbers (A=1, B=2, etc.) to match your sheet layout.
    *   (Other constants like keywords can be adjusted later, defaults are usually fine to start).
    *   **What this does:** Tailors the script's behavior to match your specific sheet name, desired Gmail labels, and sheet column structure.
*   **Action:** Click the "Save project" icon (💾). Give the script project a name (e.g., "Job Application Email Parser").

**Step 2: Setup Script 2 (`auto_reject_stale_applications.gs`)**

*   **Action:** Go to [script.google.com](https://script.google.com/home/my) in a new browser tab.
*   **Action:** Click "+ New project".
    *   **What this does:** Creates a *standalone* script project, not directly attached to any specific Sheet initially. This is necessary because this script identifies the sheet by its ID.
*   **Action:** Delete any default code.
*   **Action:** Paste the entire code from `auto_reject_stale_applications.gs` into the editor.
*   **Action:** **Configure Script 2 Constants:** Near the top, find the `USER CONFIGURATION SECTION`. Carefully review and **edit** these constants:
    *   `SPREADSHEET_ID`: **CRITICAL:** Replace `"YOUR_SPREADSHEET_ID_HERE"` with the actual ID of your Google Sheet from Step 0. Find the ID in the Sheet's URL: `.../spreadsheets/d/THIS_IS_THE_ID/edit...`.
    *   `SHEET_NAME`: Ensure this matches the sheet tab name (e.g., "Applications").
    *   `WEEKS_THRESHOLD`: Adjust if you want a different period before marking as stale (default is 7).
    *   `REJECTED_STATUS`, `FINAL_STATUSES`: Ensure these match the status terms used in your sheet/Script 1.
    *   `Column Indices (...)`: **CRITICAL:** Ensure `STATUS_COL` and `LAST_UPDATE_DATE_COL` match your sheet layout (defaults F=6, G=7).
    *   **What this does:** Tells the standalone script *which* spreadsheet and tab to operate on, how long to wait before rejecting, what status terms to use, and which columns contain the necessary data.
*   **Action:** Click the "Save project" icon (💾). Give it a name (e.g., "Stale Application Rejector").

**Step 3: Setup Gmail Labels & Filter**

*   **Action:** Open your Gmail.
*   **Action:** Create the three Gmail labels you configured in Script 1 (e.g., `AppToProcess`, `AppDONEProcess`, `ManualReviewNeeded`). You can create labels by scrolling down the left sidebar, clicking "More", then "Create new label".
    *   **What this does:** Creates the necessary "folders" for the scripts to use for managing the email processing workflow.
*   **Action:** Set up the Gmail Filter for Automatic Labeling (Recommended):
    *   Go to Gmail Settings (⚙️ icon) > "See all settings".
    *   Go to the "Filters and Blocked Addresses" tab.
    *   Click "Create a new filter".
    *   In the `From` field, paste the following (or adjust based on the emails you receive):
        `indeedapply@indeed.com OR notifications@linkedin.com OR donotreply@workday.com OR no-reply@greenhouse.io OR greenhouse.io OR lever.co OR myworkday.com OR icims.com OR wellfound.com OR angel.co OR hi.wellfound.com OR linkedin.com OR indeed.com OR ziprecruiter.com OR glassdoor.com OR hired.com`
    *   Click "Create filter".
    *   Check the box for `Skip the Inbox (Archive it)`.
    *   Check the box for `Apply the label:` and select your processing label (e.g., `AppToProcess`).
    *   Click "Create filter".
    *   **What this does:** Automatically labels incoming emails from common job platforms with `AppToProcess` and archives them, keeping your inbox cleaner and feeding emails directly into the script's queue.

**Step 4: Authorize the Scripts**

*   **Authorize Script 1:**
    *   **Action:** Go back to your Google Sheet. **Refresh the page.**
    *   **Action:** A new menu item (e.g., "Job Processor (v8.1)") should appear. Click it, then select "1. Process Emails Now".
    *   **Action:** An "Authorization Required" dialog will pop up. Click "Continue".
    *   **Action:** Choose your Google Account.
    *   **Action:** You might see a "Google hasn't verified this app" screen. Click "Advanced", then "Go to [Your Script Name] (unsafe)". (This is normal for personal scripts).
    *   **Action:** Review the permissions (it needs access to Gmail and Sheets). Click "Allow".
    *   **What this does:** Grants Script 1 permission to read your emails (with specific labels), modify labels, and read/write data to the Google Sheet it's attached to.
*   **Authorize Script 2:**
    *   **Action:** Go back to the Apps Script editor for `auto_reject_stale_applications.gs` (from Step 2).
    *   **Action:** In the editor's menu bar, select the function `markStaleApplicationsAsRejected` from the dropdown list next to the Debug (bug) icon.
    *   **Action:** Click the "Run" button (looks like a play icon ▶️).
    *   **Action:** Follow the same "Authorization Required" steps as above (Continue > Choose Account > Advanced > Go to... > Allow).
    *   **What this does:** Grants Script 2 permission to access and modify the specific Google Sheet you identified by its ID in the configuration.

**Step 5: Set Up Automatic Triggers**

*   **Script 1 Trigger (Hourly):**
    *   **Action:** In the Apps Script editor for `job_application_email_parser.gs`, click the "Triggers" icon (looks like a clock ⏰) on the left sidebar.
    *   **Action:** Click the "+ Add Trigger" button (bottom right).
    *   **Action:** Configure the settings:
        *   Choose which function to run: `processJobApplicationEmails`
        *   Choose which deployment should run: `Head`
        *   Select event source: `Time-driven`
        *   Select type of time based trigger: `Hourly timer`
        *   Error notification settings: Select `Notify me immediately` (Recommended).
    *   **Action:** Click "Save". You might need to authorize again briefly.
    *   **What this does:** Sets up Script 1 to run automatically every hour to process newly labeled emails.
*   **Script 2 Trigger (Daily):**
    *   **Action:** In the Apps Script editor for `auto_reject_stale_applications.gs`, click the "Triggers" icon (⏰).
    *   **Action:** Click "+ Add Trigger".
    *   **Action:** Configure the settings:
        *   Choose which function to run: `markStaleApplicationsAsRejected`
        *   Choose which deployment should run: `Head`
        *   Select event source: `Time-driven`
        *   Select type of time based trigger: `Daily timer`
        *   Select time of day: Choose a time when you're unlikely to be actively using the sheet (e.g., 2am - 3am).
        *   Error notification settings: Select `Notify me immediately`.
    *   **Action:** Click "Save".
    *   **What this does:** Sets up Script 2 to run automatically once a day to check for and update stale applications.

**Setup Complete! 🎉** The system should now be running automatically.

## Configuration Details 🔧

While the setup covers the basics, you can fine-tune the scripts via the constants in their respective `USER CONFIGURATION SECTION`s:

**`job_application_email_parser.gs`:**

*   `DEBUG_MODE`: Set to `true` for verbose logging in Apps Script Logs (useful for troubleshooting).
*   `SHEET_NAME`: Target sheet tab name.
*   `GMAIL_LABEL_*`: Names of labels used for the workflow.
*   `*_COL` Constants: **Must match your sheet layout (A=1, B=2...).**
*   `TOTAL_COLUMNS_IN_SHEET`: Should be >= the highest column index number used.
*   `DEFAULT_STATUS`, `REJECTED_STATUS`, etc.: The exact text strings for different application statuses.
*   `MANUAL_REVIEW_NEEDED`: Text used when parsing fails.
*   `DEFAULT_PLATFORM`: Fallback platform name.
*   `*_KEYWORDS`: Lists of lowercase words/phrases used to detect status changes in email bodies. **Expand these for better accuracy!**
*   `PLATFORM_DOMAIN_KEYWORDS`: Maps keywords found in sender email domains to platform names.
*   `IGNORED_DOMAINS`: Domains (like `gmail.com`, `lever.co`) that shouldn't be used as the primary Company Name when parsing the sender.

**`auto_reject_stale_applications.gs`:**

*   `SPREADSHEET_ID`: **Must be the correct ID of your target Google Sheet.**
*   `SHEET_NAME`: Target sheet tab name (must match Script 1).
*   `WEEKS_THRESHOLD`: How many weeks of inactivity before marking as stale.
*   `STATUS_COL`, `LAST_UPDATE_DATE_COL`: Column numbers for Status and Last Update Date in your sheet.
*   `REJECTED_STATUS`: The exact status text to apply when an application becomes stale.
*   `FINAL_STATUSES`: A `Set` containing status strings (e.g., "Offer/Accepted", "Withdrawn", your `REJECTED_STATUS`) that should *not* be overwritten by the stale rejection logic.

## Usage 🧑‍💻

1.  **Label Emails:** Either let the Gmail filter automatically apply the `AppToProcess` label, or manually apply it to relevant job application emails (initial applications, updates, rejections, interview invites).
2.  **Wait:** The hourly trigger will run `job_application_email_parser.gs` to process labeled emails.
3.  **Check Sheet:** Review your Google Sheet periodically. New applications will be added, and existing ones updated.
4.  **Manual Review:** Check emails labeled `ManualReviewNeeded`. These likely had parsing issues (couldn't find Company/Title) or encountered an error. You may need to manually update the sheet for these.
5.  **Monitor Stale:** The daily trigger runs `auto_reject_stale_applications.gs` to automatically mark old, inactive applications as "Rejected".

## Troubleshooting / Notes ⚠️

*   **Permissions:** If scripts fail, double-check that you granted all necessary permissions during authorization. You might need to re-authorize via the steps above.
*   **Typos:** Carefully check `SHEET_NAME`, `SPREADSHEET_ID`, label names, and column numbers in the script configurations for typos. They must match *exactly*.
*   **Sheet Structure:** Ensure your Google Sheet columns match the `*_COL` constants defined in *both* scripts.
*   **Gmail Filter:** Ensure your Gmail filter uses the correct `AppToProcess` label name and that the `From:` criteria capture the emails you expect.
*   **Logs:** Check the script execution logs for errors. Go to the Apps Script editor > "Executions" (stopwatch icon ⏱️) to see run history and logs. `DEBUG_MODE = true` in Script 1 provides more detail.
*   **Quotas:** Google Apps Script has quotas (e.g., execution time per day). If you process hundreds of emails at once, you might hit limits. Processing hourly usually avoids this.
*   **Parsing Accuracy:** Email parsing relies on patterns and keywords. It won't be 100% perfect due to variations in email formats. Expect some entries to require manual review or correction. Improve accuracy by adding more keywords to the configuration.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
