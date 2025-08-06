# Career Suite AI v2 **REVISED PROJECT NAME**
# Automated Job Application Tracker & Pipeline Manager v2.0

**Author:** Francis John LiButti
**Version:** 2.0.0 (Based on Chrono v8 - Full Integration)
**Last Updated:** May 13, 2024 *(User: Please update this to the current date when you post)*

## Short Description

An enhanced Google Apps Script (v2.0) to automate your job application workflow. It parses Gmail emails for key application details, logs them to a structured Google Sheet (acting as a mini-CRM & pipeline manager), manages application stages with a nested Gmail label system, and includes a one-click setup for the sheet, labels, and automated triggers for processing and stale application flagging.

## Overview

This "Automated Job Application Tracker & Pipeline Manager v2.0" is a Google Apps Script designed to significantly streamline the often-tedious process of tracking job applications. By integrating directly with your Gmail and Google Sheets, it provides a centralized system to:

*   Automatically process emails related to your job applications that you label in Gmail.
*   Intelligently parse these emails for key information such as company name, job title, and the current application status (e.g., Applied, Interview Scheduled, Rejected) using customizable keyword lists.
*   Log all extracted details into a structured Google Sheet, which serves as your central application pipeline and mini-CRM.
*   Organize your Gmail using a nested label system specific to this tracker (default parent: "Assisted Job Application Tracker").
*   Automate the initial setup of necessary Gmail labels, the Google Sheet (complete with detailed formatting like frozen headers, specific column widths, text wrapping, and alternating row colors), and time-driven triggers for ongoing automation.
*   Offer flexibility in sheet configuration: use a specific, pre-existing Google Sheet by providing its ID, or let the script dynamically create and manage its own data sheet.
*   Automatically identify and update "stale" applications (e.g., those with no updates for several weeks) to a "Rejected" status, helping keep your pipeline current.

## Features

*   **Automated Email Parsing:** Extracts company, job title, and infers status from application-related emails based on customizable keywords.
*   **Google Sheet as CRM/Pipeline:** Creates and maintains a formatted spreadsheet for comprehensive application tracking.
*   **Nested Gmail Label System:** Provides clear visual organization of application emails within Gmail (default parent: `Assisted Job Application Tracker`).
*   **One-Click Initial Setup:** A dedicated menu command (`0. Initial Setup...`) automates the creation of labels, the sheet (with detailed styling), and essential time-driven triggers.
*   **Hybrid Sheet Management:** Supports using a user-specified `FIXED_SPREADSHEET_ID` or dynamic sheet creation/discovery by filename (with ID stored in `PropertiesService`).
*   **Detailed Sheet Formatting:** New data sheets are automatically formatted with:
    *   Bold, centered, and frozen header row.
    *   Generous default row heights and specific column widths for readability.
    *   Text wrapping for most data, with clipping specifically for the "Email Link" column.
    *   Alternating row colors (blue and white banding) for data rows.
    *   Automatic hiding of unused columns beyond those defined.
*   **Stale Application Flagging:** A function (and optional trigger) to automatically mark applications as "Rejected" if they haven't been updated beyond a configurable time threshold (default 7 weeks) and are not already in a pre-defined "final" status.
*   **Customizable Keyword Engine:** Relies on user-defined global keyword lists within the script for more accurate email parsing.
*   **Programmatic Trigger Setup:** Automates the creation of:
    *   An **hourly** trigger for processing new job application emails.
    *   A **daily** (around midnight) trigger for checking and marking stale applications.
*   **User-Friendly Menu:** Adds a custom menu ("Job App Pipeline v2 Tools") in Google Sheets for easy access to setup, manual processing, and trigger verification.

## Prerequisites

*   A Google Account (for Gmail, Google Sheets, Google Drive, and Apps Script).
*   Basic ability to copy/paste code into the Google Apps Script editor and authorize permissions.

## Setup Instructions

**1. Copy the Script Code:**
   *   Navigate to the `Code.gs` file (or the primary `.gs` file) in this repository.
   *   Select and copy all the code within that file.

**2. Open or Create a Google Sheet:**
   *   It's recommended to create a new, blank Google Sheet (e.g., name it "Job Tracker Control Panel"). This sheet will host the script's custom menu for easy access.
   *   Alternatively, you can use an existing sheet.

**3. Access the Apps Script Editor:**
   *   With your Google Sheet open, click on **Extensions > Apps Script**. This will open the editor bound to that sheet.

**4. Paste and Configure the Script:**
   *   In the Apps Script editor, if there's a `Code.gs` file with default content, delete that content.
   *   Paste the entire script code you copied from GitHub.
   *   **CRITICAL: Configure Constants (near the top of the script):**
        *   **`FIXED_SPREADSHEET_ID`**:
            *   **Option A (Use a specific existing sheet for data):** Find the ID of your existing Google Sheet (from its URL: `.../d/SHEET_ID_IS_HERE/edit`) and replace `"YOUR_SPREADSHEET_ID_HERE"` with that ID.
               ```javascript
               const FIXED_SPREADSHEET_ID = "123abcYourActualSheetID456xyz"; // Example
               ```
            *   **Option B (Let the script create/manage its data sheet - Recommended for new users):** Leave `FIXED_SPREADSHEET_ID` as `"YOUR_SPREADSHEET_ID_HERE"` or set it to `""`. The script will then use `TARGET_SPREADSHEET_FILENAME` (default: "Assisted Job Application Tracker Data").
        *   **Keyword Lists (`REJECTION_KEYWORDS`, `ACCEPTANCE_KEYWORDS`, etc.):** **VERY IMPORTANT.** Review and extensively customize these global constant lists to match the common phrasing and senders of emails you receive. Better, more comprehensive keyword lists significantly improve parsing accuracy.
        *   Review other constants like `GMAIL_LABEL_PARENT`, `TARGET_SPREADSHEET_FILENAME`, `WEEKS_THRESHOLD` if you wish to change their default values.

**5. Save the Script:**
   *   Click the "Save project" icon (floppy disk) in the Apps Script editor.
   *   Give the Apps Script project a name if prompted (e.g., "JobApplicationManagerV2").

**6. Run Initial Setup (One-Time Essential Step):**
   *   Go back to your Google Sheet (the one from Step 2). **Refresh the browser tab/window.**
   *   After a few moments, a new menu named **"Job App Pipeline v2 Tools"** should appear.
   *   Click **Job App Pipeline v2 Tools > 0. Initial Setup (Labels, Sheet, Triggers)**. (The function called is `initialSetup_Full`).
   *   **Authorization Required:**
        *   A dialog box will prompt for authorization. Click "Continue."
        *   Select your Google Account.
        *   If you see a "Google hasn’t verified this app" screen (this is normal for your own scripts), click **"Advanced"** (you may need to scroll), then click **"Go to [Your Script Name] (unsafe)"**.
        *   Review the list of permissions the script requests (access to Gmail for labels/reading emails, Sheets for data logging, Drive if dynamic sheet creation is enabled, and services for creating triggers). Click **"Allow"**.
   *   The setup function will execute. It will attempt to create Gmail labels, set up the Google Sheet with formatting, and create the automated triggers.
   *   An alert box should appear summarizing the setup actions. **Crucially, also check the script execution logs** for detailed results of each step (In Apps Script Editor: `View > Executions`, select the `initialSetup_Full` run).

**7. Verify Setup Components:**
    *   **Gmail:** Look for the parent label (default: "Assisted Job Application Tracker") and its nested children: "/To Process", "/Processed", "/Manual Review Needed".
    *   **Google Drive/Sheets:** Find your data spreadsheet. Open it and check for the "Applications" tab with the headers, frozen row, column styles, banding, etc.
    *   **Apps Script Triggers:** In the Apps Script editor, click the "Triggers" icon (alarm clock) on the left. Confirm the hourly trigger for `processJobApplicationEmails` and the daily trigger for `markStaleApplicationsAsRejected`.

## How to Use

**1. Label Emails in Gmail:**
   *   For any job application email (sent or received) that you want the script to process, apply the Gmail label: **`[Your Parent Label]/To Process`** (e.g., `Assisted Job Application Tracker/To Process`) to that email thread.

**2. Processing Application Emails:**
   *   **Manually:** From the Google Sheet menu, select **Job App Pipeline v2 Tools > 1. Process New Emails Now**.
   *   **Automatically:** The hourly trigger (if set up) will run `processJobApplicationEmails`.
   *   Data is logged/updated in the "Applications" sheet. Gmail labels on the thread are updated from `/To Process` to `/Processed` or `/Manual Review Needed`.

**3. Marking Stale Applications:**
   *   **Manually:** Select **Job App Pipeline v2 Tools > 2. Mark Stale Applications as Rejected**.
   *   **Automatically:** The daily trigger (if set up) runs `markStaleApplicationsAsRejected`.

**4. Managing Triggers:**
   *   The menu includes items (`3.` and `4.`) to verify or re-attempt the creation of specific triggers if you only want to manage one.
   *   View, edit, or delete triggers on the "Triggers" page in the Apps Script editor.

## Shortcomings and Considerations

*   **Parsing Accuracy:** Email content varies wildly. The script's parsing (based on keywords and regex) will not be perfect for all emails. **Extensive customization of the keyword lists is essential for good performance.** Expect to use the "/Manual Review Needed" label and refine keywords over time.
*   **Google Service Quotas:** For typical personal use, daily quotas for Gmail, Sheets, and script execution time (6 min/execution for consumer accounts) should not be an issue. Very high volumes might hit limits.
*   **Trigger Precision:** Time-driven triggers run approximately at the scheduled time; Google may randomize the exact minute slightly.
*   **Error Logs:** If something goes wrong, the first place to check is **View > Executions** in the Apps Script editor.
*   **Consolidation Placeholder:** `performProcessConsolidation_()` is a stub for custom logic.

## Future Ideas

*   More advanced NLP/AI for parsing (beyond Apps Script's current capabilities without external APIs).
*   UI for keyword management directly in the Sheet.
*   Direct integration for sending follow-up email templates.

---
