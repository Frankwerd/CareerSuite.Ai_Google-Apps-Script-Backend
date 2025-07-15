# CareerSuite.ai - Google Apps Script Backend

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

This repository contains the full open-source code for the server-side backend of the [CareerSuite.ai Chrome Extension](https://chrome.google.com/webstore/detail/careersuiteai/your-extension-id). This Google Apps Script project runs entirely within the user's own Google Account, ensuring that the user always owns and controls their data. It does not run on any developer-owned servers.

The primary purpose of this script is to automate the creation and management of a personal job application tracker within the user's Google Sheets by securely processing specific Gmail messages with their permission.

## Our Commitment to Privacy & User Control

The architecture of this backend is guided by a steadfast commitment to user privacy and data security.

-   **Data Sovereignty**: You, the user, always own and control 100% of your data. All data processed by this script—including the Google Sheet tracker, the content of your emails, and your personal API key—resides exclusively within your personal Google Account.
-   **No Developer Data Access**: The developers of CareerSuite.ai do not have access to, collect, or store any of the user's personally identifiable information (PII), resume content, email data, or API keys. The script acts as a secure automation tool that you install and run in your own cloud environment.
-   **Full Transparency**: This project is fully open-source to allow for public auditing and independent verification of all our data handling claims. We encourage you to review the code.

## Key Features

-   **Automated Sheet Creation**: On first authorization, the script copies a master template to create a new "CareerSuite.ai Data" spreadsheet in the user's own Google Drive.
-   **Automated Email Processing**: The script creates specific labels in your Gmail (e.g., `CareerSuite.AI/Applications/To Process`). A corresponding Gmail filter, also created by the script, automatically moves relevant job application emails to this label. The script reads *only* from this designated label, analyzes the emails, and then updates their labels to `.../Processed` to prevent re-processing.
-   **AI-Powered Data Extraction**: To accurately parse details like Company Name, Job Title, and Application Status from unstructured emails, the script utilizes the user's provided Google Gemini API key. This key is stored securely within the user's own account using Google's `PropertiesService`.
-   **Dashboard & Analytics**: The script automatically populates a "Dashboard" tab in the Google Sheet, providing users with at-a-glance metrics, charts, and visualizations about their job search progress.

## System Architecture

The following diagram illustrates the flow of data. All components under "User's Google Cloud Account" are created and managed by this script and are under the user's exclusive control.

```
[User's Browser]                                  [User's Google Cloud Account]
+------------------------------------+           +---------------------------------------------+
|                                    |           |                                             |
|  [Chrome Extension]                |  ----1--> |  [Google Apps Script Web App]               |
|   - UI (Popup, Options)            |           |   - doGet(e)                                |
|   - Securely stores user API Key   |           |   - doPost(e)                               |
|                                    |           |                                             |
+------------------------------------+           |             |                               |
                                                 |             |                               |
                                                 |             2                               |
                                                 |             |                               |
                                                 |   +-------------------+                     |
                                                 |   | [GeminiService.gs]| --------3--------> [Google Gemini API]
                                                 |   +-------------------+                     |
                                                 |             |                               |
                                                 |             4                               |
                                                 |             |                               |
                                                 |   +-------------------+   +-----------------+
                                                 |   |  [SheetUtils.gs]|-->| [Google Sheet]  |
                                                 |   +-------------------+   +-----------------+
                                                 |             |                               |
                                                 |             5                               |
                                                 |             |                               |
                                                 |   +-------------------+   +-----------------+
                                                 |   |   [GmailUtils.gs] |-->|   [Gmail]       |
                                                 |   +-------------------+   +-----------------+
                                                 |                                             |
                                                 +---------------------------------------------+
```

1.  **Authenticated Request**: The CareerSuite.ai Chrome Extension makes a secure `fetch` request to the Apps Script Web App endpoint (e.g., to create a sheet or sync an API key). This request is authenticated with the user's Google OAuth token, ensuring only the user can trigger their own backend.
2.  **Trigger Backend Logic**: The Apps Script Web App (`WebApp_Endpoints.gs`) receives the request and triggers the appropriate function, such as `runFullProjectInitialSetup` or `processJobApplicationEmails`.
3.  **AI Analysis**: The `GeminiService.gs` module sends the text of a relevant email to the Google Gemini API for parsing. This request is made using the user's *own* Gemini API key, which is stored securely in their account's `UserProperties`.
4.  **Sheet Update**: The script, via `SheetUtils.gs` and `Main.gs`, writes the processed data (e.g., Company Name, Job Title, Status) into the user's private "CareerSuite.ai Data" Google Sheet.
5.  **Gmail State Management**: To complete the cycle and prevent duplicate processing, `GmailUtils.gs` modifies the email's label from `.../To Process` to `.../Processed`.

## Project File Structure

*   `Main.gs`: The central orchestration file. It contains the `onOpen()` function to create the spreadsheet menu, the primary `runFullProjectInitialSetup()` function that calls other modules, and the main email processing loop `processJobApplicationEmails()`.
*   `WebApp_Endpoints.gs`: Handles all incoming HTTP `doGet` and `doPost` requests from the companion Chrome Extension. This is the primary entry point for the extension to communicate with the backend.
*   `Config.gs`: A centralized configuration file containing all global constants, such as sheet names, column headers, status types, AI model endpoints, and Gmail label names.
*   `GeminiService.gs`: Manages all interactions with the Google Gemini API. It constructs the prompts, sends the requests for email parsing, and handles the responses.
*   `SheetUtils.gs`: A collection of utility functions for interacting with Google Sheets, including creating new sheets, applying formatting, and managing data ranges.
*   `GmailUtils.gs`: Contains helper functions for interacting with Gmail, primarily for creating and managing labels (`getOrCreateLabel`).
*   `Leads_Main.gs`: Contains the primary functions for the Job Leads Tracker module, including initial setup of the leads sheet/labels/filters and the ongoing processing of job lead emails.
*   `Leads_SheetUtils.gs`: Contains utility functions specifically for the "Potential Job Leads" sheet, such as writing new job data, retrieving processed email IDs, and mapping column headers.
*   `Dashboard.gs`: Manages the creation, formatting, and data population of the "Dashboard" and "DashboardHelperData" sheets, including chart creation and formula setup.
*   `ParsingUtils.gs`: Contains functions dedicated to parsing email content (subject, body, sender) using regular expressions and keyword matching as a fallback or supplement to AI parsing.
*   `Triggers.gs`: Includes functions for creating, verifying, and managing the time-driven triggers that automate the script's execution (e.g., checking for new emails every hour).
*   `AdminUtils.gs`: Provides utility functions for project setup and configuration, such as managing API keys stored in `UserProperties`.
*   `appsscript.json`: The project's manifest file. It defines the necessary OAuth scopes (permissions), time zone, and dependencies on advanced Google services required for the script to function.

## Manual Setup & Verification

For developers, security auditors, or users who wish to deploy this backend manually:

1.  **Create an Apps Script Project**: Navigate to [script.google.com](https://script.google.com) and create a new project.
2.  **Copy Code**: Copy the code from each `.gs` file in this repository into corresponding new script files in your Apps Script project. Ensure all filenames match exactly.
3.  **Enable Advanced Google Services**: In the Apps Script Editor, click on **Services (+)** in the left sidebar. Find and add the **Gmail API**. The Drive and Sheets APIs are typically available by default but ensure they are not disabled.
4.  **Configure the Script**: Open `Config.gs`. If you are forking the project and creating your own template, you must update the `TEMPLATE_SHEET_ID` with the ID of your own template spreadsheet. For initial testing, the `AdminUtils.gs` file contains functions for manually setting a Gemini API key via the script editor.
5.  **Deploy as a Web App**:
    *   Click the blue **Deploy** button > **New deployment**.
    *   Click the gear icon (⚙️) and select **Web app**.
    *   Configure the deployment with the following settings:
        *   **Execute as**: `Me`
        *   **Who has access**: `Anyone with Google account`
    *   Click **Deploy**.
    *   Google will prompt you to authorize the script's permissions. Review and grant them.
    *   Copy the generated **Web app URL**. This is the endpoint the companion Chrome Extension needs to communicate with.

## Privacy, Security, and Permission Justification

This script operates as a **zero-knowledge system** from the developer's perspective. We cannot see, access, or manage any of your data. The permissions you grant during the one-time authorization are required for the script to function *on your behalf* within your own account.

Here is the justification for the sensitive permissions required:

*   **Google Drive (`https://www.googleapis.com/auth/drive` scope)**: This permission is required for the script to create and manage the "CareerSuite.ai Data" spreadsheet in the user's Google Drive. It allows the script to create the sheet, and then open it by its unique ID to add new rows and update charts.

*   **Google Sheets (`https://www.googleapis.com/auth/spreadsheets` scope)**: This permission is necessary for the script to write data to and read data from the "CareerSuite.ai Data" spreadsheet. It is used to populate the "Applications" and "Potential Job Leads" sheets, as well as to update the dashboard with the latest metrics.

*   **Gmail (`https://www.googleapis.com/auth/gmail.readonly`, `https://www.googleapis.com/auth/gmail.labels`, and `https://www.googleapis.com/auth/gmail.settings.basic` scopes)**: These permissions are used to automate the job tracking process. The script uses these permissions to:
    1.  Create the necessary `CareerSuite.AI/...` labels in the user's Gmail account.
    2.  Create a filter to automatically route job-related emails to the appropriate labels.
    3.  Read emails that have been placed in the `.../To Process` label.
    4.  Modify an email's labels to move it from `.../To Process` to `.../Processed` after it has been analyzed.

For a complete overview of our data practices, please see our full [Privacy Policy](https://careersuiteai.vercel.app/privacy).
