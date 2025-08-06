# CareerSuite.ai - Google Apps Script Backend

This repository contains the complete open-source Google Apps Script code that serves as the backend for the [CareerSuite.ai Chrome Extension](https://[LINK_TO_YOUR_EXTENSION_STORE_PAGE_ONCE_LIVE]). This script runs entirely within a user's own Google Account, ensuring complete data privacy and user control.

## Overview & Architecture

The primary purpose of this backend is to automate the tracking of job applications by intelligently processing emails and managing a dedicated Google Sheet. This script is deployed as a Google Web App and is called securely by the Chrome Extension.

**Our Core Principles:**
1.  **Data Sovereignty:** The user owns their data. Nothing is stored on our servers. All data (the user's resume, application emails, the tracker sheet) resides exclusively within the user's personal Google Account.
2.  **No-Cost Utility:** The system is powered by the user's personal Google Gemini API key. This allows us to provide powerful AI features for free, subject only to Google's own API usage limits.
3.  **Transparency:** This entire codebase is open-source. Users and reviewers can audit every line of code to verify our data handling practices.

## Key Features Powered by This Script

-   **Automated Sheet Creation:** On first use, the script creates a copy of a master "CareerSuite.ai Data" Google Sheet in the user's Google Drive. The extension is sandboxed to only have access to this one file it creates.
-   **Automated Email Processing:**
    -   The script creates a filter and a set of labels (e.g., `CareerSuite.AI/To Process`) in the user's Gmail.
    -   It periodically scans *only* the emails within this specific label.
    -   It uses the Gemini API to analyze these emails, extracting key information like Company Name, Job Title, and Application Status (e.g., "Interview Scheduled", "Rejected").
    -   It then updates the Google Sheet with this information, providing a real-time view of the user's application pipeline.
-   **Dashboard & Analytics:** The script populates a "Dashboard" tab within the Google Sheet, automatically generating charts and key metrics like application counts, interview rates, and weekly trends.

---

## File Structure

This repository contains all the `.gs` files for the Google Apps Script project.

-   **`Main.gs`**: The primary orchestration file. Contains the `onOpen` function to create the spreadsheet menu and the main setup function `runFullProjectInitialSetup`.
-   **`WebApp_Endpoints.gs`**: Handles all incoming web requests (`doGet` and `doPost`) from the Chrome Extension. This is the main entry point for the client.
-   **`GeminiService.gs`**: Contains the logic for making API calls to the Google Gemini service for parsing email content.
-   **`GmailUtils.gs`**: Helper functions for creating and managing Gmail labels.
-   **`SheetUtils.gs`**: General helper functions for creating, formatting, and interacting with the Google Sheet.
-   **`Dashboard.gs`**: Contains the logic specifically for creating and updating the charts and metrics on the Dashboard sheet.
-   **`ParsingUtils.gs`**: Contains fallback regex and keyword-based parsing logic for when the Gemini AI is unavailable or fails.
-   **`AdminUtils.gs`**: Utility functions for the project owner, such as managing API keys (for development).
-   **`Leads_Main.gs` & `Leads_SheetUtils.gs`**: The files that handle the new "Potential Job Leads" tracking feature.
-   **`Triggers.gs`**: Manages the time-driven triggers that run the email processing functions automatically.
-   **`Config.gs`**: A central file containing all configuration constants for the project, such as sheet names, label names, and API properties.

---

## Setup for Reviewers & Contributors

This script is designed to be set up automatically by the CareerSuite.ai Chrome Extension. However, for manual setup or review:

1.  **Create a new Google Apps Script project.**
2.  **Copy the code** from each `.gs` file in this repository into corresponding files in your Apps Script project. Ensure filenames match.
3.  **Enable Advanced Google Services:**
    -   In the Apps Script Editor, click on "Services" in the left sidebar.
    -   Find and enable both the **Gmail API** and the **Google Drive API**.
4.  **Set Configuration:**
    -   In `Config.gs`, update the `TEMPLATE_SHEET_ID` with the ID of your own template sheet if you are forking the project.
    -   The Gemini API key is intended to be set via the extension's UI. For testing, you can use the functions in `AdminUtils.gs`.
5.  **Deploy as a Web App:**
    -   Click **Deploy > New deployment**.
    -   Configure the Web App with **Execute as: "Me"** and **Who has access: "Anyone with Google account"**.
    -   Authorize the necessary permissions when prompted.
    -   The resulting Web App URL is the endpoint that the companion Chrome Extension communicates with.

## Privacy & Security

As stated in our official [Privacy Policy](https://[https://careersuiteai.vercel.app/privacy]), this script adheres to a strict user-first privacy model. **The security of the user's data is maintained because the script and the data it processes never leave the user's own Google Account.** Our role as developers is simply to provide the open-source code that enables this powerful, private automation.
