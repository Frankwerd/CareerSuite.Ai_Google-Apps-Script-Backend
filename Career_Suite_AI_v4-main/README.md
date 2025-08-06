# Career_Suite_AI_v4
# 🤖 Master Job Manager: An AI-Powered Job Application CRM in Google Sheets **REVISED NAME**

**The Master Job Manager transforms your Gmail and Google Sheets into an intelligent, automated system that manages the entire job application lifecycle, from lead generation to final status.**
---

## 🚀 The Problem

The modern job search is a high-volume activity. It's incredibly easy to lose track of which applications have been viewed, which jobs you've sourced from email alerts, and which applications have gone stale. Manually updating a spreadsheet is tedious, time-consuming, and prone to human error, taking valuable time away from preparing for interviews.

## ✨ The Solution

This project automates the entire process, acting as a personal CRM for your job search that lives right where you work: your Google Workspace.

The **Master Job Manager** actively monitors your Gmail for new application submissions, status updates (like "your application was viewed"), and job alert emails. It uses **Google's Gemini AI** to intelligently parse these emails, then populates a centralized, dynamic dashboard in Google Sheets, giving you a bird's-eye view of your entire pipeline.

### Key Features

*   **🧠 AI-Powered Email Parsing:** Leverages the Google Gemini API to accurately extract Company Name, Job Title, and Application Status from complex emails. Includes a robust, multi-layered regex parser as a reliable fallback.

*   **🗂️ Dual-Module System:**
    *   **Application Tracker:** Processes application confirmations and status updates from submission to offer or rejection.
    *   **Job Leads Tracker:** A separate module that uses AI to parse job alert emails (e.g., from LinkedIn, Indeed) and extracts multiple job leads into a separate database for review.

*   **📊 Automated BI Dashboard:** The dashboard provides at-a-glance KPIs and visualizations, including:
    *   Key metrics like Total Applications, Active Applications, and Interview/Offer Rates.
    *   A dynamic application funnel showing progression from `Applied` to `Offer`.
    *   Charts for applications over time and distribution by job platform.

*   **⚙️ Full-Cycle Automation:**
    *   **One-Click Setup:** A custom menu item runs a comprehensive setup function that creates all necessary Gmail labels, Gmail filters, Sheet tabs, and time-driven triggers.
    *   **Automated Processing:** Hourly and daily triggers automatically process new emails without any user intervention.
    *   **Data Maintenance:** A "stale application" handler automatically marks applications as rejected after a set period of inactivity to keep the pipeline clean.

*   **💻 Professional Architecture:** The code is organized into a clean, modular structure with a separation of concerns (e.g., `GeminiService.gs`, `Dashboard.gs`, `SheetUtils.gs`), demonstrating best practices for maintainability and scalability.

## 🛠️ Tech Stack

*   **Core Logic:** Google Apps Script (JavaScript ES5)
*   **APIs & Services:**
    *   **Google Gemini API:** For AI-driven data extraction.
    *   **Advanced Gmail Service:** For programmatic creation and management of labels and filters.
    *   **Standard Services:** `GmailApp`, `SpreadsheetApp`, `UrlFetchApp`, `PropertiesService`, `ScriptApp`.
*   **Development & Tooling:**
    *   **`clasp`:** Google's command-line tool for local development, version control with Git, and deployments.

## 💡 How to Use

1.  **Make Your Own Copy of the Sheet**
    *   [**➡️ Click here to create your own copy of the Master Job Manager Sheet**](https://your-google-sheet-link-here-ending-in/copy)
    **(Action Item: Create a sanitized "Public Demo" sheet as we discussed and put the `/copy` link here.)**

2.  **Get a Gemini API Key**
    *   The script uses Google's Gemini API for its AI features. You can get a free API key from Google AI Studio.
    *   Follow the instructions here: [**Get a Gemini API Key**](https://aistudio.google.com/app/apikey).

3.  **Configure the Script**
    *   In your new spreadsheet, go to the top menu and click `⚙️ Master Job Manager` > `Admin & Config` > `Set Shared Gemini API Key`.
    *   Paste your API key when prompted.

4.  **Run the Full Setup**
    *   From the same menu, click `⚙️ Master Job Manager` > `▶️ RUN FULL PROJECT SETUP`.
    *   You will be asked to authorize the script's permissions. This is a one-time step that allows the script to manage your sheets and emails.

5.  **Done!**
    *   The system is now live! It will automatically process new job-related emails as they arrive.
