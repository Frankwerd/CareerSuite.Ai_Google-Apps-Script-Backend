# Career Suite AI v2.5 **REVISED PROJECT NAME**
# AI-Powered Gmail Job Lead Processor (Google Apps Script & Gemini API)

Tired of sifting through endless "job alert" emails? This Google Apps Script automates the processing of job notifications **you already receive** from your favorite job boards. It uses Google's Gemini AI to intelligently scan these emails, extract key information (Job Title, Company, Location, Link), and organizes it all neatly into a Google Spreadsheet.

**How This Tool Respects Terms of Service & User Control:**

*   **User-Initiated Job Alerts:** This script **DOES NOT** scrape job boards or websites. You, the user, are responsible for setting up job alert notifications from your preferred job boards (e.g., LinkedIn, Indeed, company career pages) to be delivered to your Gmail.
*   **Processing Existing Emails:** The script only processes emails that **land in your Gmail inbox** as a result of the alerts you've configured.
*   **Gmail & Google API Compliance:** The script utilizes official Google APIs (Gmail, Google Sheets, Gemini) within Google's ecosystem, adhering to their standard terms of service for API usage.
*   **Manual Tracking Supplement:** While it automates initial lead capture and organization, the detailed tracking of applications, interviews, etc., within the generated spreadsheet is still a manual process managed by you.

**Features:**

*   **Processes Your Existing Job Alerts:** Works with job notification emails you set up from various job boards.
*   **Automatic Gmail Filtering & Labeling:** Identifies these job alert emails, applies custom labels (e.g., "Job Application Potential/NeedsProcess"), and archives them from your inbox.
*   **AI-Powered Data Extraction:** Leverages the Google Gemini API to parse email content and extract job details.
*   **Google Sheet Integration:** Creates a new, organized Google Sheet to track all your potential job leads.
*   **Automated Daily Processing:** Sets up a daily trigger to process new emails automatically.
*   **Easy Setup:** Runs directly within your Google account – no complex local software installations required.
*   **Open Source:** Free to use, modify, and contribute to!

## Screenshots (Optional but Highly Recommended)

*(Consider adding a couple of screenshots here showing:*
*   *The Gmail labels created (e.g., "Job Application Potential/NeedsProcess").*
*   *An example of the populated Google Sheet with job leads.)*

---

## Table of Contents

*   [Prerequisites](#prerequisites)
*   [Setup Instructions](#setup-instructions)
    *   [Part 0: Setting Up Job Board Alerts (Crucial First Step!)](#part-0-setting-up-job-board-alerts-crucial-first-step)
    *   [Part 1: Getting the Script Code into Google Apps Script](#part-1-getting-the-script-code-into-google-apps-script)
    *   [Part 2: Getting Your Gemini API Key](#part-2-getting-your-gemini-api-key)
    *   [Part 3: Running the Initial Setup in Apps Script](#part-3-running-the-initial-setup-in-apps-script)
    *   [Part 4: Understanding How It Works](#part-4-understanding-how-it-works)
*   [Usage](#usage)
*   [Troubleshooting](#troubleshooting)
*   [Contributing](#contributing)
*   [License](#license)
*   [Disclaimer](#disclaimer)

---

## Prerequisites

*   A Google Account (Gmail, Google Drive).
*   **You must have already set up job alert notifications from your desired job boards (e.g., LinkedIn, Indeed, company career pages) to be sent to your Gmail address.** This script *processes* these incoming alerts; it does not generate them.
*   Basic familiarity with navigating Google services.

---

## Setup Instructions

**Part 0: Setting Up Job Board Alerts (Crucial First Step!)**

Before you set up this script, ensure you have configured job alerts from your preferred job boards to send email notifications to your Gmail account. The script relies on these emails (typically with subjects like "job alert," "new job matches," etc.) as its input.

**Examples:**
*   Go to LinkedIn Jobs and set up alerts for specific roles/locations.
*   Go to Indeed and create email alerts for your searches.
*   Subscribe to notifications from specific company career pages.

**The script's filter is looking for subjects containing "job alert" or "jobalert".** If your job board emails use a different common subject line, you may need to adjust the `filterQuery` variable within the `runInitialSetup_createNewSpreadsheetWithJobLabels` function in the script code (advanced users only).

**Part 1: Getting the Script Code into Google Apps Script**

1.  **Open the Script File:**
    *   Navigate to the `[YOUR_SCRIPT_FILENAME_HERE.gs]` file in this repository (e.g., `JobProcessor.gs` or `Code.gs`).
    *   Click the "Raw" button to view the plain text code.
    *   Select all the code (`Ctrl+A` or `Cmd+A`) and copy it (`Ctrl+C` or `Cmd+C`).

2.  **Open Google Apps Script:**
    *   Go to [script.google.com](https://script.google.com).
    *   Click on **"+ New project"**.
    *   Delete any placeholder code (like `function myFunction() {}`) in the editor.
    *   Paste the code you copied in Step 1 into the editor.
    *   Give your project a name at the top left (e.g., "Gemini Gmail Job Processor"). Click "Rename".
    *   Click the **Save icon** (floppy disk) or `Ctrl+S` (Windows) / `Cmd+S` (Mac).

3.  **Enable the Gmail Advanced Service:**
    *   In the Apps Script editor, look at the left-hand sidebar.
    *   Next to "Services," click the **"+"** icon.
    *   Scroll down and find "**Gmail API**". Select it.
    *   The "Identifier" should remain `Gmail`. Click "**Add**".
    *   *(This allows the script to create filters and manage labels more effectively).*

**Part 2: Getting Your Gemini API Key**

This script uses Google's Gemini AI to understand your emails. You'll need an API key.

1.  **Get a Gemini API Key:**
    *   Go to [Google AI Studio](https://aistudio.google.com/app/apikey).
    *   You might need to sign in with your Google account.
    *   Click "**Create API key in new project**" (or select an existing project if you have one).
    *   **Copy the generated API key** carefully. Keep it safe and private.
    *   *(Note: Use of the Gemini API is subject to Google's terms and may involve costs depending on usage. Check Google's pricing page for the Gemini API for details. Free tiers are often available for initial/low usage.)*

**Part 3: Running the Initial Setup in Apps Script**

1.  **Set Your Gemini API Key in the Script:**
    *   Back in the Apps Script editor:
    *   Near the top of the editor, find the dropdown menu that currently says "Select function". Click it.
    *   Choose the function `setGeminiApiKey_UI`.
    *   Click the **"Run"** button (it looks like a play ► icon).
    *   **Authorization:** The first time you run any function that needs permissions, you'll see an "Authorization required" dialog.
        *   Click **"Review permissions"**.
        *   Choose your Google account.
        *   If you see a "Google hasn’t verified this app" screen, click **"Advanced"** and then **"Go to [Your Project Name] (unsafe)"**. (This project name will be what you named your Apps Script project, e.g., "Gemini Gmail Job Processor").
        *   Finally, click **"Allow"**. This is standard for personal Apps Scripts you create.
    *   A dialog box will pop up: "Gemini API Key".
    *   Paste the Gemini API Key you copied in Part 2, Step 1 into the text field.
    *   Click **"OK"**. You should see a confirmation "API Key Saved."

2.  **Run the Full Setup Process:**
    *   In the "Select function" dropdown, now choose `runInitialSetup_createNewSpreadsheetWithJobLabels`.
    *   Click the **"Run"** button.
    *   A confirmation dialog "Confirm Full Setup" will appear, detailing what the script will do. Read it and click **"YES"** if you wish to proceed.
    *   **Permissions:** You'll likely be asked for permissions again for Gmail and Google Drive. Follow the same authorization steps as above (Review permissions, select account, Advanced, Go to... (unsafe), Allow).
    *   The script will now:
        *   Create a new Google Spreadsheet named "Job Leads Tracker - [Date]".
        *   Create Gmail labels ("Job Application Potential", ".../NeedsProcess", ".../DoneProcess").
        *   Set up a Gmail filter for emails with "job alert" (or "jobalert") in the subject, labeling them and archiving them from the inbox.
        *   Schedule itself to run automatically once a day (around 3 AM in the script's timezone, check script logs for exact timing if needed).
    *   You'll see a "Setup Complete!" message with a link to your new spreadsheet. **Copy this spreadsheet URL and bookmark it!**

**Part 4: Understanding How It Works**

*   **Job Board Alerts First:** You set up job alerts from your preferred job boards to send emails to your Gmail.
*   **Automatic Filtering:** The script's filter (created during setup) watches for incoming emails with subjects like "job alert." It applies the "Job Application Potential/NeedsProcess" label and archives the email (removes it from the inbox).
*   **Daily Processing:** Once a day, the `processJobLeadsWithGemini` function (the one triggered automatically):
    1.  Looks for emails in the "Job Application Potential/NeedsProcess" label.
    2.  Sends the content of these emails to the Gemini AI to extract job title, company, location, and a job link.
    3.  Writes this information into your "Job Leads Tracker" Google Sheet.
    4.  Moves successfully processed email threads to the "Job Application Potential/DoneProcess" label.
*   **Sheet Management:** The spreadsheet is your central hub. You'll manually update statuses, add notes, and manage your application process there.

---

## Usage

*   **Check Your Spreadsheet:** Regularly open your "Job Leads Tracker" Google Sheet to see newly processed job leads.
*   **Manage Gmail Labels:** Emails processed will be in the "Job Application Potential/DoneProcess" label. Emails waiting for processing are in ".../NeedsProcess".
*   **Manual Processing (Optional):**
    *   If you want to process emails immediately without waiting for the daily trigger:
        1.  Go back to the Apps Script editor ([script.google.com](https://script.google.com) and open your project).
        2.  Select the `processJobLeadsWithGemini` function from the "Select function" dropdown.
        3.  Click **"Run"**. (You may need to authorize again the very first time you run it manually this way).
*   **Updating API Key:** If you need to change your Gemini API Key, run the `setGeminiApiKey_UI` function again from the Apps Script editor.

---

## Troubleshooting

*   **Script Not Running / No New Leads:**
    *   **Check Job Board Alerts:** Are you actually receiving job alert emails from job boards into your Gmail?
    *   **Check Email Subjects:** Do your job alert emails have "job alert" or "jobalert" in the subject? If not, the filter won't catch them. You might need to adjust the filter criteria in the script.
    *   **Check Apps Script Executions:** In the Apps Script editor, go to "Executions" (clock icon on the left). Look for recent runs of `processJobLeadsWithGemini`.
        *   If "Status" is "Failed," click on the run to see the error message. This can give clues.
        *   If there are no recent runs, ensure the trigger was set up correctly (the "Setup Complete!" dialog during `runInitialSetup...` confirms this). You can also check "Triggers" (alarm clock icon) in Apps Script.
    *   **Check Gmail Labels:** Are emails piling up in the ".../NeedsProcess" label but not moving to ".../DoneProcess" or the sheet? This could indicate an issue with Gemini processing or writing to the sheet. Check the Executions logs.
    *   **Gemini API Key:** Is your API key correct and active? Does it have sufficient quota?
    *   **Permissions:** Ensure all required permissions were granted during setup.
*   **"Authorization required" Looping:** Sometimes, especially if you have multiple Google accounts, authorization can get tricky. Try signing out of all Google accounts, then sign back into ONLY the one you want to use for this script and retry the setup/run.
*   **"Google hasn't verified this app":** This is normal for personal Apps Script projects that haven't gone through Google's official verification process (which is complex and usually for widely distributed add-ons). Clicking "Advanced" and "Go to [Project Name] (unsafe)" is the standard way to proceed for your own scripts.
*   **Spreadsheet Issues:**
    *   Can't find the spreadsheet? Check your Google Drive for a file named "Job Leads Tracker - [Date]". The setup log also provides the URL.
    *   Headers messed up? Avoid manually changing the header row names in the sheet if possible, as the script uses them to find where to write data.

---

## Contributing

Contributions are welcome! If you have ideas for improvements, bug fixes, or new features:

1.  **Fork the repository.**
2.  **Create a new branch** (`git checkout -b feature/your-amazing-feature`).
3.  **Make your changes.**
4.  **Commit your changes** (`git commit -m 'Add some amazing feature'`).
5.  **Push to the branch** (`git push origin feature/your-amazing-feature`).
6.  **Open a Pull Request.**

Please ensure your code is well-commented and, if you're adding new functionality, consider how it impacts the existing setup and user experience.

---

## License

This project is licensed under the **MIT License**. See the `LICENSE` file for details.

---

## Disclaimer

*   This script is provided "as-is" without warranty of any kind.
*   You are responsible for setting up job alert emails from third-party job boards. This script processes emails you receive; it does not interact directly with job board websites or APIs beyond processing the emails they send you.
*   Ensure your use of this script and the Gemini API complies with all applicable Google Terms of Service and the terms of service of any job boards whose email alerts you are processing.
*   The accuracy of data extraction depends on the Gemini API and the format of incoming emails. Results may vary.
*   Always safeguard your Gemini API Key.
*   The developer of this script is not responsible for any misuse, data loss, or issues arising from its use.
