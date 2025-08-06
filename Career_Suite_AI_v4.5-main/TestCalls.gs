// TestCalls.gs
// This file contains various test functions for different parts of the AI Resume Tailoring script,
// as well as UI-related functions like onOpen() for creating custom menus.
// It relies on global constants defined in 'Constants.gs' (e.g., MASTER_RESUME_SPREADSHEET_ID, MASTER_RESUME_DATA_SHEET_NAME).

// The MASTER_RESUME_SPREADSHEET_ID constant (formerly SCRIPT_TEST_SPREADSHEET_ID)
// is now expected to be defined in 'Constants.gs'.
// Example from Constants.gs:
// const MASTER_RESUME_SPREADSHEET_ID = "1_CQOG_FBovoYzudxZgsaVeZmn5XyAGJU82LgB9bFfd0"; // USER UPDATES IN Constants.gs

/**
 * LEGACY TEST: Tests the original, single-run full AI resume tailoring process.
 * NOTE: This function calls `orchestrateTailoringProcess` which may be commented out
 * or removed from Main.gs. Uncomment that function in Main.gs to use this test.
 */
function testFullResumeTailoring_Legacy() {
  Logger.log("--- Starting LEGACY Full Resume Tailoring Test (Using Groq) ---");
  Logger.log("NOTE: This test relies on the 'orchestrateTailoringProcess' function being active in Main.gs.");

  Logger.log("Step 1: Loading Master Resume from Sheet...");
  // Assumes MASTER_RESUME_SPREADSHEET_ID and MASTER_RESUME_DATA_SHEET_NAME are globally defined in Constants.gs
  const masterResumeObject = getMasterResumeData(MASTER_RESUME_SPREADSHEET_ID, MASTER_RESUME_DATA_SHEET_NAME);
  if (!masterResumeObject || !masterResumeObject.personalInfo) {
    Logger.log("FAILURE: Failed to load Master Resume Data. Aborting legacy test.");
    return;
  }
  Logger.log("SUCCESS: Master Resume loaded. Full Name: " + masterResumeObject.personalInfo.fullName);

  const sampleJobDescription = `
    **Job Title: Junior Data Analyst**
    **Company: Innovatech Solutions Inc.**
    **Location: New York, NY**
    **Responsibilities:**
    - Collect and interpret data from various sources.
    - Analyze results using statistical techniques and provide ongoing reports.
    - Develop and implement databases, data collection systems, data analytics, and other strategies that optimize statistical efficiency and quality.
    **Requirements:**
    - Proven working experience as a Data Analyst or similar role.
    - Strong knowledge of SQL and reporting packages (e.g., Tableau, Power BI).
    - Technical expertise regarding data models, database design development, data mining and segmentation techniques.
    - BS in Mathematics, Economics, Computer Science, Information Management or Statistics.
  `;
  Logger.log("Using Sample Job Description for 'Junior Data Analyst'.");

  Logger.log("Step 2: Calling orchestrateTailoringProcess()...");
  // Check if orchestrateTailoringProcess is defined before calling
  if (typeof orchestrateTailoringProcess === 'function') {
    const orchestrationResult = orchestrateTailoringProcess(masterResumeObject, sampleJobDescription);

    if (orchestrationResult && !orchestrationResult.error && orchestrationResult.tailoredResumeObject) {
      const tailoredResumeObject = orchestrationResult.tailoredResumeObject;
      const jdAnalysisForDocTitle = orchestrationResult.jdAnalysis;

      Logger.log("Step 3: orchestrateTailoringProcess completed successfully. Preparing to create document...");
      // Logger.log_DEBUG("Preview of tailored object (first 2000 chars): " + JSON.stringify(tailoredResumeObject, null, 2).substring(0, 2000) + "...");
      
      let jobTitleForDoc = "TargetJob";
      if (jdAnalysisForDocTitle && jdAnalysisForDocTitle.jobTitle && jdAnalysisForDocTitle.jobTitle.trim()) {
        jobTitleForDoc = jdAnalysisForDocTitle.jobTitle.trim().replace(/[^a-zA-Z0-9\s-]/g, "").substring(0, 30);
      } else if (tailoredResumeObject.summary) {
        const summaryWords = tailoredResumeObject.summary.split(/\s+/);
        jobTitleForDoc = summaryWords.slice(0, 3).join(" ").replace(/[^a-zA-Z0-9\s-]/g, "").substring(0, 30) || "Role";
      }
      const docTitle = `Tailored Resume (Legacy) - ${masterResumeObject.personalInfo.fullName} for ${jobTitleForDoc}`;

      Logger.log("Step 4: Calling createFormattedResumeDoc() with TAILORED (Legacy) Resume Data...");
      const docUrl = createFormattedResumeDoc(tailoredResumeObject, docTitle);

      if (docUrl) Logger.log(`SUCCESS: TAILORED (Legacy) resume doc created! URL: ${docUrl}`);
      else Logger.log("FAILURE: TAILORED (Legacy) resume doc creation failed.");
    } else {
      Logger.log("FAILURE: orchestrateTailoringProcess() returned an error or no tailored object.");
      if (orchestrationResult && orchestrationResult.error) {
        Logger.log("Orchestration Error: " + JSON.stringify(orchestrationResult.error, null, 2));
        if (orchestrationResult.details) Logger.log("Details: " + JSON.stringify(orchestrationResult.details, null, 2));
      }
    }
  } else {
    Logger.log("ERROR: orchestrateTailoringProcess function is not defined or accessible. Please ensure it's active in Main.gs if you intend to run this legacy test.");
  }
  Logger.log("--- LEGACY Full Resume Tailoring Test Finished ---");
}

/**
 * Tests ONLY the document formatting part using the master resume data.
 * This helps isolate issues with DocumentService.gs.
 */
function testMasterResumeFormattingOnly() {
  Logger.log("--- Starting MASTER RESUME FORMATTING ONLY Test ---");

  // Assumes MASTER_RESUME_SPREADSHEET_ID and MASTER_RESUME_DATA_SHEET_NAME are globally defined
  const masterResumeObject = getMasterResumeData(MASTER_RESUME_SPREADSHEET_ID, MASTER_RESUME_DATA_SHEET_NAME);
  if (!masterResumeObject || !masterResumeObject.personalInfo) {
    Logger.log("FAILURE: Could not load Master Resume data.");
    return;
  }
  Logger.log("SUCCESS: Master Resume loaded. Full Name: " + masterResumeObject.personalInfo.fullName);

  const docTitle = `Master Resume Format Test - ${masterResumeObject.personalInfo.fullName}`;
  const docUrl = createFormattedResumeDoc(masterResumeObject, docTitle);

  if (docUrl) {
    Logger.log(`SUCCESS: Master resume document created! URL: ${docUrl}`);
  } else {
    Logger.log("FAILURE: Master resume document creation failed.");
  }
  Logger.log("--- MASTER RESUME FORMATTING ONLY Test Finished ---");
}

/**
 * Tests core AI functionalities from TailoringLogic.gs (e.g., JD analysis, bullet matching/tailoring, summary generation).
 * Uses Groq as the backend.
 */
function testCoreFunctionality() {
  Logger.log("--- Starting Core AI Functionality Test (Groq Focus) ---");

  // Assumes MASTER_RESUME_SPREADSHEET_ID and MASTER_RESUME_DATA_SHEET_NAME are globally defined
  const masterResumeObject = getMasterResumeData(MASTER_RESUME_SPREADSHEET_ID, MASTER_RESUME_DATA_SHEET_NAME);
  if (!masterResumeObject || !masterResumeObject.personalInfo) {
    Logger.log("FAILURE: Could not load Master Resume for Core Functionality Test.");
    return;
  }
  Logger.log("Master Resume loaded for Core AI tests.");
  const candidateFullNameForSummary = masterResumeObject.personalInfo.fullName;
  const sampleExperienceBullet = (masterResumeObject.sections.find(s => s.title === "EXPERIENCE")
                                 ?.items[0]?.responsibilities[0]) || "Led a team to achieve project goals.";


  const sampleJobDescriptionShort = `
    **Job Title: Software Engineer**
    **Company: FutureTech Inc.**
    **Key Responsibilities:** Develop web applications, collaborate with team.
    **Required Skills:** JavaScript, React, Node.js.
  `;
  Logger.log("Using short sample JD for Core AI tests.");

  Logger.log("Testing analyzeJobDescription()...");
  const jdAnalysis = analyzeJobDescription(sampleJobDescriptionShort);
  if (!jdAnalysis || jdAnalysis.error) {
    Logger.log(`FAIL: analyzeJobDescription. Error: ${JSON.stringify(jdAnalysis)}`);
    return; // Stop test if JD analysis fails
  }
  Logger.log(`SUCCESS: JD Analysis (Groq):\n${JSON.stringify(jdAnalysis, null, 2).substring(0, 500)}...`);

  Logger.log("\nTesting matchResumeSection()...");
  const matchResult = matchResumeSection(sampleExperienceBullet, jdAnalysis);
  if (!matchResult || matchResult.error) {
    Logger.log(`FAIL: matchResumeSection. Error: ${JSON.stringify(matchResult)}`);
  } else {
    Logger.log(`SUCCESS: Match Result (Groq):\n${JSON.stringify(matchResult, null, 2)}`);
  }

  Logger.log("\nTesting tailorBulletPoint()...");
  const tailoredBullet = tailorBulletPoint(sampleExperienceBullet, jdAnalysis, jdAnalysis.jobTitle || "Target Role");
  if (!tailoredBullet || tailoredBullet.startsWith("ERROR:")) {
    Logger.log(`FAIL: tailorBulletPoint. Output: ${tailoredBullet}`);
  } else {
    Logger.log(`SUCCESS: Tailored Bullet (Groq):\n${tailoredBullet}`);
  }

  Logger.log("\nTesting generateTailoredSummary()...");
  const highlights = `${sampleExperienceBullet} ${tailoredBullet || ""}`; // Simple highlights
  const tailoredSummary = generateTailoredSummary(highlights, jdAnalysis, candidateFullNameForSummary);
  if (!tailoredSummary || tailoredSummary.startsWith("ERROR:")) {
    Logger.log(`FAIL: generateTailoredSummary. Output: ${tailoredSummary}`);
  } else {
    Logger.log(`SUCCESS: Tailored Summary (Groq):\n${tailoredSummary}`);
  }

  Logger.log("--- Core AI Functionality Test Finished ---");
}


// --- STAGED PROCESSING TESTS ---

/**
 * Tests Stage 1: JD Analysis & Master Resume Bullet Scoring.
 * Writes results to Google Sheets.
 */
function testStage1_AnalysisAndScoring_New() {
  Logger.log("--- Testing STAGE 1: JD Analysis & Scoring ---");

  // MASTER_RESUME_SPREADSHEET_ID is expected to be defined in Constants.gs
  if (!MASTER_RESUME_SPREADSHEET_ID || MASTER_RESUME_SPREADSHEET_ID.toUpperCase().includes("YOUR_ACTUAL")) {
    Logger.log("ERROR: MASTER_RESUME_SPREADSHEET_ID is not correctly set in Constants.gs.");
    SpreadsheetApp.getUi().alert("Configuration Error", "MASTER_RESUME_SPREADSHEET_ID not set in Constants.gs. Please inform the script developer.");
    return;
  }

  const sampleJD_For_Stage1_Test = `
    **Job Title: Entry Level Software Engineer**
    **Company: Tech Solutions LLC**
    **Responsibilities:** Assist senior developers, write clean code, test software.
    **Requirements:** BS in Computer Science, knowledge of Java or Python.
  `;
  Logger.log("Using SHORT Sample JD for Stage 1 Test.");

  // Calls runStage1_AnalyzeAndScore from Main.gs
  const result = runStage1_AnalyzeAndScore(sampleJD_For_Stage1_Test, MASTER_RESUME_SPREADSHEET_ID);

  Logger.log("\n--- Stage 1 Test Result ---");
  Logger.log(JSON.stringify(result, null, 2));
  if (result && result.success) {
    // JD_ANALYSIS_SHEET_NAME and BULLET_SCORING_RESULTS_SHEET_NAME are from Constants.gs
    Logger.log(`SUCCESS: Stage 1 completed. Check sheets "${JD_ANALYSIS_SHEET_NAME}" & "${BULLET_SCORING_RESULTS_SHEET_NAME}" in https://docs.google.com/spreadsheets/d/${MASTER_RESUME_SPREADSHEET_ID}/edit`);
  } else {
    Logger.log("FAILURE: Stage 1 did not complete successfully.");
  }
  Logger.log("--- Test for STAGE 1 Finished ---");
}

/**
 * Tests Stage 2: Tailors bullets selected (manually in the sheet) after Stage 1.
 * IMPORTANT PREREQUISITES:
 * 1. Run `testStage1_AnalysisAndScoring_New()` successfully.
 * 2. Manually open the 'BulletScoringResults' sheet.
 * 3. In the 'SelectToTailor(Manual)' column, enter "YES" or "TRUE" for a few bullets.
 */
function testStage2_TailoringSelected() {
  Logger.log("--- Testing STAGE 2: Tailoring Selected Bullets ---");

  // MASTER_RESUME_SPREADSHEET_ID and BULLET_SCORING_RESULTS_SHEET_NAME are expected from Constants.gs
  if (!MASTER_RESUME_SPREADSHEET_ID || MASTER_RESUME_SPREADSHEET_ID.toUpperCase().includes("YOUR_ACTUAL")) {
    Logger.log("ERROR: MASTER_RESUME_SPREADSHEET_ID is not correctly set in Constants.gs.");
    SpreadsheetApp.getUi().alert("Configuration Error", "MASTER_RESUME_SPREADSHEET_ID not set in Constants.gs.");
    return;
  }
  Logger.log(`Processing selections from spreadsheet ID: ${MASTER_RESUME_SPREADSHEET_ID}, expected sheet: ${BULLET_SCORING_RESULTS_SHEET_NAME}`);
  Logger.log("IMPORTANT: Ensure you've run Stage 1 and manually selected bullets in the sheet before running this test.");

  // Calls runStage2_TailorSelectedBullets from Main.gs
  const result = runStage2_TailorSelectedBullets(MASTER_RESUME_SPREADSHEET_ID);

  Logger.log("\n--- Stage 2 Test Result ---");
  Logger.log(JSON.stringify(result, null, 2));
  if (result && result.success) {
    Logger.log(`SUCCESS: Stage 2 completed. Check the "TailoredBulletText(Stage2)" column in sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" in your spreadsheet: https://docs.google.com/spreadsheets/d/${MASTER_RESUME_SPREADSHEET_ID}/edit`);
  } else {
    Logger.log("FAILURE: Stage 2 did not complete successfully.");
  }
  Logger.log("--- Test for STAGE 2 Finished ---");
}

/**
 * Tests Stage 3: Assembles the resume object and generates the Google Document.
 * IMPORTANT PREREQUISITES:
 * 1. Run `testStage1_AnalysisAndScoring_New()` successfully.
 * 2. Run `testStage2_TailoringSelected()` successfully (after manually selecting bullets).
 */
function testStage3_GenerateDocument() {
  Logger.log("--- Testing STAGE 3: Assemble & Generate Document ---");

  // MASTER_RESUME_SPREADSHEET_ID, JD_ANALYSIS_SHEET_NAME, BULLET_SCORING_RESULTS_SHEET_NAME from Constants.gs
  if (!MASTER_RESUME_SPREADSHEET_ID || MASTER_RESUME_SPREADSHEET_ID.toUpperCase().includes("YOUR_ACTUAL")) {
    Logger.log("ERROR: MASTER_RESUME_SPREADSHEET_ID is not correctly set in Constants.gs.");
    SpreadsheetApp.getUi().alert("Configuration Error", "MASTER_RESUME_SPREADSHEET_ID not set in Constants.gs.");
    return;
  }

  // Check prerequisite sheets
  try {
    const ss = SpreadsheetApp.openById(MASTER_RESUME_SPREADSHEET_ID);
    const jdSheet = ss.getSheetByName(JD_ANALYSIS_SHEET_NAME);
    const scoreSheet = ss.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME);

    if (!jdSheet || !jdSheet.getRange(2, 1).getValue()) {
      const msg = `Sheet "${JD_ANALYSIS_SHEET_NAME}" is missing or empty. Please run Stage 1 first.`;
      Logger.log(`ERROR (Prerequisite Check): ${msg}`);
      SpreadsheetApp.getUi().alert("Prerequisite Missing", msg);
      return;
    }
    if (!scoreSheet || scoreSheet.getLastRow() < 2) {
      const msg = `Sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" is missing or has no data. Run Stage 1 & 2 (with manual selections) first.`;
      Logger.log(`ERROR (Prerequisite Check): ${msg}`);
      SpreadsheetApp.getUi().alert("Prerequisite Missing", msg);
      return;
    }
    Logger.log(`Prerequisite sheets "${JD_ANALYSIS_SHEET_NAME}" and "${BULLET_SCORING_RESULTS_SHEET_NAME}" found with data.`);
  } catch (e) {
    Logger.log(`ERROR (Prerequisite Check): Could not verify sheets. ${e.toString()}`);
    return;
  }

  Logger.log(`Processing using Spreadsheet ID: ${MASTER_RESUME_SPREADSHEET_ID}`);
  
  // Calls runStage3_BuildAndGenerateDocument from Main.gs
  const result = runStage3_BuildAndGenerateDocument(MASTER_RESUME_SPREADSHEET_ID);
  
  Logger.log("\n--- Stage 3 Test Result ---");
  Logger.log(JSON.stringify(result, null, 2)); // Full result object
  if (result && result.success && result.docUrl) {
    Logger.log(`SUCCESS: Stage 3 completed! Document URL: ${result.docUrl}`);
  } else {
    Logger.log("FAILURE: Stage 3 did not complete successfully or document URL was not returned.");
  }
  Logger.log("--- Test for STAGE 3 Finished ---");
}


// --- UI RELATED FUNCTIONS ---
// These functions are typically triggered by user interaction with the Google Workspace UI (e.g., a custom menu).

/**
 * Creates a custom menu in the Google Spreadsheet UI when the spreadsheet is opened.
 * This should be defined only once in your project (e.g., in Code.gs or a dedicated UI.gs).
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Resume Tools')
    .addItem('SETUP: Create/Format Resume Data Sheet', 'setupMasterResumeSheet') // Assumes setupMasterResumeSheet is defined
    .addSeparator()
    .addItem('STEP 1: Analyze JD & Score All Bullets', 'triggerStage1FromUi')
    // Future UI items for Stage 2 (Review & Select) and Stage 3 (Generate Doc) would go here.
    // .addItem('STEP 2: Tailor Selected & Generate Resume', 'triggerStage2And3FromUi_Combined') 
    .addSeparator()
    .addSubMenu(ui.createMenu('Run Tests')
        .addItem('Test: Stage 1 - Analyze & Score', 'testStage1_AnalysisAndScoring_New')
        .addItem('Test: Stage 2 - Tailor Selected', 'testStage2_TailoringSelected')
        .addItem('Test: Stage 3 - Generate Document', 'testStage3_GenerateDocument')
        .addSeparator()
        .addItem('Test: Format Master Resume Only', 'testMasterResumeFormattingOnly')
        .addItem('Test: Core AI Logic Blocks', 'testCoreFunctionality')
        .addItem('Test: Legacy Full Tailoring (if active)', 'testFullResumeTailoring_Legacy') 
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Configuration')
        .addItem('Set Groq API Key', 'SET_GROQ_API_KEY_UI')       // Assumes function in GroqService.gs
        // .addItem('Set Gemini API Key', 'SET_GEMINI_API_KEY_UI') // Assumes function in GeminiService.gs
    )
    .addToUi();
}

/**
 * UI-triggered function to run Stage 1: JD Analysis and Master Resume Scoring.
 * Prompts the user for Job Description text.
 */
function triggerStage1FromUi() {
  const ui = SpreadsheetApp.getUi();
  
  // Check for MASTER_RESUME_SPREADSHEET_ID from Constants.gs
  if (typeof MASTER_RESUME_SPREADSHEET_ID === 'undefined' || !MASTER_RESUME_SPREADSHEET_ID || MASTER_RESUME_SPREADSHEET_ID.toUpperCase().includes("YOUR_ACTUAL")) {
    ui.alert("Configuration Error", "MASTER_RESUME_SPREADSHEET_ID is not correctly set in Constants.gs. Please inform the script developer.");
    return;
  }

  const jdResponse = ui.prompt(
    "Job Description for Stage 1",
    "Paste the full Job Description text here:",
    ui.ButtonSet.OK_CANCEL
  );

  if (jdResponse.getSelectedButton() === ui.Button.OK) {
    const jdText = jdResponse.getResponseText();
    if (jdText && jdText.trim() !== "") {
      ui.alert(
        "Processing Stage 1",
        "Analyzing Job Description and scoring master resume bullets. This may take several minutes. Please check Execution Logs (View > Logs) for progress. You will receive a completion alert.",
        ui.ButtonSet.OK
      );
      // Calls runStage1_AnalyzeAndScore from Main.gs
      const result = runStage1_AnalyzeAndScore(jdText, MASTER_RESUME_SPREADSHEET_ID);
      
      if (result.success) {
        // JD_ANALYSIS_SHEET_NAME and BULLET_SCORING_RESULTS_SHEET_NAME are from Constants.gs
        ui.alert("Stage 1 Complete", `${result.message} Review sheets: "${JD_ANALYSIS_SHEET_NAME}" and "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);
      } else {
        ui.alert("Stage 1 Failed", `${result.message}${result.details ? ` Details: ${JSON.stringify(result.details)}` : ""}`);
      }
    } else {
      ui.alert("Input Error", "No Job Description text was provided.");
    }
  } else {
    ui.alert("Cancelled", "Stage 1 process was cancelled.");
  }
}
