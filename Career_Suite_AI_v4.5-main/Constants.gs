// Constants.gs
// This file centralizes global constants and configuration settings
// for the AI-Powered Resume Tailoring Script project.

// --- SPREADSHEET & DATA CONFIGURATION ---

/** 
 * @const {string} The ID of the Google Spreadsheet containing the master resume data 
 *  and where intermediate processing sheets (JD Analysis, Bullet Scoring) will be written. 
 *  Also used by SheetSetup.gs as the default ID to format or the ID to update if a new sheet is made.
 *  IMPORTANT: USER MUST UPDATE THIS VALUE if not allowing dynamic creation.
 */
const MASTER_RESUME_SPREADSHEET_ID = ""; // << REPLACE WITH YOUR ACTUAL SPREADSHEET ID or let SheetSetup create one

/** 
 * @const {string} The name of the sheet tab within MASTER_RESUME_SPREADSHEET_ID 
 *  that holds the structured master resume content. Also used by SheetSetup.gs.
 */
const MASTER_RESUME_DATA_SHEET_NAME = "ResumeData";

/** 
 * @const {string} The name of the sheet tab used to store the JSON output of the 
 *  Job Description analysis (Stage 1 output).
 */
const JD_ANALYSIS_SHEET_NAME = "JDAnalysisData";

/** 
 * @const {string} The name of the sheet tab used to store the results of master resume
 *  bullet scoring against the JD (Stage 1 output, Stage 2 input/output).
 */
const BULLET_SCORING_RESULTS_SHEET_NAME = "BulletScoringResults";


// --- DOCUMENT TEMPLATE ---

/** 
 * @const {string} The ID of the Google Document to be used as the template for 
 *  generating the final formatted resume.
 */
const RESUME_TEMPLATE_DOC_ID = "18eX765FWVBHpOZ2jzxwNdGVKSMOzmyPJmS_Kfzp8JSk";


// --- API & LLM CONFIGURATION ---

/** @const {string} The default Groq model name to be used for AI tasks. */
const GROQ_MODEL_NAME = "gemma2-9b-it"; 
// Note: The line `const GROQ_MODEL_NAME = "gemma2-9b-it";` was in your TailoringLogic.gs.
// For consistency, I'm using GROQ_DEFAULT_MODEL_NAME here as the standard. Ensure TailoringLogic.gs
// and GroqService.gs default parameter now uses this constant.

/** @const {string} The default Gemini model name to be used if GeminiService.gs is called. */
const GEMINI_DEFAULT_MODEL = "gemini-1.5-flash-latest"; // Default for GeminiService.gs

// Delays for Staged Processing (in milliseconds)
/** @const {number} Delay between LLM calls during Stage 1 (JD Analysis & Bullet Scoring). */
const API_CALL_DELAY_STAGE1_MS = 2000;
/** @const {number} Delay between LLM calls during Stage 2 (Tailoring Selected Bullets). */
const API_CALL_DELAY_STAGE2_MS = 2000;
/** @const {number} Delay before the LLM call for generating the tailored summary in Stage 3. */
const API_CALL_DELAY_STAGE3_SUMMARY_MS = 2000;


// --- STAGE 3: RESUME ASSEMBLY & DOCUMENT GENERATION CONFIGURATION ---

/** @const {number} Minimum relevance score for a bullet to be included in Stage 3. */
const STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD = 0.01;
/** @const {number} Max bullets per job in the Experience section during Stage 3. */
const STAGE3_MAX_BULLETS_PER_JOB = 5;
/** @const {number} Max bullets per project in the Projects section during Stage 3. */
const STAGE3_MAX_BULLETS_PER_PROJECT = 5;
/** @const {number} Max highlights for AI Summary input in Stage 3. */
const STAGE3_MAX_HIGHLIGHTS_FOR_SUMMARY = 4;


// --- RESUME STRUCTURE & SHEET SETUP CONFIGURATION ---

/**
 * @const {number} Defines how many dedicated columns (e.g., Responsibility1, Responsibility2)
 * are created in the "ResumeData" sheet for itemized bullet points and parsed.
 */
const NUM_DEDICATED_BULLET_COLUMNS = 3;

/** 
 * @const {Array<Object>} Defines the structure for sections in the master resume sheet.
 * Each object contains a 'title' and a 'headers' array (or null for non-tabular).
 * Headers listed here are base headers; numbered bullet columns like "Responsibility1"
 * will be dynamically added by SheetSetup.gs and parsed by MasterResumeData.gs based on NUM_DEDICATED_BULLET_COLUMNS.
 */
const RESUME_STRUCTURE = [
  { title: "PERSONAL INFO", headers: ["Key", "Value"] }, // SheetSetup will populate keys like "Full Name", "Email"
  { title: "SUMMARY", headers: null }, // No tabular headers; content usually in one cell or merged
  { 
    title: "EXPERIENCE", 
    headers: ["JobTitle", "Company", "Location", "StartDate", "EndDate"] // Responsibility1,2,3 added dynamically
  },
  { 
    title: "EDUCATION", 
    headers: ["Institution", "Degree", "Location", "StartDate", "EndDate", "GPA", "RelevantCoursework"] 
  },
  { 
    title: "TECHNICAL SKILLS & CERTIFICATES", 
    headers: ["CategoryName", "SkillItem", "Details", "Issuer", "IssueDate"] 
  },
  { 
    title: "PROJECTS", 
    headers: ["ProjectName", "Organization", "Role", "StartDate", "EndDate", "Technologies", "GitHubName1", "GitHubURL1", "Impact", "FutureDevelopment"] // DescriptionBullet1,2,3 added dynamically
  },
  { 
    title: "LEADERSHIP & UNIVERSITY INVOLVEMENT", 
    headers: ["Organization", "Role", "Location", "StartDate", "EndDate", "Description"] // Responsibility1,2,3 added dynamically
  },
  { 
    title: "HONORS & AWARDS", 
    headers: ["AwardName", "Details", "Date"] 
  }
];

// Styling Constants for SheetSetup.gs
/** @const {string} Hex color code for main section headers in the sheet. */
const HEADER_BACKGROUND_COLOR = "#F8BBD0"; 
/** @const {string} Hex color code for font in main section headers. */
const HEADER_FONT_COLOR = "#4E342E";       
/** @const {string} Hex color code for sub-headers (column titles) in the sheet. */
const SUB_HEADER_BACKGROUND_COLOR = "#FCE4EC"; 
/** @const {string} Hex color code for cell borders in the sheet. */
const BORDER_COLOR = "#C8A2C8";


// --- LEGACY ORCHESTRATION CONSTANTS (Comment out or remove if legacy function is fully deprecated) ---
/*
const ORCH_API_CALL_DELAY = 10000; // Original name, use _MS for consistency if uncommented
const ORCH_SKIP_TAILORING_IF_SCORE_BELOW = 0.3;
const ORCH_FINAL_INCLUSION_THRESHOLD = 0.7;
const ORCH_MAX_BULLETS_TO_SCORE_PER_JOB = 3;
const ORCH_MAX_BULLETS_TO_INCLUDE_PER_JOB = 2;
const ORCH_MAX_PROJECT_DESC_BULLETS_TO_SCORE = 2;
const ORCH_MAX_SPR_BULLETS_TO_SCORE_PER_CATEGORY = 1;
const ORCH_MAX_PROJECT_FALLBACK_BULLETS_TO_SCORE = 1;
const ORCH_MAX_BULLETS_TO_INCLUDE_PER_PROJECT_TOTAL = 2;
*/

// --- OTHER GLOBAL SETTINGS ---

/** @const {string} A version string for the resume data schema produced by MasterResumeData.gs. */
const CURRENT_RESUME_DATA_SCHEMA_VERSION = "1.3.0_Refactored"; // Example updated version

// Note: API Key Property NAMES (e.g., 'GEMINI_API_KEY', 'GROQ_API_KEY') are best kept
// as local consts within their respective service files (GeminiService.gs, GroqService.gs)
// as they are specific to how those services retrieve their secrets from PropertiesService.
