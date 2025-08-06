// Main.gs
// This file contains the main orchestration logic for the AI Resume Tailoring script,
// focusing on the current staged processing (Stage 1, 2, 3).
// It relies on global constants defined in 'Constants.gs'.

// --- LEGACY ORCHESTRATION FUNCTION (Commented Out - For Reference Only) ---
/*
function orchestrateTailoringProcess(masterResumeObject, jobDescriptionText) {
  Logger.log("--- Starting LEGACY Resume Tailoring Orchestration ---");

  if (ORCH_FINAL_INCLUSION_THRESHOLD < ORCH_SKIP_TAILORING_IF_SCORE_BELOW) {
    Logger.log("WARNING (Legacy Orch): ORCH_FINAL_INCLUSION_THRESHOLD is less than ORCH_SKIP_TAILORING_IF_SCORE_BELOW. This might lead to unexpected exclusions.");
  }
  if (!masterResumeObject || !masterResumeObject.personalInfo) {
    Logger.log("ERROR (Legacy Orch): Invalid or incomplete Master Resume object provided.");
    return { error: "Invalid Master Resume object." };
  }
  const pi = masterResumeObject.personalInfo;
  if (!jobDescriptionText || !jobDescriptionText.trim()) {
    Logger.log("ERROR (Legacy Orch): Invalid or empty Job Description text provided.");
    return { error: "Invalid Job Description text." };
  }

  // Step 1: Analyze Job Description
  Logger.log("Legacy Orch - Step 1: Analyzing Job Description...");
  Utilities.sleep(ORCH_API_CALL_DELAY); // Uses legacy constant
  const jdAnalysis = analyzeJobDescription(jobDescriptionText);
  if (!jdAnalysis || jdAnalysis.error) {
    Logger.log(`ERROR (Legacy Orch): Job Description analysis failed. Details: ${JSON.stringify(jdAnalysis)}`);
    return { error: "Job Description analysis failed.", details: jdAnalysis };
  }
  Logger.log("Legacy Orch - Step 1: Job Description Analyzed successfully.");

  const tailoredResumeObject = {
    resumeSchemaVersion: masterResumeObject.resumeSchemaVersion, // Consider using global CURRENT_RESUME_SCHEMA_VERSION
    personalInfo: JSON.parse(JSON.stringify(pi)), // Deep copy
    summary: "",
    sections: []
  };

  // --- EXPERIENCE SECTION (LEGACY ORCHESTRATION) ---
  Logger.log("\nLegacy Orch - Step 2: Processing EXPERIENCE section...");
  const masterExperienceSection = masterResumeObject.sections.find(s => s && s.title === "EXPERIENCE");
  const tailoredExperienceSection = { title: "EXPERIENCE", items: [] };

  if (masterExperienceSection && masterExperienceSection.items) {
    for (const job of masterExperienceSection.items) {
      const tailoredJob = {
        jobTitle: job.jobTitle,
        company: job.company,
        location: job.location,
        startDate: job.startDate,
        endDate: job.endDate,
        responsibilities: []
      };
      let bulletsScoredForThisJob = 0;
      if (job.responsibilities && Array.isArray(job.responsibilities)) {
        for (const originalBullet of job.responsibilities) {
          if (tailoredJob.responsibilities.length >= ORCH_MAX_BULLETS_TO_INCLUDE_PER_JOB) break;
          if (bulletsScoredForThisJob >= ORCH_MAX_BULLETS_TO_SCORE_PER_JOB) break;
          
          bulletsScoredForThisJob++;
          Utilities.sleep(ORCH_API_CALL_DELAY);
          const matchResult = matchResumeSection(originalBullet, jdAnalysis);
          let iScore = 0;
          if (matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number') {
            iScore = matchResult.relevanceScore;
          }
          if (iScore < ORCH_SKIP_TAILORING_IF_SCORE_BELOW) continue;

          Utilities.sleep(ORCH_API_CALL_DELAY);
          const tailoredOut = tailorBulletPoint(originalBullet, jdAnalysis, jdAnalysis.jobTitle);
          let bulletToConsider = originalBullet;
          if (tailoredOut && !tailoredOut.startsWith("ERROR:") && tailoredOut.toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
            bulletToConsider = tailoredOut.trim();
          }
          if (iScore >= ORCH_FINAL_INCLUSION_THRESHOLD) {
            tailoredJob.responsibilities.push(bulletToConsider);
          }
        }
      }
      if (tailoredJob.responsibilities.length > 0) {
        tailoredExperienceSection.items.push(tailoredJob);
      }
    }
  }
  if (tailoredExperienceSection.items.length > 0) {
    tailoredResumeObject.sections.push(tailoredExperienceSection);
  }
  Logger.log("Legacy Orch - Step 2: EXPERIENCE section processed.");

  // --- PROJECTS SECTION (LEGACY ORCHESTRATION - condensed helper, minimal logs retained) ---
  Logger.log("\nLegacy Orch - Step 3: Processing PROJECTS section...");
  const masterProjectsSection = masterResumeObject.sections.find(s => s && s.title === "PROJECTS");
  const tailoredProjectsSection = { title: "PROJECTS", subsections: [] };

  function _processLegacyProjectBullets(originalBulletList, listType, targetArray, projectOverallTracker, maxToScoreForThisListType) {
    // ... (implementation for brevity in this refactored view, was very dense) ...
    if (!originalBulletList || !Array.isArray(originalBulletList)) return; let scoredFromThisList = 0;
    for(const oBul of originalBulletList){ if(projectOverallTracker.count >= ORCH_MAX_BULLETS_TO_INCLUDE_PER_PROJECT_TOTAL)break; if(scoredFromThisList >= maxToScoreForThisListType)break; scoredFromThisList++; Utilities.sleep(ORCH_API_CALL_DELAY); const mR=matchResumeSection(oBul,jdAnalysis); let iS=0; if(mR&&!mR.error&&typeof mR.relevanceScore==='number')iS=mR.relevanceScore; if(iS<ORCH_SKIP_TAILORING_IF_SCORE_BELOW){continue;} Utilities.sleep(ORCH_API_CALL_DELAY); const tO=tailorBulletPoint(oBul,jdAnalysis,jdAnalysis.jobTitle); let bC=oBul; if(tO&&!tO.startsWith("ERROR:")&&tO.toLowerCase()!=="original bullet not suitable for significant tailoring towards this role.")bC=tO.trim(); if(iS>=ORCH_FINAL_INCLUSION_THRESHOLD){if(projectOverallTracker.count<ORCH_MAX_BULLETS_TO_INCLUDE_PER_PROJECT_TOTAL){targetArray.push(bC);projectOverallTracker.count++;}}}    
  }

  if (masterProjectsSection && masterProjectsSection.subsections) {
    masterProjectsSection.subsections.forEach(masterSubSection => {
      const tailoredSubSection = { name: masterSubSection.name, items: [] };
      if (masterSubSection.items && Array.isArray(masterSubSection.items)) {
        masterSubSection.items.forEach(project => {
          const tailoredProject = { // ... project structure ...  descriptionBullets: [] };
          let projectOverallBulletTracker = { count: 0 };
          _processLegacyProjectBullets(project.descriptionBullets, "P-Desc", tailoredProject.descriptionBullets, projectOverallBulletTracker, ORCH_MAX_PROJECT_DESC_BULLETS_TO_SCORE);
          // ... more calls to _processLegacyProjectBullets for SPRs and fallbacks ...
          if (projectOverallBulletTracker.count > 0) tailoredSubSection.items.push(tailoredProject);
        });
      }
      if (tailoredSubSection.items.length > 0) tailoredProjectsSection.subsections.push(tailoredSubSection);
    });
  }
  if (tailoredProjectsSection.subsections && tailoredProjectsSection.subsections.some(ss => ss.items && ss.items.length > 0)) {
    tailoredResumeObject.sections.push(tailoredProjectsSection);
  }
  Logger.log("Legacy Orch - Step 3: PROJECTS section processed.");

  // --- SUMMARY GENERATION (LEGACY ORCHESTRATION) ---
  // ... (Legacy summary generation logic, heavily condensed for this view) ...
  Logger.log("\nLegacy Orch - Step 4 & 5: Generating Highlights and Tailored Summary...");
  let topHighlightsText = ""; // build highlights...
  Utilities.sleep(ORCH_API_CALL_DELAY);
  const finalSummary = generateTailoredSummary(topHighlightsText.trim(), jdAnalysis, pi.fullName);
  if (finalSummary && !finalSummary.startsWith("ERROR:")) {
    tailoredResumeObject.summary = finalSummary.trim();
  } else {
    tailoredResumeObject.summary = masterResumeObject.summary || "";
    Logger.log("WARNING (Legacy Orch): Tailored summary generation failed.");
  }
  
  // --- STATIC SECTIONS & FINAL ORDERING (LEGACY ORCHESTRATION) ---
  // ... (Legacy static section and ordering logic) ...
  Logger.log("\nLegacy Orch - Step 6: Finalizing sections...");
  const finalSectionsArray = [];
  const desiredSectionOrder = ["TECHNICAL SKILLS & CERTIFICATES","EXPERIENCE","PROJECTS","LEADERSHIP & UNIVERSITY INVOLVEMENT","HONORS & AWARDS","EDUCATION"];
  // Loop through desiredSectionOrder, add from tailoredResumeObject if present and not empty,
  // else add from masterResumeObject if present and not empty.
  order.forEach(t=>{ // ... condensed logic for brevity ...  });
  tailoredResumeObject.sections = finalSectionsArray;


  Logger.log("\n--- LEGACY Resume Tailoring Orchestration Finished ---");
  return {
    tailoredResumeObject: tailoredResumeObject,
    jdAnalysis: jdAnalysis
  };
}
*/
// --- END OF LEGACY ORCHESTRATION ---


// --- STAGE 1: ANALYZE JD & SCORE MASTER RESUME BULLETS ---
/**
 * Stage 1 of the resume tailoring process.
 * Analyzes the provided job description and scores all relevant bullets
 * from the master resume against it. Writes results to designated Google Sheets.
 * Relies on global constants from 'Constants.gs' (e.g., API_CALL_DELAY_STAGE1_MS, JD_ANALYSIS_SHEET_NAME).
 *
 * @param {string} jobDescriptionText The full text of the job description.
 * @param {string} resumeSpreadsheetId The ID of the Google Spreadsheet.
 * @return {Object} An object indicating success or failure:
 *                  { success: boolean, message: string, details?: any, jdAnalysis?: object }
 */
function runStage1_AnalyzeAndScore(jobDescriptionText, resumeSpreadsheetId) {
  Logger.log("--- STAGE 1: Starting JD Analysis & Master Resume Scoring ---");

  // Input Validation
  if (!jobDescriptionText || !jobDescriptionText.trim()) {
    Logger.log("ERROR (Stage 1): Job Description text is empty.");
    return { success: false, message: "Job Description text cannot be empty." };
  }
  if (!resumeSpreadsheetId || resumeSpreadsheetId.toUpperCase().includes("YOUR_ACTUAL") || resumeSpreadsheetId.trim() === "") { // Basic check for placeholder
    Logger.log("ERROR (Stage 1): Master Resume Spreadsheet ID is invalid or placeholder.");
    return { success: false, message: "Master Resume Spreadsheet ID is required and appears to be a placeholder." };
  }

  // Load Master Resume
  // Assumes MASTER_RESUME_DATA_SHEET_NAME is defined in Constants.gs
  const masterResumeObject = getMasterResumeData(resumeSpreadsheetId, MASTER_RESUME_DATA_SHEET_NAME);
  if (!masterResumeObject || !masterResumeObject.personalInfo) {
    Logger.log(`ERROR (Stage 1): Failed to load master resume data from SSID: ${resumeSpreadsheetId}, Sheet: "${MASTER_RESUME_DATA_SHEET_NAME}".`);
    return { success: false, message: `Could not load master resume data. Check sheet ID, name, and contents.` };
  }
  Logger.log("STAGE 1: Master Resume loaded successfully.");

  // Analyze Job Description
  Logger.log("STAGE 1: Analyzing Job Description via LLM...");
  Utilities.sleep(API_CALL_DELAY_STAGE1_MS); // Uses constant from Constants.gs
  const jdAnalysis = analyzeJobDescription(jobDescriptionText); // Calls function in TailoringLogic.gs
  if (!jdAnalysis || jdAnalysis.error) {
    Logger.log(`ERROR (Stage 1): Job Description analysis failed. Details: ${JSON.stringify(jdAnalysis)}`);
    return { success: false, message: "Job Description analysis failed.", details: jdAnalysis };
  }
  Logger.log("STAGE 1: Job Description Analyzed successfully.");

  // Prepare and Write JD Analysis to Sheet
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(resumeSpreadsheetId);
    let jdSheet = spreadsheet.getSheetByName(JD_ANALYSIS_SHEET_NAME); // Uses constant from Constants.gs
    if (!jdSheet) {
      jdSheet = spreadsheet.insertSheet(JD_ANALYSIS_SHEET_NAME);
      Logger.log(`Sheet "${JD_ANALYSIS_SHEET_NAME}" created for JD Analysis.`);
    }
    jdSheet.clearContents();
    jdSheet.getRange(1, 1).setValue("JD Analysis JSON:").setFontWeight("bold");
    const jdAnalysisString = JSON.stringify(jdAnalysis, null, 2);
    if (jdAnalysisString.length > 49000) {
      jdSheet.getRange(2, 1).setValue(jdAnalysisString.substring(0, 49000) + "\n... (TRUNCATED due to cell limit)");
      Logger.log("WARNING (Stage 1): JDAnalysis JSON was very large and truncated in the sheet.");
    } else {
      jdSheet.getRange(2, 1).setValue(jdAnalysisString);
    }
    try { jdSheet.autoResizeColumn(1); } catch (e) { Logger.log(`Note (Stage 1): Could not auto-resize column for ${JD_ANALYSIS_SHEET_NAME}. ${e.message}`); }
    Logger.log(`STAGE 1: JD Analysis data stored in sheet "${JD_ANALYSIS_SHEET_NAME}".`);
  } catch (e) {
    Logger.log(`ERROR (Stage 1): Could not store JD Analysis data in sheet: ${e.toString()}`);
    return { success: false, message: `Error storing JD Analysis: ${e.toString()}` };
  }

  // Prepare Bullet Scoring Results Sheet
  let scoringSheet;
  try {
    scoringSheet = spreadsheet.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME); // Uses constant from Constants.gs
    if (!scoringSheet) {
      scoringSheet = spreadsheet.insertSheet(BULLET_SCORING_RESULTS_SHEET_NAME);
      Logger.log(`Sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" created for bullet scores.`);
    }
    scoringSheet.clearContents();
    const headers = ["UniqueID", "Section", "ItemIdentifier", "OriginalBulletText", "RelevanceScore", "MatchingKeywords", "Justification", "SelectToTailor(Manual)", "TailoredBulletText(Stage2)"];
    scoringSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    scoringSheet.setFrozenRows(1);
    Logger.log(`STAGE 1: Bullet Scoring Results sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" prepared.`);
  } catch (e) {
    Logger.log(`ERROR (Stage 1): Could not prepare Bullet Scoring sheet: ${e.toString()}`);
    return { success: false, message: `Error preparing scoring sheet: ${e.toString()}` };
  }

  // Score Master Resume Bullets
  Logger.log("STAGE 1: Starting LLM scoring of master resume bullets...");
  const allScoredBulletsDataForSheet = [];
  let masterBulletCounter = 0;

  // --- Process EXPERIENCE Bullets ---
  Logger.log("STAGE 1: Processing 'EXPERIENCE' bullets for scoring...");
  const experienceSection = masterResumeObject.sections.find(s => s.title === "EXPERIENCE");
  if (experienceSection && experienceSection.items) {
    experienceSection.items.forEach((job, jobIndex) => {
      if (job.responsibilities && Array.isArray(job.responsibilities)) {
        job.responsibilities.forEach((originalBullet, bulletIndex) => {
          if (!originalBullet || !originalBullet.trim()) return;
          masterBulletCounter++;
          const uniqueId = `EXP_J${jobIndex}_B${bulletIndex}`;
          // Logger.log_DEBUG(`  Scoring Exp. Bullet #${masterBulletCounter} (ID:${uniqueId}): "${originalBullet.substring(0, 60)}..."`);
          Utilities.sleep(API_CALL_DELAY_STAGE1_MS); 
          const matchResult = matchResumeSection(originalBullet, jdAnalysis);
          
          let score = 0.0; let keywords = []; let justification = "Match error/no numeric score.";
          if (matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number') {
            score = matchResult.relevanceScore; keywords = matchResult.matchingKeywords || [];
            justification = matchResult.justification || "No justification by AI.";
          } else if (matchResult && matchResult.error) {
            justification = `Error: ${matchResult.error}. Raw: ${matchResult.rawOutput || ""}`;
          }
          allScoredBulletsDataForSheet.push([
            uniqueId, "EXPERIENCE", job.company || `ExperienceItem${jobIndex}`, originalBullet,
            score.toFixed(2), keywords.join("; "), justification, "", ""
          ]);
        });
      }
    });
  } else {
    Logger.log("STAGE 1: No 'EXPERIENCE' section or items found in master resume.");
  }

  // --- Process PROJECTS Bullets ---
  Logger.log("STAGE 1: Processing 'PROJECTS' bullets for scoring...");
  const projectsSection = masterResumeObject.sections.find(s => s.title === "PROJECTS");
  if (projectsSection && projectsSection.subsections) {
    projectsSection.subsections.forEach((subsection, subIndex) => { // Added subIndex for more unique IDs if needed
      if (subsection.items && Array.isArray(subsection.items)) {
        subsection.items.forEach((project, projectIndex) => {
          const baseProjectId = `PROJ_S${subIndex}_P${projectIndex}_${(project.projectName || "NoName").replace(/\s+/g, "_").substring(0, 10)}`;
          let projectBulletsToProcess = [];
          if (project.descriptionBullets && Array.isArray(project.descriptionBullets)) {
            projectBulletsToProcess.push(...project.descriptionBullets.map(b => ({ text: b, type: "Desc" }))); // Shortened type
          }
          if (project.subProjectsOrRoles && Array.isArray(project.subProjectsOrRoles)) {
            project.subProjectsOrRoles.forEach((spr, sprIndex) => {
              if (spr.responsibilities && Array.isArray(spr.responsibilities)) {
                projectBulletsToProcess.push(...spr.responsibilities.map(b => ({ text: b, type: `SPR${sprIndex}` })));
              }
            });
          }
          if (project.responsibilities && Array.isArray(project.responsibilities) && projectBulletsToProcess.length === 0) { 
            projectBulletsToProcess.push(...project.responsibilities.map(b => ({ text: b, type: "MainRespFallbk" })));
          }

          projectBulletsToProcess.forEach((bulletObject, bulletIndexInList) => {
            if (!bulletObject.text || !bulletObject.text.trim()) return;
            masterBulletCounter++;
            const uniqueId = `${baseProjectId}_${bulletObject.type}_B${bulletIndexInList}`;
            // Logger.log_DEBUG(`  Scoring Proj. Bullet #${masterBulletCounter} (ID:${uniqueId}): "${bulletObject.text.substring(0, 60)}..."`);
            Utilities.sleep(API_CALL_DELAY_STAGE1_MS); 
            const matchResult = matchResumeSection(bulletObject.text, jdAnalysis);

            let score = 0.0; let keywords = []; let justification = "Match error/no numeric score.";
            if (matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number') {
              score = matchResult.relevanceScore; keywords = matchResult.matchingKeywords || [];
              justification = matchResult.justification || "No justification by AI.";
            } else if (matchResult && matchResult.error) {
              justification = `Error: ${matchResult.error}. Raw: ${matchResult.rawOutput || ""}`;
            }
            allScoredBulletsDataForSheet.push([
              uniqueId, "PROJECTS", project.projectName || `ProjectItem${projectIndex}`, bulletObject.text,
              score.toFixed(2), keywords.join("; "), justification, "", ""
            ]);
          });
        });
      }
    });
  } else {
    Logger.log("STAGE 1: No 'PROJECTS' section or subsections/items found in master resume.");
  }
  
  // --- Process LEADERSHIP & UNIVERSITY INVOLVEMENT Bullets ---
  Logger.log("STAGE 1: Processing 'LEADERSHIP & UNIVERSITY INVOLVEMENT' bullets for scoring...");
  const leadershipSection = masterResumeObject.sections.find(s => s.title === "LEADERSHIP & UNIVERSITY INVOLVEMENT");
  if (leadershipSection && leadershipSection.items) {
    leadershipSection.items.forEach((item, itemIndex) => {
      // Use organization or role as identifier, fall back to indexed name
      const itemIdentifier = item.organization || item.role || `LeadershipItem${itemIndex}`; 
      if (item.responsibilities && Array.isArray(item.responsibilities)) {
        item.responsibilities.forEach((originalBullet, bulletIndex) => {
          if (!originalBullet || !originalBullet.trim()) return; 
          
          masterBulletCounter++;
          const uniqueId = `LEAD_I${itemIndex}_B${bulletIndex}`;
          // Logger.log_DEBUG(`  Scoring Leadership Bullet #${masterBulletCounter} (ID:${uniqueId}): "${originalBullet.substring(0, 60)}..."`);
          
          Utilities.sleep(API_CALL_DELAY_STAGE1_MS);
          const matchResult = matchResumeSection(originalBullet, jdAnalysis);
          
          let score = 0.0; let keywords = []; let justification = "Match error/no numeric score.";
          if (matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number') {
            score = matchResult.relevanceScore; keywords = matchResult.matchingKeywords || [];
            justification = matchResult.justification || "No justification provided by AI.";
          } else if (matchResult && matchResult.error) {
            justification = `Error: ${matchResult.error}. Raw: ${matchResult.rawOutput || ""}`;
          }
          
          allScoredBulletsDataForSheet.push([
            uniqueId, "LEADERSHIP & UNIVERSITY INVOLVEMENT", itemIdentifier, originalBullet,
            score.toFixed(2), keywords.join("; "), justification, "", ""
          ]);
        });
      }
    });
  } else {
    Logger.log("STAGE 1: No 'LEADERSHIP & UNIVERSITY INVOLVEMENT' section or items found in master resume.");
  }
  
  // --- Process TECHNICAL SKILLS & CERTIFICATES ---
  Logger.log("STAGE 1: Processing 'TECHNICAL SKILLS & CERTIFICATES' for scoring...");
  const techSkillsSection = masterResumeObject.sections.find(s => s.title === "TECHNICAL SKILLS & CERTIFICATES");
  if (techSkillsSection && techSkillsSection.subsections) {
    techSkillsSection.subsections.forEach((subsection, subIndex) => {
      const categoryName = subsection.name || `TechCategory${subIndex}`;
      if (subsection.items && Array.isArray(subsection.items)) {
        subsection.items.forEach((item, itemIndex) => {
          let originalTextToScore = "";
          let itemIdentifier = ""; // For the sheet
          let itemType = ""; // For UniqueID, helps differentiate skill vs cert

          if (item.skill) { // It's a skill item
            originalTextToScore = item.skill.trim();
            if (item.details && item.details.trim()) {
              originalTextToScore += ` ${item.details.trim()}`; // Add details for context
            }
            itemIdentifier = item.skill.trim();
            itemType = "Skill";
          } else if (item.name) { // It's likely a certificate/license item
            originalTextToScore = item.name.trim();
            if (item.issuer && item.issuer.trim()) {
              originalTextToScore += ` (Issuer: ${item.issuer.trim()})`;
            }
            // Could also add issueDate if relevant for scoring context, e.g., `Issued: ${item.issueDate}`
            itemIdentifier = item.name.trim();
            itemType = "Cert";
          }

          if (!originalTextToScore) {
            // Logger.log_DEBUG(`  Skipping empty skill/cert item in category "${categoryName}".`);
            return; // Skip if no text to score
          }
          
          masterBulletCounter++;
          const uniqueId = `TECH_Cat${subIndex}_${itemType}${itemIndex}`;
          // Logger.log_DEBUG(`  Scoring Tech Item #${masterBulletCounter} (ID:${uniqueId}, Cat:${categoryName}): "${originalTextToScore.substring(0, 70)}..."`);
          
          Utilities.sleep(API_CALL_DELAY_STAGE1_MS); // From Constants.gs
          const matchResult = matchResumeSection(originalTextToScore, jdAnalysis);
          
          let score = 0.0;
          let keywords = [];
          let justification = "Matching process failed, returned an error, or score was not a number.";
          
          if (matchResult && !matchResult.error && typeof matchResult.relevanceScore === 'number') {
            score = matchResult.relevanceScore;
            keywords = matchResult.matchingKeywords || [];
            justification = matchResult.justification || "No justification provided by AI.";
          } else if (matchResult && matchResult.error) {
            justification = `Error during matching: ${matchResult.error}. Raw: ${matchResult.rawOutput || ""}`;
          }
          
          allScoredBulletsDataForSheet.push([
            uniqueId,
            "TECHNICAL SKILLS & CERTIFICATES", // Section Name
            itemIdentifier,                    // ItemIdentifier (the skill or cert name itself)
            originalTextToScore,               // OriginalText (might include details/issuer)
            score.toFixed(2),
            keywords.join("; "),
            justification,
            "", // Placeholder for SelectToTailor (Can skills/certs be "tailored" by LLM, or just included?)
            ""  // Placeholder for TailoredText (Likely N/A if not tailoring them)
          ]);
        });
      }
    });
  } else {
    Logger.log("STAGE 1: No 'TECHNICAL SKILLS & CERTIFICATES' section or subsections/items found in master resume.");
  }

  // Write Scored Bullets to Sheet
  if (allScoredBulletsDataForSheet.length > 0) {
    scoringSheet.getRange(2, 1, allScoredBulletsDataForSheet.length, allScoredBulletsDataForSheet[0].length)
                .setValues(allScoredBulletsDataForSheet);
    try {
      for (let i = 1; i <= allScoredBulletsDataForSheet[0].length; i++) {
        scoringSheet.autoResizeColumn(i);
      }
    } catch (e) { Logger.log(`Note (Stage 1): Could not auto-resize columns for scoring sheet. ${e.message}`); }
    Logger.log(`STAGE 1: Scored ${allScoredBulletsDataForSheet.length} bullets. Results saved to sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);
  } else {
    Logger.log("STAGE 1: No bullets were processed or scored. Ensure Experience/Projects sections in master resume have content.");
  }

  Logger.log("--- STAGE 1: JD Analysis & Master Resume Scoring Finished ---");
  return {
    success: true,
    message: `Stage 1 Complete. JD Analyzed. ${allScoredBulletsDataForSheet.length} bullets scored. Results in sheet '${BULLET_SCORING_RESULTS_SHEET_NAME}'.`,
    jdAnalysis: jdAnalysis // Pass jdAnalysis for potential immediate use by a caller managing stages
  };
}
// --- END OF STAGE 1 FUNCTION ---


// --- STAGE 2: TAILOR SELECTED BULLETS ---
/**
 * Stage 2 of the resume tailoring process.
 * Reads bullets marked for tailoring from the 'BulletScoringResults' sheet,
 * uses an LLM to tailor them, and updates the sheet with the tailored text.
 * Relies on global constants from 'Constants.gs' (e.g., API_CALL_DELAY_STAGE2_MS, JD_ANALYSIS_SHEET_NAME).
 *
 * @param {string} resumeSpreadsheetId The ID of the Google Sheet.
 * @return {Object} An object indicating success or failure:
 *                  { success: boolean, message: string, bulletsTailored?: number }
 */
function runStage2_TailorSelectedBullets(resumeSpreadsheetId) {
  Logger.log("--- STAGE 2: Starting Tailoring of Selected Bullets ---");

  // Input Validation & Data Loading
  if (!resumeSpreadsheetId) {
    Logger.log("ERROR (Stage 2): Spreadsheet ID is required.");
    return { success: false, message: "Spreadsheet ID required." };
  }

  let spreadsheet, jdAnalysis;
  let scoringSheetDataRows; // To hold the array of data rows for updating
  let scoringSheetHeaders; // To hold the header row

  try {
    spreadsheet = SpreadsheetApp.openById(resumeSpreadsheetId);
    const jdAnalysisSheet = spreadsheet.getSheetByName(JD_ANALYSIS_SHEET_NAME); // Uses constant from Constants.gs
    const scoringSheetFromStage1 = spreadsheet.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME); // Uses constant from Constants.gs

    if (!jdAnalysisSheet) {
      Logger.log(`ERROR (Stage 2): Sheet "${JD_ANALYSIS_SHEET_NAME}" for JD analysis not found.`);
      return { success: false, message: `Sheet "${JD_ANALYSIS_SHEET_NAME}" missing.` };
    }
    if (!scoringSheetFromStage1) {
      Logger.log(`ERROR (Stage 2): Sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" for bullet scores not found.`);
      return { success: false, message: `Sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" missing.` };
    }

    const jdAnalysisJsonString = jdAnalysisSheet.getRange(2, 1).getValue();
    if (!jdAnalysisJsonString || typeof jdAnalysisJsonString !== 'string' || !jdAnalysisJsonString.trim()) {
      Logger.log("ERROR (Stage 2): JD Analysis JSON string not found or empty in its sheet.");
      return { success: false, message: "JD Analysis JSON string not found." };
    }
    jdAnalysis = JSON.parse(jdAnalysisJsonString);
    if (!jdAnalysis || !jdAnalysis.jobTitle) {
      Logger.log("ERROR (Stage 2): Parsed JD Analysis from sheet is invalid or incomplete.");
      return { success: false, message: "Parsed JD Analysis invalid." };
    }
    Logger.log("STAGE 2: JD Analysis loaded successfully from sheet.");

    const allScoringDataWithHeaders = scoringSheetFromStage1.getDataRange().getValues();
    if (allScoringDataWithHeaders.length < 2) { // Must have headers + at least one data row
      Logger.log("STAGE 2: No data rows found in scoring sheet. Nothing to tailor.");
      return { success: true, message: "No data in scoring sheet to tailor.", bulletsTailored: 0 };
    }
    scoringSheetHeaders = allScoringDataWithHeaders.shift(); // Get headers, allScoringDataWithHeaders now only data
    scoringSheetDataRows = allScoringDataWithHeaders;

  } catch (e) {
    Logger.log(`ERROR (Stage 2): Exception during data loading: ${e.toString()}`);
    return { success: false, message: `Error accessing Stage 2 data sheets: ${e.toString()}` };
  }

  // Get column indices
  const originalBulletColIndex = scoringSheetHeaders.indexOf("OriginalBulletText");
  const selectColIndex = scoringSheetHeaders.indexOf("SelectToTailor(Manual)");
  const tailoredColIndex = scoringSheetHeaders.indexOf("TailoredBulletText(Stage2)");
  const scoreColIndex = scoringSheetHeaders.indexOf("RelevanceScore");
  const uniqueIdColIndex = scoringSheetHeaders.indexOf("UniqueID");

  if ([originalBulletColIndex, selectColIndex, tailoredColIndex, scoreColIndex, uniqueIdColIndex].includes(-1)) {
    const missingCols = ["OriginalBulletText", "SelectToTailor(Manual)", "TailoredBulletText(Stage2)", "RelevanceScore", "UniqueID"]
      .filter(h => scoringSheetHeaders.indexOf(h) === -1).join(", ");
    Logger.log(`ERROR (Stage 2): Required columns missing in scoring sheet: ${missingCols}.`);
    return { success: false, message: `Required columns missing in scoring sheet: ${missingCols}.` };
  }

  // Tailor Selected Bullets
  let bulletsAttemptedToTailor = 0;
  let bulletsSuccessfullyTailored = 0;
  let modifiedDataForSheet = false;

  for (let i = 0; i < scoringSheetDataRows.length; i++) {
    const rowData = scoringSheetDataRows[i];
    const shouldTailorSelection = String(rowData[selectColIndex]).toUpperCase().trim();

    if (["YES", "TRUE", "1", "X"].includes(shouldTailorSelection)) {
      bulletsAttemptedToTailor++;
      modifiedDataForSheet = true; // Mark that we've modified something
      const originalBullet = rowData[originalBulletColIndex];
      const uniqueId = rowData[uniqueIdColIndex];
      // const originalScore = rowData[scoreColIndex]; // DEBUG
      // Logger.log_DEBUG(`  Tailoring bullet ID:${uniqueId} (Score ${originalScore}): "${originalBullet.substring(0,60)}..."`);
      
      Utilities.sleep(API_CALL_DELAY_STAGE2_MS); // Uses constant from Constants.gs
      const tailoredOutput = tailorBulletPoint(originalBullet, jdAnalysis, jdAnalysis.jobTitle);

      if (tailoredOutput && !tailoredOutput.startsWith("ERROR:") && tailoredOutput.toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
        scoringSheetDataRows[i][tailoredColIndex] = tailoredOutput.trim(); // Update in the array for batch write
        bulletsSuccessfullyTailored++;
        // Logger.log_DEBUG(`    SUCCESS tailoring ID:${uniqueId}.`);
      } else {
        scoringSheetDataRows[i][tailoredColIndex] = `TAILOR_FAIL/SKIP: ${tailoredOutput}`;
        Logger.log(`WARN (Stage 2): Failed or skipped tailoring for bullet ID:${uniqueId}. LLM Message: ${tailoredOutput}`);
      }
    }
  }
  
  // Write updated data back to the sheet if any modifications were made
  if (modifiedDataForSheet && scoringSheetDataRows.length > 0) {
    try {
      const targetSheet = spreadsheet.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME);
      targetSheet.getRange(2, 1, scoringSheetDataRows.length, scoringSheetHeaders.length).setValues(scoringSheetDataRows);
      Logger.log(`STAGE 2: Updated "${BULLET_SCORING_RESULTS_SHEET_NAME}" with tailored bullet text.`);
    } catch (e) {
      Logger.log(`ERROR (Stage 2): Failed to write tailored bullets back to sheet. ${e.toString()}`);
      // Continue to return status but acknowledge write failure.
    }
  }

  Logger.log(`--- STAGE 2: Tailoring Finished. Attempted: ${bulletsAttemptedToTailor}, Succeeded: ${bulletsSuccessfullyTailored} ---`);
  return {
    success: true,
    message: `Stage 2 Complete. Attempted to tailor ${bulletsAttemptedToTailor} bullets, successfully tailored ${bulletsSuccessfullyTailored}. Sheet '${BULLET_SCORING_RESULTS_SHEET_NAME}' updated.`,
    bulletsTailored: bulletsSuccessfullyTailored
  };
}
// --- END OF STAGE 2 FUNCTION ---


// --- STAGE 3: ASSEMBLE RESUME OBJECT & GENERATE DOCUMENT ---
/**
 * Stage 3 of the resume tailoring process.
 * Assembles the final `finalTailoredResumeObject` and generates the Google Document.
 * Relies on global constants from 'Constants.gs' (e.g., STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD).
 *
 * @param {string} resumeSpreadsheetId The ID of the Google Sheet.
 * @return {Object} An object indicating success or failure.
 */

// In Main.gs
// This is the refined runStage3_BuildAndGenerateDocument function
// incorporating strict "Score + YES" selection logic for all dynamic sections.

/**
 * Stage 3 of the resume tailoring process.
 * Assembles the final `finalTailoredResumeObject` based on data from Stage 1 & 2 sheets
 * (master resume, JD analysis, scored/selected/tailored bullets from all relevant sections).
 * This version implements strict filtering: items (jobs, projects, leadership roles) and their
 * bullets are only included if they meet score thresholds AND are marked "YES" by the user.
 * Technical skills also follow this strict selection.
 * Generates an AI-powered summary.
 * Creates the final, formatted Google Document by calling createFormattedResumeDoc.
 * Relies on global constants from 'Constants.gs'.
 *
 * @param {string} resumeSpreadsheetId The ID of the Google Sheet.
 * @return {Object} An object indicating success or failure:
 *                  { success: boolean, message: string, docUrl?: string, tailoredResumeObjectForDebug?: Object }
 */
function runStage3_BuildAndGenerateDocument(resumeSpreadsheetId) {
  Logger.log("--- STAGE 3: Starting Final Resume Assembly & Document Generation (Strict Selection v61 Logic) ---");

  if (!resumeSpreadsheetId) {
    Logger.log("ERROR (Stage 3): Spreadsheet ID is required.");
    return { success: false, message: "Spreadsheet ID required for Stage 3." };
  }

  // --- 1. Load Prerequisite Data (Master Resume, JD Analysis, Scored/Tailored Bullets) ---
  let masterResumeObject, jdAnalysis, allScoredBulletRows;
  try {
    Logger.log("STAGE 3 - Step 1a: Loading Master Resume Data...");
    masterResumeObject = getMasterResumeData(resumeSpreadsheetId, MASTER_RESUME_DATA_SHEET_NAME); // Uses constants
    if (!masterResumeObject || !masterResumeObject.personalInfo) {
      Logger.log("ERROR (Stage 3): Failed to load master resume data or personalInfo is missing.");
      return { success: false, message: "Failed to load master resume data for Stage 3." };
    }
    Logger.log("STAGE 3: Master Resume loaded successfully.");

    const spreadsheet = SpreadsheetApp.openById(resumeSpreadsheetId);
    const jdAnalysisSheet = spreadsheet.getSheetByName(JD_ANALYSIS_SHEET_NAME); // Uses constants
    const scoringSheet = spreadsheet.getSheetByName(BULLET_SCORING_RESULTS_SHEET_NAME); // Uses constants

    if (!jdAnalysisSheet) {
        Logger.log(`ERROR (Stage 3): JD Analysis sheet "${JD_ANALYSIS_SHEET_NAME}" not found.`);
        return { success: false, message: `Sheet "${JD_ANALYSIS_SHEET_NAME}" not found.` };
    }
    if (!scoringSheet) {
        Logger.log(`ERROR (Stage 3): Bullet Scoring sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" not found.`);
        return { success: false, message: `Sheet "${BULLET_SCORING_RESULTS_SHEET_NAME}" not found.` };
    }

    Logger.log("STAGE 3 - Step 1b: Loading JD Analysis from sheet...");
    const jdAnalysisJsonString = jdAnalysisSheet.getRange(2, 1).getValue();
    if (!jdAnalysisJsonString || typeof jdAnalysisJsonString !== 'string' || jdAnalysisJsonString.trim() === "") {
      Logger.log("ERROR (Stage 3): JD Analysis JSON string not found or empty in its sheet.");
      return { success: false, message: "JD Analysis JSON string not found or empty." };
    }
    jdAnalysis = JSON.parse(jdAnalysisJsonString);
    if (!jdAnalysis || !jdAnalysis.jobTitle) {
      Logger.log("ERROR (Stage 3): Parsed JD Analysis data from sheet is invalid or missing jobTitle.");
      return { success: false, message: "JD Analysis data invalid." };
    }
    Logger.log("STAGE 3: JD Analysis loaded successfully from sheet.");

    Logger.log("STAGE 3 - Step 1c: Loading Scored/Tailored Bullets from sheet...");
    const scoringDataRange = scoringSheet.getDataRange();
    const allScoringDataWithHeaders = scoringDataRange.getValues();
    if (allScoringDataWithHeaders.length < 2) {
      Logger.log(`INFO (Stage 3): No scored bullet data found in "${BULLET_SCORING_RESULTS_SHEET_NAME}". Dynamic sections might be sparse or use static fallback.`);
      allScoredBulletRows = [];
    } else {
      const headers = allScoringDataWithHeaders.shift(); // Remove header row
      allScoredBulletRows = allScoringDataWithHeaders.map(row => {
        let obj = {}; headers.forEach((header, index) => obj[header] = row[index]); return obj;
      });
      Logger.log(`STAGE 3: Loaded ${allScoredBulletRows.length} scored entries from "${BULLET_SCORING_RESULTS_SHEET_NAME}".`);
    }
  } catch (e) {
    Logger.log(`ERROR (Stage 3 - Data Loading): ${e.toString()}\nStack: ${e.stack}`);
    return { success: false, message: `Error loading prerequisite data for Stage 3: ${e.toString()}` };
  }
  // --- End of Prerequisite Data Loading ---


  // Initialize finalTailoredResumeObject
  const personalInfoCopy = JSON.parse(JSON.stringify(masterResumeObject.personalInfo));
  const finalTailoredResumeObject = {
    resumeSchemaVersion: (typeof CURRENT_RESUME_DATA_SCHEMA_VERSION !== 'undefined' ? CURRENT_RESUME_DATA_SCHEMA_VERSION : masterResumeObject.resumeSchemaVersion),
    personalInfo: personalInfoCopy,
    summary: "", // To be AI generated
    sections: []
  };
  let allIncludedContentForSummaryHighlights = [];

  // --- Helper function to get the bullet text (tailored or original) ---
  function getBulletTextToUse(scoredBulletEntry) {
    let bulletTextToUse = scoredBulletEntry.OriginalBulletText;
    const tailoredTextFromSheet = scoredBulletEntry["TailoredBulletText(Stage2)"];
    if (tailoredTextFromSheet &&
        String(tailoredTextFromSheet).trim() !== "" &&
        !String(tailoredTextFromSheet).toUpperCase().startsWith("TAILOR_FAIL") &&
        String(tailoredTextFromSheet).toLowerCase() !== "original bullet not suitable for significant tailoring towards this role.") {
      bulletTextToUse = String(tailoredTextFromSheet).trim();
    }
    return bulletTextToUse;
  }

  // --- DYNAMIC SECTION ASSEMBLY ---
  // Logic: Iterate master items. For each, find its scored bullets. Filter by score AND "YES". 
  // If any bullets pass, include the master item shell with ONLY those selected/processed bullets.

  // Section: EXPERIENCE
  Logger.log("STAGE 3 - Assembling 'EXPERIENCE' (Strict Selection: Score + User 'YES')");
  const tailoredExperienceSection = { title: "EXPERIENCE", items: [] };
  const masterExperienceItems = (masterResumeObject.sections.find(s => s.title === "EXPERIENCE") || {}).items || [];
  masterExperienceItems.forEach(masterJob => {
    const jobBulletsFromSheet = allScoredBulletRows.filter(r => r.Section === "EXPERIENCE" && r.ItemIdentifier === masterJob.company);
    const selectedScoredBullets = jobBulletsFromSheet
      .filter(sb => parseFloat(sb.RelevanceScore) >= STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD && 
                     ["YES", "TRUE", "1", "X"].includes(String(sb["SelectToTailor(Manual)"]).toUpperCase().trim()))
      .sort((a, b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore));

    if (selectedScoredBullets.length > 0) {
      let includedBulletsText = selectedScoredBullets
        .slice(0, STAGE3_MAX_BULLETS_PER_JOB) 
        .map(sb => getBulletTextToUse(sb));
      
      if (includedBulletsText.length > 0) { 
        tailoredExperienceSection.items.push({ 
            ...masterJob, 
            responsibilities: includedBulletsText 
        });
        allIncludedContentForSummaryHighlights.push(...includedBulletsText);
        Logger.log(`  Added Experience: "${masterJob.company}" with ${includedBulletsText.length} selected/scored bullets.`);
      }
    }
  });
  if (tailoredExperienceSection.items.length > 0) finalTailoredResumeObject.sections.push(tailoredExperienceSection);

  // Section: PROJECTS
  Logger.log("STAGE 3 - Assembling 'PROJECTS' (Strict Selection: Score + User 'YES' for bullets)");
  const tailoredProjectsSection = { title: "PROJECTS", subsections: [] };
  const masterProjectsData = masterResumeObject.sections.find(s => s.title === "PROJECTS");
  if (masterProjectsData && masterProjectsData.subsections) {
    masterProjectsData.subsections.forEach(masterSubSection => {
      const tailoredProjectItemsForSubSection = [];
      (masterSubSection.items || []).forEach(masterProject => {
        const projectBulletsFromSheet = allScoredBulletRows.filter(r => r.Section === "PROJECTS" && r.ItemIdentifier === masterProject.projectName);
        const selectedScoredBullets = projectBulletsFromSheet
          .filter(sb => parseFloat(sb.RelevanceScore) >= STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD &&
                         ["YES", "TRUE", "1", "X"].includes(String(sb["SelectToTailor(Manual)"]).toUpperCase().trim()))
          .sort((a, b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore));

        if (selectedScoredBullets.length > 0) {
            let includedBulletsText = selectedScoredBullets
              .slice(0, STAGE3_MAX_BULLETS_PER_PROJECT)
              .map(sb => getBulletTextToUse(sb));

            if (includedBulletsText.length > 0) {
                tailoredProjectItemsForSubSection.push({
                    ...masterProject,
                    descriptionBullets: includedBulletsText
                });
                allIncludedContentForSummaryHighlights.push(...includedBulletsText);
                Logger.log(`  Added Project: "${masterProject.projectName}" with ${includedBulletsText.length} selected/scored bullets.`);
            }
        }
      });
      if (tailoredProjectItemsForSubSection.length > 0) {
        tailoredProjectsSection.subsections.push({ name: masterSubSection.name || "General Projects", items: tailoredProjectItemsForSubSection });
      }
    });
    if (tailoredProjectsSection.subsections.some(ss => ss.items && ss.items.length > 0)) {
      finalTailoredResumeObject.sections.push(tailoredProjectsSection);
    }
  }

  // Section: LEADERSHIP & UNIVERSITY INVOLVEMENT
  Logger.log("STAGE 3 - Assembling 'LEADERSHIP & UNIVERSITY INVOLVEMENT' (Strict Selection: Score + User 'YES')");
  const tailoredLeadershipSection = { title: "LEADERSHIP & UNIVERSITY INVOLVEMENT", items: [] };
  const masterLeadershipItems = (masterResumeObject.sections.find(s => s.title === "LEADERSHIP & UNIVERSITY INVOLVEMENT") || {}).items || [];
  masterLeadershipItems.forEach(masterItem => {
    const itemIdentifierInMaster = masterItem.organization || masterItem.role; 
    const leadershipBulletsFromSheet = allScoredBulletRows.filter(r => 
        r.Section === "LEADERSHIP & UNIVERSITY INVOLVEMENT" && r.ItemIdentifier === itemIdentifierInMaster);
    const selectedScoredBullets = leadershipBulletsFromSheet
      .filter(sb => parseFloat(sb.RelevanceScore) >= STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD &&
                     ["YES", "TRUE", "1", "X"].includes(String(sb["SelectToTailor(Manual)"]).toUpperCase().trim()))
      .sort((a, b) => parseFloat(b.RelevanceScore) - parseFloat(a.RelevanceScore));

    const MAX_BULLETS_PER_LEADERSHIP_ITEM = STAGE3_MAX_BULLETS_PER_JOB; // Using same limit as jobs, or define new in Constants.gs

    if (selectedScoredBullets.length > 0) {
      let includedBulletsText = selectedScoredBullets
        .slice(0, MAX_BULLETS_PER_LEADERSHIP_ITEM)
        .map(sb => getBulletTextToUse(sb));
            
      if (includedBulletsText.length > 0) {
        tailoredLeadershipSection.items.push({
            ...masterItem,
            responsibilities: includedBulletsText // Assuming master leadership items also use 'responsibilities' for their bullets
        });
        allIncludedContentForSummaryHighlights.push(...includedBulletsText);
        Logger.log(`  Added Leadership: "${itemIdentifierInMaster}" with ${includedBulletsText.length} selected/scored bullets.`);
      }
    }
  });
  if (tailoredLeadershipSection.items.length > 0) finalTailoredResumeObject.sections.push(tailoredLeadershipSection);

  // Section: TECHNICAL SKILLS & CERTIFICATES
  Logger.log("STAGE 3 - Assembling 'TECHNICAL SKILLS & CERTIFICATES' (Strict Selection: Score + User 'YES')");
  const tailoredTechSkillsSection = { title: "TECHNICAL SKILLS & CERTIFICATES", subsections: [] };
  const selectedTechItemsFromSheet = allScoredBulletRows.filter(row =>
    row.Section === "TECHNICAL SKILLS & CERTIFICATES" &&
    parseFloat(row.RelevanceScore) >= STAGE3_FINAL_INCLUSION_SCORE_THRESHOLD &&
    ["YES", "TRUE", "1", "X"].includes(String(row["SelectToTailor(Manual)"]).toUpperCase().trim())
  );

  if (selectedTechItemsFromSheet.length > 0) {
    const skillsBySubCategoryMap = new Map();
    const masterTechSkillsData = masterResumeObject.sections.find(s => s.title === "TECHNICAL SKILLS & CERTIFICATES");

    selectedTechItemsFromSheet.forEach(selectedSheetItem => {
      let categoryName = "Uncategorized"; // Default
      let originalMasterItemStructure = null;

      if (masterTechSkillsData && masterTechSkillsData.subsections) {
        for (const masterSub of masterTechSkillsData.subsections) {
          const foundItem = masterSub.items.find(masterItm =>
            (masterItm.skill && masterItm.skill.trim() === selectedSheetItem.ItemIdentifier) || // ItemIdentifier is skill/cert name
            (masterItm.name && masterItm.name.trim() === selectedSheetItem.ItemIdentifier)
          );
          if (foundItem) {
            categoryName = masterSub.name;
            originalMasterItemStructure = { ...foundItem }; // Get a copy of the original item's structure
            break; 
          }
        }
      }
      
      if (!originalMasterItemStructure) {
        Logger.log(`WARN (Stage 3 Skills Assembly): Original master structure not found for selected skill/cert: "${selectedSheetItem.ItemIdentifier}". Attempting to use OriginalBulletText from sheet: "${selectedSheetItem.OriginalBulletText}"`);
        // Fallback: Create a minimal skill structure. This relies on OriginalBulletText containing skill +/- details.
        // If your renderTechnicalSkillsList needs distinct {skill, details} or {name, issuer}, this fallback is very basic.
        originalMasterItemStructure = { skill: selectedSheetItem.OriginalBulletText, details: "" }; 
      }

      if (!skillsBySubCategoryMap.has(categoryName)) {
        skillsBySubCategoryMap.set(categoryName, { name: categoryName, items: [] });
      }
      skillsBySubCategoryMap.get(categoryName).items.push(originalMasterItemStructure);
    });

    tailoredTechSkillsSection.subsections = Array.from(skillsBySubCategoryMap.values())
                                              .filter(sub => sub.items.length > 0); // Only include subsections that ended up with items
    if (tailoredTechSkillsSection.subsections.length > 0) {
      finalTailoredResumeObject.sections.push(tailoredTechSkillsSection);
      Logger.log(`  Assembled Skills: ${selectedTechItemsFromSheet.length} selected items into ${tailoredTechSkillsSection.subsections.length} categories.`);
    }
  }
  // --- End of Dynamic Section Assembly ---

  // --- Generate AI Summary ---
  Logger.log(`STAGE 3 - Generating highlights for AI Summary. Total relevant content pieces: ${allIncludedContentForSummaryHighlights.length}`);
  let topMatchedHighlights = allIncludedContentForSummaryHighlights
                              .sort((a,b) => b.length - a.length) // Simple sort by length for "more detail"
                              .slice(0, STAGE3_MAX_HIGHLIGHTS_FOR_SUMMARY) // From Constants.gs
                              .join(" ");
  if (!topMatchedHighlights.trim()) {
    topMatchedHighlights = masterResumeObject.summary ? masterResumeObject.summary.substring(0, 350) + "..." : "Proactive professional adept at leveraging diverse skills.";
    Logger.log("  No dynamic highlights generated; using master summary snippet for AI input.");
  }

  Logger.log("STAGE 3 - Generating AI Tailored Summary...");
  Utilities.sleep(API_CALL_DELAY_STAGE3_SUMMARY_MS); // From Constants.gs
  const tailoredSummary = generateTailoredSummary(topMatchedHighlights.trim(), jdAnalysis, personalInfoCopy.fullName);
  if (tailoredSummary && !tailoredSummary.startsWith("ERROR:")) {
    finalTailoredResumeObject.summary = tailoredSummary.trim();
    Logger.log("  AI Tailored Summary generated successfully.");
  } else {
    finalTailoredResumeObject.summary = masterResumeObject.summary || ""; // Fallback
    Logger.log(`  WARN (Stage 3): AI Summary generation failed/skipped. Using master summary. LLM Response: ${tailoredSummary}`);
  }

  // --- Add Static Sections (if not dynamically processed) & Ensure Final Order ---
  Logger.log("STAGE 3 - Adding static sections and ensuring final document order...");
  const processedDynamicSectionTitles = finalTailoredResumeObject.sections.map(s => s.title.toUpperCase());
  const desiredFinalSectionOrder = ["EDUCATION", "EXPERIENCE", "PROJECTS", "TECHNICAL SKILLS & CERTIFICATES", "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"]; 
  
  desiredFinalSectionOrder.forEach(sectionTitle => {
    if (!processedDynamicSectionTitles.includes(sectionTitle.toUpperCase())) {
      const staticSectionData = masterResumeObject.sections.find(s => s && s.title && s.title.toUpperCase() === sectionTitle.toUpperCase());
      if (staticSectionData) {
        // Basic check if the static section is empty before adding
        let isStaticSectionEffectivelyEmpty = false;
        if ((staticSectionData.items && staticSectionData.items.length === 0) && 
            (!staticSectionData.subsections || staticSectionData.subsections.length === 0 || staticSectionData.subsections.every(ss => !ss.items || ss.items.length === 0))) {
          isStaticSectionEffectivelyEmpty = true;
        } else if ((!staticSectionData.items || staticSectionData.items.length === 0) && 
                   (staticSectionData.subsections && staticSectionData.subsections.every(ss => !ss.items || ss.items.length === 0))) {
          isStaticSectionEffectivelyEmpty = true;
        } else if (!staticSectionData.items && !staticSectionData.subsections) { // If section has neither items nor subsections arrays
            isStaticSectionEffectivelyEmpty = true;
        }
        
        if (!isStaticSectionEffectivelyEmpty) {
          Logger.log(`  Adding static section to final resume: ${sectionTitle}`);
          finalTailoredResumeObject.sections.push(JSON.parse(JSON.stringify(staticSectionData)));
        } else {
          // Logger.log_DEBUG(`  Skipping empty static section: ${sectionTitle}`);
        }
      } else {
         // Logger.log_DEBUG(`WARN (Stage 3): Static section "${sectionTitle}" (from desired order) not found in master.`);
      }
    }
  });
  
  // Sort all sections into the desired final order
  finalTailoredResumeObject.sections.sort((a, b) => {
    // Handle cases where a title might not be in desiredFinalSectionOrder (e.g., if one was misspelled in master)
    const indexA = desiredFinalSectionOrder.indexOf(a.title.toUpperCase());
    const indexB = desiredFinalSectionOrder.indexOf(b.title.toUpperCase());
    if (indexA === -1 && indexB === -1) return 0; // both not in order, keep relative
    if (indexA === -1) return 1; // a is not in order, push to end
    if (indexB === -1) return -1; // b is not in order, push to end
    return indexA - indexB;
  });

  Logger.log(`STAGE 3: Final 'finalTailoredResumeObject' assembled with ${finalTailoredResumeObject.sections.length} sections.`);
  // Logger.log_DEBUG(`Final Tailored Object Preview: ${JSON.stringify(finalTailoredResumeObject).substring(0,500)}`);
  
  // --- Generate Google Document (using the ORIGINAL createFormattedResumeDoc for now) ---
  Logger.log("STAGE 3 - Calling createFormattedResumeDoc to generate Google Document...");
  let documentGeneratedTitle = `Tailored Resume - ${personalInfoCopy.fullName}`;
  if (jdAnalysis && jdAnalysis.jobTitle) {
    const cleanJobTitle = jdAnalysis.jobTitle.replace(/[^a-zA-Z0-9\s-]/g, "").substring(0,30);
    documentGeneratedTitle += ` for ${cleanJobTitle}`;
  }
  
  // Calls your original document creation function (not the enhanced one with API tabs yet)
  const documentUrl = createFormattedResumeDoc(finalTailoredResumeObject, documentGeneratedTitle); 

  if (documentUrl) {
    Logger.log(`STAGE 3: SUCCESS! Document created: ${documentUrl}`);
    return { 
      success: true, 
      message: "Stage 3 Complete: Tailored resume document generated using refined selection logic.", 
      docUrl: documentUrl, 
      tailoredResumeObjectForDebug: finalTailoredResumeObject 
    };
  } else {
    Logger.log("ERROR (Stage 3): Document generation failed (createFormattedResumeDoc returned null).");
    return { 
      success: false, 
      message: "Stage 3 Failed: Document generation did not return a URL.", 
      tailoredResumeObjectForDebug: finalTailoredResumeObject 
    };
  }
}
// --- END OF STAGE 3 FUNCTION ---
