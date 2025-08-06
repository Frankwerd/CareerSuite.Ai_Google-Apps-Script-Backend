// TailoringLogic.gs
// This file contains functions responsible for constructing prompts and making calls
// to the LLM (via GroqService.gs) for specific resume tailoring tasks such as
// Job Description analysis, bullet point relevance matching, bullet point tailoring,
// and professional summary generation.

// The GROQ_MODEL_NAME constant is now expected to be defined in 'Constants.gs'.
// Example from Constants.gs:
// const GROQ_MODEL_NAME = "gemma2-9b-it";

/**
 * Analyzes a job description using the configured LLM (Groq) to extract key information.
 *
 * @param {string} jobDescriptionText The full text of the job description.
 * @return {Object} A JavaScript object with the extracted JD information (e.g., jobTitle, keyResponsibilities, etc.),
 *                  or an error object { error: string, rawOutput?: string } on failure.
 */
function analyzeJobDescription(jobDescriptionText) {
  if (!jobDescriptionText || !jobDescriptionText.trim()) {
    Logger.log("ERROR (analyzeJobDescription): Input job description text is empty.");
    return { error: "Input job description text is empty." };
  }
  // Uses GROQ_MODEL_NAME from Constants.gs for the model.
  // If you need to specify different models per task, you can pass it directly to callGroq.
  const modelToUse = GROQ_MODEL_NAME;

  const prompt = `
    Analyze the following job description text.
    Extract the specified information and return it ONLY as a single, valid JSON object.
    Do not include any explanatory text, markdown, or anything outside the JSON structure.
    The JSON object should have keys: "jobTitle", "companyName", "location", "keyResponsibilities", "requiredTechnicalSkills", "requiredSoftSkills", "experienceLevel", "educationRequirements", "primaryKeywords", "companyCultureClues".
    If info for a key isn't present in the JD, use "" for string values or [] for array values. Maintain all keys.

    Job Description Text:
    ---
    ${jobDescriptionText}
    ---

    Output JSON:
  `;

  Logger.log(`analyzeJobDescription: Calling Groq (model: ${modelToUse}) to analyze JD.`);
  const groqResponseString = callGroq(prompt, modelToUse, 0.1, 1500); // Low temperature for structured JSON output

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    try {
      // Clean potential markdown formatting from LLM response
      let cleanedResponse = groqResponseString.trim();
      if (cleanedResponse.startsWith("```json")) cleanedResponse = cleanedResponse.substring(7).trim();
      else if (cleanedResponse.startsWith("```")) cleanedResponse = cleanedResponse.substring(3).trim();
      if (cleanedResponse.endsWith("```")) cleanedResponse = cleanedResponse.substring(0, cleanedResponse.length - 3).trim();
      
      const parsedJD = JSON.parse(cleanedResponse);
      // Basic validation of the parsed JSON structure
      if (parsedJD && typeof parsedJD.jobTitle !== 'undefined' && Array.isArray(parsedJD.keyResponsibilities)) {
        Logger.log("analyzeJobDescription: Successfully parsed JD analysis from Groq.");
        return parsedJD;
      } else {
        Logger.log(`ERROR (analyzeJobDescription): Parsed JSON from Groq has an unexpected structure or is missing required fields (jobTitle, keyResponsibilities). Raw Output: ${cleanedResponse}`);
        return { error: "Groq output structure mismatch for JD analysis.", rawOutput: cleanedResponse };
      }
    } catch (e) {
      Logger.log(`ERROR (analyzeJobDescription): Failed to parse JSON from Groq. Error: ${e.toString()}. Raw Groq Output: ${groqResponseString}`);
      return { error: "Failed to parse Groq JSON output for JD analysis.", rawOutput: groqResponseString };
    }
  } else {
    Logger.log(`ERROR (analyzeJobDescription): Groq API call failed or returned an error string. Response: ${groqResponseString}`);
    return { error: "Groq API call failed for JD analysis.", rawOutput: groqResponseString };
  }
}


/**
 * Asks the LLM (Groq) to evaluate the relevance of a resume section/bullet point
 * to the analyzed job description.
 *
 * @param {string} resumeSectionText The text from a resume section (e.g., a single bullet point).
 * @param {Object} jdAnalysisResults The structured analysis of the job description from `analyzeJobDescription`.
 * @return {Object} An object with { relevanceScore: number, matchingKeywords: string[], justification: string },
 *                  or an error object { error: string, rawOutput?: string } on failure.
 */
function matchResumeSection(resumeSectionText, jdAnalysisResults) {
  if (!resumeSectionText || !resumeSectionText.trim()) {
    Logger.log("ERROR (matchResumeSection): Input resume section text is empty.");
    return { error: "Input resume section text is empty." };
  }
  if (!jdAnalysisResults || !jdAnalysisResults.primaryKeywords || !jdAnalysisResults.jobTitle) { // Added jobTitle check for robust context
    Logger.log("ERROR (matchResumeSection): Invalid or incomplete JD analysis results provided.");
    return { error: "Invalid JD analysis results (missing primaryKeywords or jobTitle)." };
  }
  
  const modelToUse = GROQ_MODEL_NAME; // Uses constant from Constants.gs
  const jdAnalysisContextString = JSON.stringify(jdAnalysisResults, null, 2); // For prompt context

  const prompt = `
    You are an expert resume analyst. Evaluate the relevance of the "Resume Section Text" 
    to the "Analyzed Job Description Data" provided below. Focus on how well the resume text
    demonstrates skills, experiences, and keywords mentioned or implied in the job description.
    
    Return ONLY a single, valid JSON object with the following structure:
    {"relevanceScore": number (a float between 0.0 for no relevance and 1.0 for perfect relevance), "matchingKeywords": array of strings (list any keywords from the JD that directly match or are strongly represented in the resume text), "justification": string (a brief 1-2 sentence explanation for your score, highlighting key matches or mismatches)}.
    
    If the resume text is completely irrelevant, assign a relevanceScore of 0.0 and provide a clear justification.

    --- Analyzed Job Description Data ---
    ${jdAnalysisContextString}

    --- Resume Section Text to Evaluate ---
    ${resumeSectionText}

    --- Output JSON ---
  `;

  Logger.log(`matchResumeSection: Calling Groq (model: ${modelToUse}) to evaluate relevance.`);
  const groqResponseString = callGroq(prompt, modelToUse, 0.2, 512); // Low temperature, expecting structured JSON

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    try {
      let cleanedResponse = groqResponseString.trim();
      if (cleanedResponse.startsWith("```json")) cleanedResponse = cleanedResponse.substring(7).trim();
      else if (cleanedResponse.startsWith("```")) cleanedResponse = cleanedResponse.substring(3).trim();
      if (cleanedResponse.endsWith("```")) cleanedResponse = cleanedResponse.substring(0, cleanedResponse.length - 3).trim();
      
      const parsedMatch = JSON.parse(cleanedResponse);
      if (parsedMatch && typeof parsedMatch.relevanceScore === 'number' && 
          Array.isArray(parsedMatch.matchingKeywords) && typeof parsedMatch.justification === 'string') {
        // Logger.log_DEBUG(`matchResumeSection: Groq relevance parsed successfully. Score: ${parsedMatch.relevanceScore}`);
        return parsedMatch;
      } else {
        Logger.log(`ERROR (matchResumeSection): Parsed JSON from Groq has an unexpected structure for matching. Raw Output: ${cleanedResponse}`);
        return { error: "Groq output structure mismatch for matching.", rawOutput: cleanedResponse };
      }
    } catch (e) {
      Logger.log(`ERROR (matchResumeSection): Failed to parse JSON from Groq for matching. Error: ${e.toString()}. Raw Groq Output: ${groqResponseString}`);
      return { error: "Failed to parse Groq JSON for matching.", rawOutput: groqResponseString };
    }
  } else {
    Logger.log(`ERROR (matchResumeSection): Groq API call failed or returned an error string for matching. Response: ${groqResponseString}`);
    return { error: "Groq API call failed for matching.", rawOutput: groqResponseString };
  }
}

/**
 * Asks the LLM (Groq) to rewrite/tailor an original resume bullet point to better
 * align with the target job description and role.
 *
 * @param {string} originalBullet The original bullet point text.
 * @param {Object} jdAnalysisResults The structured analysis of the job description.
 * @param {string} targetRoleTitle The specific job title being targeted.
 * @return {string} The tailored bullet point string, the specific message "Original bullet not suitable...", 
 *                  or an "ERROR: ..." string on failure.
 */
function tailorBulletPoint(originalBullet, jdAnalysisResults, targetRoleTitle) {
  if (!originalBullet || !originalBullet.trim()) {
    Logger.log("ERROR (tailorBulletPoint): Original bullet text is empty.");
    return "ERROR: Original bullet empty.";
  }
  if (!jdAnalysisResults || !jdAnalysisResults.primaryKeywords || !jdAnalysisResults.jobTitle) {
    Logger.log("ERROR (tailorBulletPoint): Invalid or incomplete JD analysis results provided.");
    return "ERROR: Invalid JD analysis.";
  }
  if (!targetRoleTitle || !targetRoleTitle.trim()) {
    Logger.log("ERROR (tailorBulletPoint): Target role title is empty.");
    return "ERROR: Target role empty.";
  }

  const modelToUse = GROQ_MODEL_NAME; // Uses constant from Constants.gs
  // Construct a concise context string from JD analysis for the prompt
  const jdContextString = `
Target Role: ${targetRoleTitle}
Key Responsibilities (sample): ${(jdAnalysisResults.keyResponsibilities || []).slice(0, 3).join('; ')}
Primary Keywords from JD: ${(jdAnalysisResults.primaryKeywords || []).join(', ')}
Required Technical Skills from JD: ${(jdAnalysisResults.requiredTechnicalSkills || []).join(', ')}
  `.trim();

  const prompt = `
    You are an expert resume writer. Your task is to rewrite the "Original Resume Bullet" provided below 
    to be highly effective for a candidate applying for the "Target Role" detailed in the "Job Description Context".

    Follow these guidelines strictly:
    1. Start the rewritten bullet with a strong action verb.
    2. Quantify achievements using metrics (numbers, percentages, etc.) whenever possible. If the original bullet had metrics, retain or enhance them. Do not invent metrics if they are not present or reasonably inferable.
    3. Naturally incorporate relevant skills or keywords from the "Job Description Context" IF AND ONLY IF they genuinely fit the context of the original bullet's achievement. Do NOT force keywords where they don't belong.
    4. Maintain the core truthfulness of the original bullet. Do not fabricate experiences or outcomes.
    5. Aim for a concise, professional, and impactful bullet point, ideally 1-2 lines long.
    6. If the original bullet is already excellent and highly relevant as-is, you may return it with minor enhancements or return it unchanged.
    7. If the original bullet is completely irrelevant to the target role and cannot be reasonably or truthfully tailored, you MUST return the exact phrase: "Original bullet not suitable for significant tailoring towards this role."

    --- Job Description Context ---
    ${jdContextString} 

    --- Original Resume Bullet to Tailor ---
    ${originalBullet}

    --- Rewritten (Tailored) Resume Bullet ---
    (Important: Return ONLY the rewritten bullet text itself OR the specific "not suitable" message. 
    You MAY optionally return the rewritten bullet as a JSON object: {"rewritten_bullet": "Your tailored bullet text here."} 
    If you use the JSON format, ensure it's the ONLY thing in your response.)
  `;

  Logger.log(`tailorBulletPoint: Calling Groq (model: ${modelToUse}) to tailor bullet.`);
  // Logger.log_DEBUG(`  tailorBulletPoint prompt preview: ${prompt.substring(0, 300)}...`); // For debugging prompt
  const groqResponseString = callGroq(prompt, modelToUse, 0.5, 256); // Temp for creativity, max tokens for a bullet

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    // Logger.log_DEBUG(`tailorBulletPoint: Groq raw response: ${groqResponseString}`); // Can be very verbose
    let tailoredText = groqResponseString.trim();
    try {
      // Attempt to parse if LLM opted for the JSON structure
      let potentialJson = tailoredText;
      if (potentialJson.startsWith("```json")) potentialJson = potentialJson.substring(7).trim();
      else if (potentialJson.startsWith("```")) potentialJson = potentialJson.substring(3).trim();
      if (potentialJson.endsWith("```")) potentialJson = potentialJson.substring(0, potentialJson.length - 3).trim();
      
      const parsedResponse = JSON.parse(potentialJson);
      if (parsedResponse && parsedResponse.rewritten_bullet && typeof parsedResponse.rewritten_bullet === 'string') {
        Logger.log("tailorBulletPoint: Successfully extracted tailored bullet from Groq JSON response.");
        return parsedResponse.rewritten_bullet.trim();
      } else {
        // If not the specific JSON, or missing key, assume the whole (cleaned) response is the direct bullet text or "not suitable" message.
        Logger.log("tailorBulletPoint: Groq response was not valid JSON with 'rewritten_bullet', using the cleaned raw string output.");
        return tailoredText; // Return the cleaned original response
      }
    } catch (e) {
      // Not JSON, assume it's the direct string output (either tailored bullet or "not suitable" message)
      Logger.log("tailorBulletPoint: Groq response was not JSON, using as direct string. (This is expected if LLM returns plain text).");
      return tailoredText; // Return the cleaned original response
    }
  } else {
    Logger.log(`ERROR (tailorBulletPoint): Groq API call failed or returned an error string for tailoring. Response: ${groqResponseString}`);
    return `ERROR: Groq API call failed during tailoring for bullet: "${originalBullet.substring(0,50)}..."`;
  }
}


/**
 * Generates a tailored professional summary using the LLM (Groq).
 *
 * @param {string} topMatchedExperiencesHighlights A string containing key highlights/achievements from the resume.
 * @param {Object} jdAnalysisResults The structured analysis of the job description.
 * @param {string} candidateFullName The full name of the candidate.
 * @return {string} The tailored professional summary string, or an "ERROR: ..." string on failure.
 */
function generateTailoredSummary(topMatchedExperiencesHighlights, jdAnalysisResults, candidateFullName) {
  if (!topMatchedExperiencesHighlights || !topMatchedExperiencesHighlights.trim()) {
    Logger.log("ERROR (generateTailoredSummary): Input 'topMatchedExperiencesHighlights' is empty.");
    return "ERROR: Highlights empty.";
  }
  if (!jdAnalysisResults || !jdAnalysisResults.jobTitle) {
    Logger.log("ERROR (generateTailoredSummary): Invalid or incomplete JD analysis results provided (missing jobTitle).");
    return "ERROR: Invalid JD analysis.";
  }
  if (!candidateFullName || !candidateFullName.trim()) {
    Logger.log("ERROR (generateTailoredSummary): Candidate's full name is empty.");
    return "ERROR: Candidate name empty.";
  }

  const modelToUse = GROQ_MODEL_NAME; // Uses constant from Constants.gs
  const jdContextString = `
Target Role: ${jdAnalysisResults.jobTitle}${jdAnalysisResults.companyName ? ' at ' + jdAnalysisResults.companyName : ''}.
Key Responsibilities (sample): ${(jdAnalysisResults.keyResponsibilities || []).slice(0, 2).join('; ')}.
Primary Keywords from JD: ${(jdAnalysisResults.primaryKeywords || []).join(', ')}.
Required Technical Skills from JD: ${(jdAnalysisResults.requiredTechnicalSkills || []).join(', ')}.
Experience Level Sought: ${jdAnalysisResults.experienceLevel || 'Not specified'}.
  `.trim();

  const prompt = `
    You are an expert resume writer. Your task is to craft a compelling professional summary 
    for a candidate named ${candidateFullName}.
    The candidate is applying for the "Target Role" as described in the "Job Description Context".
    The candidate's key relevant strengths and achievements, derived from their resume, are provided in "Candidate Highlights".

    Instructions for the summary:
    1. Write a concise, impactful, and professional summary. Aim for 2-4 sentences, with a maximum of 80 words.
    2. Directly address the "Target Role".
    3. Naturally integrate relevant keywords and phrases from the "Job Description Context".
    4. Leverage the "Candidate Highlights" to showcase specific skills and achievements.
    5. The tone should be confident and compelling.
    6. Avoid clichés and overly generic statements (e.g., "results-oriented professional").
    7. STRICTLY AVOID using first-person pronouns (I, me, my). Write in the third person or impersonally.

    --- Job Description Context ---
    ${jdContextString}

    --- Candidate Highlights (Key skills/achievements from their resume relevant to this role) ---
    ${topMatchedExperiencesHighlights}

    --- Professional Summary ---
    (Important: Return ONLY the summary text. 
    You MAY optionally return it as a JSON object: {"professionalSummary": "Your summary text here."}
    If you use the JSON format, ensure it's the ONLY thing in your response.)
  `;

  Logger.log(`generateTailoredSummary: Calling Groq (model: ${modelToUse}) to generate summary.`);
  const groqResponseString = callGroq(prompt, modelToUse, 0.6, 200); // Higher temp for creativity, reasonable token limit for summary

  if (groqResponseString && !groqResponseString.startsWith("ERROR:")) {
    // Logger.log_DEBUG(`generateTailoredSummary: Groq raw response: ${groqResponseString}`);
    let summaryText = groqResponseString.trim();
    try {
      // Attempt to parse if LLM opted for the JSON structure
      let potentialJson = summaryText;
      if (potentialJson.startsWith("```json")) potentialJson = potentialJson.substring(7).trim();
      else if (potentialJson.startsWith("```")) potentialJson = potentialJson.substring(3).trim();
      if (potentialJson.endsWith("```")) potentialJson = potentialJson.substring(0, potentialJson.length - 3).trim();
      
      const parsedResponse = JSON.parse(potentialJson);
      if (parsedResponse && parsedResponse.professionalSummary && typeof parsedResponse.professionalSummary === 'string') {
        Logger.log("generateTailoredSummary: Successfully extracted summary from Groq JSON response.");
        return parsedResponse.professionalSummary.trim();
      } else {
        Logger.log("generateTailoredSummary: Groq response was not valid JSON with 'professionalSummary', using the cleaned raw string output.");
        return summaryText; // Return the cleaned original response
      }
    } catch (e) {
      Logger.log("generateTailoredSummary: Groq response was not JSON, using as direct string. (This is expected if LLM returns plain text).");
      return summaryText; // Return the cleaned original response
    }
  } else {
    Logger.log(`ERROR (generateTailoredSummary): Groq API call failed or returned an error string for summary generation. Response: ${groqResponseString}`);
    return `ERROR: Groq API call failed during summary generation.`;
  }
}
