/**
 * @file Handles all interactions with the Google Gemini API for
 * AI-powered parsing of email content to extract job application details and job leads.
 */

/**
 * Calls the Gemini API to parse job application details from an email.
 * @param {string} emailSubject The subject of the email.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key.
 * @returns {{company: string, title: string, status: string}|null} An object with the parsed details or null on failure.
 */
function callGemini_forApplicationDetails(emailSubject, emailBody, apiKey) {
  if (!apiKey) {
    Logger.log("[INFO] GEMINI_PARSE_APP: API Key not provided. Skipping Gemini call.");
    return null;
  }
  if ((!emailSubject || emailSubject.trim() === "") && (!emailBody || emailBody.trim() === "")) {
    Logger.log("[WARN] GEMINI_PARSE_APP: Both email subject and body are empty. Skipping Gemini call.");
    return null;
  }

  const API_ENDPOINT = GEMINI_API_ENDPOINT_TEXT_ONLY + "?key=" + apiKey;
  if (DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE_APP: Using API Endpoint: ${API_ENDPOINT.split('key=')[0] + "key=..."}`);

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : ""; // Max 12k chars for body snippet

// --- START: Replacement for the prompt in callGemini_forApplicationDetails ---
const prompt = `You are a highly specialized AI assistant expert in parsing job application-related emails for a tracking system. Your sole purpose is to analyze the provided email Subject and Body, and extract three key pieces of information: "company_name", "job_title", and "status". You MUST return this information ONLY as a single, valid JSON object, with no surrounding text, explanations, apologies, or markdown.

CRITICAL INSTRUCTIONS - READ AND FOLLOW CAREFULLY:

1.  "company_name":
    *   Extract the full, official name of the HIRING COMPANY.
    *   Do NOT extract the name of the ATS (e.g., "Greenhouse") or the job board (e.g., "LinkedIn").
    *   If the company name is genuinely unclear, use the exact string "${MANUAL_REVIEW_NEEDED}".

2.  "job_title":
    *   Extract the SPECIFIC job title the user applied for as mentioned in THIS email.
    *   If the job title is not clearly present, use the exact string "${MANUAL_REVIEW_NEEDED}".

3.  "status":
    *   Determine the current status of the application based on the content of THIS email.
    *   You MUST choose a status ONLY from the following exact list. Do not invent new statuses.
        *   "${DEFAULT_STATUS}" (Use for: Application submitted, application sent, successfully applied, application received)
        *   "${REJECTED_STATUS}" (Use for: Not moving forward, unfortunately, decided not to proceed, position filled)
        *   "${OFFER_STATUS}" (Use for: Offer of employment, pleased to offer, job offer)
        *   "${INTERVIEW_STATUS}" (Use for: Invitation to interview, schedule an interview, interview request)
        *   "${ASSESSMENT_STATUS}" (Use for: Online assessment, coding challenge, technical test, skills test)
        *   "${APPLICATION_VIEWED_STATUS}" (Use for: Application was viewed by recruiter, your profile was viewed for the role)
        *   "Update/Other" (Use for: General updates or if the status is unclear)

**Output Requirements**:
*   **ONLY JSON**: Your entire response must be a single, valid JSON object.
*   **Structure**: {"company_name": "...", "job_title": "...", "status": "..."}
*   **Irrelevant Emails**: If the email is clearly NOT a job application update (e.g., a newsletter, a job alert), your output MUST be: {"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "Not an Application"}

--- EMAIL TO PROCESS START ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- EMAIL TO PROCESS END ---

JSON Output:
`;
// --- END: Replacement for the prompt in callGemini_forApplicationDetails ---

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": { "temperature": 0.2, "maxOutputTokens": 2048, "topP": 0.95, "topK": 40 },
    "safetySettings": [ 
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE_APP: Calling API for subj: "${emailSubject.substring(0,100)}". Prompt len (approx): ${prompt.length}`);
  let response; let attempt = 0; const maxAttempts = 2;

  while(attempt < maxAttempts){
    attempt++;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
      const responseCode = response.getResponseCode(); const responseBody = response.getContentText();
      if(DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE_APP (Attempt ${attempt}): RC: ${responseCode}. Body(start): ${responseBody.substring(0,200)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.startsWith("```")) extractedJsonString = extractedJsonString.substring(3).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();
          
          if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE_APP: Cleaned JSON from API: ${extractedJsonString}`);
          try {
            const extractedData = JSON.parse(extractedJsonString);
            if (typeof extractedData.company_name !== 'undefined' && 
                typeof extractedData.job_title !== 'undefined' && 
                typeof extractedData.status !== 'undefined') {
              Logger.log(`[INFO] GEMINI_PARSE_APP: Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
              return {
                  company: extractedData.company_name || MANUAL_REVIEW_NEEDED, 
                  title: extractedData.job_title || MANUAL_REVIEW_NEEDED, 
                  status: extractedData.status || MANUAL_REVIEW_NEEDED
              };
            } else {
              Logger.log(`[WARN] GEMINI_PARSE_APP: JSON from Gemini missing fields. Output: ${extractedJsonString}`);
              return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED};
            }
          } catch (e) {
            Logger.log(`[ERROR] GEMINI_PARSE_APP: Error parsing JSON: ${e.toString()}\nString: >>>${extractedJsonString}<<<`);
            return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED};
          }
        } else {
          Logger.log(`[ERROR] GEMINI_PARSE_APP: API response structure unexpected. Body (start): ${responseBody.substring(0,500)}`);
          return null; 
        }
      } else if (responseCode === 429) {
        Logger.log(`[WARN] GEMINI_PARSE_APP: Rate limit (429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        if (attempt < maxAttempts) { Utilities.sleep(5000 + Math.floor(Math.random() * 5000)); continue; }
        else { Logger.log(`[ERROR] GEMINI_PARSE_APP: Max retries for rate limit.`); return null; }
      } else {
        Logger.log(`[ERROR] GEMINI_PARSE_APP: API HTTP error. Code: ${responseCode}. Body (start): ${responseBody.substring(0,500)}`);
        if (responseCode === 404 && responseBody.includes("is not found for API version")) {
            Logger.log(`[FATAL] GEMINI_MODEL_ERROR_APP: Model ${API_ENDPOINT.split('/models/')[1].split(':')[0]} not found.`)
        }
        return null;
      }
    } catch (e) {
      Logger.log(`[ERROR] GEMINI_PARSE_APP: Exception during API call (Attempt ${attempt}): ${e.toString()}\nStack: ${e.stack}`);
      if (attempt < maxAttempts) { Utilities.sleep(3000); continue; }
      return {error: e.toString()};
    }
  }
  Logger.log(`[ERROR] GEMINI_PARSE_APP: Failed after ${maxAttempts} attempts.`);
  return {error: `Failed after ${maxAttempts} attempts.`};
}


function callGemini_forJobLeads(emailBody, apiKey) {
    if (typeof emailBody !== 'string') {
        Logger.log(`[GEMINI_LEADS CRITICAL ERR] emailBody not string. Type: ${typeof emailBody}`);
        return { success: false, data: null, error: `emailBody is not a string.` };
    }

    const API_ENDPOINT = GEMINI_API_ENDPOINT_TEXT_ONLY + "?key=" + apiKey;

    if (!apiKey || apiKey === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || apiKey.trim() === '') {
        const errorMsg = "API Key is not set. Please set it in the configuration.";
        Logger.log(`[GEMINI_LEADS ERROR] ${errorMsg}`);
        return { success: false, data: null, error: errorMsg };
    }

    // --- MODIFIED PROMPT ---
    const promptText = `You are an expert AI assistant specializing in extracting job posting details from email content, typically from job alerts or direct emails containing job opportunities.
From the following "Email Content", identify each distinct job posting.

For EACH job posting found, extract the following details:
- "jobTitle": The specific title of the job role (e.g., "Senior Software Engineer", "Product Marketing Manager").
- "company": The name of the hiring company.
- "location": The primary location of the job (e.g., "San Francisco, CA", "Remote", "London, UK", "Hybrid - New York").
- "source": If identifiable from the email content, the origin or job board where this posting was listed (e.g., "LinkedIn Job Alert", "Indeed", "Wellfound", "Company Careers Page" if mentioned). If not explicitly stated, use "N/A".
- "jobUrl": A direct URL link to the job application page or a more detailed job description, if present in the email. If no direct link for *this specific job* is found, use "N/A".
- "notes": Briefly extract 2-3 key requirements, responsibilities, or unique aspects mentioned for this specific job if readily available in the email text (e.g., "Requires Python & AWS; 5+ yrs exp", "Focus on B2B SaaS marketing", "Fast-paced startup environment"). Keep notes concise (max 150 characters). If no specific details are easily extractable for this job, use "N/A".

Strict Formatting Instructions:
- Your entire response MUST be a single, valid JSON array.
- Each element of the array MUST be a JSON object representing one job posting.
- Each JSON object MUST have exactly these keys: "jobTitle", "company", "location", "source", "jobUrl", "notes".
- If a specific field for a job is not found or not applicable, its value MUST be the string "N/A".
- If no job postings at all are found in the email content, return an empty JSON array: [].
- Do NOT include any text, explanations, apologies, or markdown (like \`\`\`json\`\`\`) before or after the JSON array.

--- EXAMPLE OUTPUT START (for an email with two jobs) ---
[
  {
    "jobTitle": "Senior Frontend Developer",
    "company": "Innovatech Solutions",
    "location": "Remote (US)",
    "source": "LinkedIn Job Alert",
    "jobUrl": "https://linkedin.com/jobs/view/12345",
    "notes": "React, TypeScript, Agile environment. 5+ years experience. UI/UX focus."
  },
  {
    "jobTitle": "Data Scientist",
    "company": "Alpha Analytics Co.",
    "location": "Boston, MA",
    "source": "Direct Email from Recruiter",
    "jobUrl": "N/A",
    "notes": "Machine learning, Python, SQL. PhD preferred. Early-stage startup."
  }
]
--- EXAMPLE OUTPUT END ---

Email Content:
---
${emailBody.substring(0, 30000)} 
---
JSON Array Output:`; // Max characters for body increased slightly

    const payload = {
        contents: [{ parts: [{ "text": promptText }] }],
        generationConfig: { 
            temperature: 0.2, 
            maxOutputTokens: 8192, // Kept high for potentially multiple listings
            // responseMimeType: "application/json" // Can try adding this if LLM still includes ```json
        },
        safetySettings: [ 
          { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
        ]
    };
    const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    let attempt = 0; const maxAttempts = 2; // Or 3 for more resilience

    Logger.log(`[GEMINI_LEADS INFO] Calling Gemini for leads. Prompt length (approx): ${promptText.length}`);

    while (attempt < maxAttempts) {
        attempt++;
        Logger.log(`[GEMINI_LEADS] Starting attempt ${attempt}/${maxAttempts}...`);
        try {
            const response = UrlFetchApp.fetch(API_ENDPOINT, options);
            const responseCode = response.getResponseCode(); 
            const responseBody = response.getContentText();
            Logger.log(`[GEMINI_LEADS DEBUG Attempt ${attempt}] RC: ${responseCode}. Body (start): ${responseBody.substring(0, 250)}...`);

            if (responseCode === 200) {
                try { 
                    // The response from Gemini might already be the JSON part.
                    // parseGeminiResponse_forJobLeads will handle cleaning ```json
                    return { success: true, data: JSON.parse(responseBody), error: null }; 
                }
                catch (jsonParseError) {
                    Logger.log(`[GEMINI_LEADS API ERROR] Parse Gemini JSON after 200 OK: ${jsonParseError}. Raw Body: ${responseBody}`);
                    return { success: false, data: null, error: `Parse API JSON (200 OK): ${jsonParseError}. Response: ${responseBody.substring(0,500)}` };
                }
            } else if (responseCode === 429 && attempt < maxAttempts) { // Rate limit
                Logger.log(`[GEMINI_LEADS] API returned 429. Sleeping before retry...`);
                Utilities.sleep(3000 + Math.random() * 2000); 
                continue;
            } else { // Other HTTP errors
                Logger.log(`[GEMINI_LEADS API ERROR ${responseCode}] Full error: ${responseBody}`);
                const parsedError = JSON.parse(responseBody); // Try to parse error for details
                if (parsedError && parsedError.error && parsedError.error.message) {
                     return { success: false, data: null, error: `API Error ${responseCode}: ${parsedError.error.message}` };
                }
                return { success: false, data: null, error: `API Error ${responseCode}: ${responseBody.substring(0,500)}` };
            }
        } catch (e) { // Catch UrlFetchApp.fetch exceptions
            Logger.log(`[GEMINI_LEADS] Caught exception on attempt ${attempt}: ${e.message}.`);
            if (attempt < maxAttempts) {
                Logger.log(`[GEMINI_LEADS] Sleeping before retry...`);
                Utilities.sleep(2000 + Math.random()*1000);
                continue;
            }
            return { success: false, data: null, error: `Fetch Error after ${maxAttempts} attempts: ${e.toString()}` };
        }
    }
    Logger.log(`[GEMINI_LEADS ERROR] Exceeded max retries for Gemini API.`);
    return { success: false, data: null, error: `Exceeded max retries (${maxAttempts}) for Gemini API.` };
}

// --- You will ALSO need to update `parseGeminiResponse_forJobLeads` ---
function parseGeminiResponse_forJobLeads(apiResponseData) {
    let jobListings = [];
    const FUNC_NAME = "parseGeminiResponse_forJobLeads";
    try {
        let jsonStringFromLLM = "";
        if (apiResponseData?.candidates?.[0]?.content?.parts?.[0]?.text) {
            jsonStringFromLLM = apiResponseData.candidates[0].content.parts[0].text.trim();
            // Clean potential markdown fences
            if (jsonStringFromLLM.startsWith("```json")) jsonStringFromLLM = jsonStringFromLLM.substring(7, jsonStringFromLLM.endsWith("```") ? jsonStringFromLLM.length - 3 : undefined).trim();
            else if (jsonStringFromLLM.startsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(3, jsonStringFromLLM.endsWith("```") ? jsonStringFromLLM.length - 3 : undefined).trim();
            else if (jsonStringFromLLM.endsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(0, jsonStringFromLLM.length - 3).trim();
        } else {
            Logger.log(`[${FUNC_NAME} WARN] No parsable content string in Gemini response for leads.`);
            if (apiResponseData?.promptFeedback?.blockReason) {
                Logger.log(`[${FUNC_NAME} WARN] Prompt Block Reason: ${apiResponseData.promptFeedback.blockReason}.`);
            }
            return jobListings; // Empty array
        }

        Logger.log(`[${FUNC_NAME} DEBUG] Cleaned JSON string from LLM: ${jsonStringFromLLM.substring(0, 500)}...`);
        try {
            const parsedData = JSON.parse(jsonStringFromLLM);
            if (Array.isArray(parsedData)) {
                parsedData.forEach(job => {
                    if (job && typeof job === 'object' && (job.jobTitle || job.company)) { // Basic validation for a job object
                        jobListings.push({
                            jobTitle: job.jobTitle || "N/A", 
                            company: job.company || "N/A",
                            location: job.location || "N/A", 
                            source: job.source || "N/A",         // <<< NEW
                            jobUrl: job.jobUrl || "N/A",         // Changed from linkToJobPosting
                            notes: job.notes || "N/A"            // <<< NEW
                        });
                    } else { Logger.log(`[${FUNC_NAME} WARN] Skipped invalid item in parsed job listings array: ${JSON.stringify(job)}`); }
                });
            } else if (typeof parsedData === 'object' && parsedData !== null && (parsedData.jobTitle || parsedData.company)) { // Handle case where LLM returns a single object instead of array
                jobListings.push({
                    jobTitle: parsedData.jobTitle || "N/A", 
                    company: parsedData.company || "N/A",
                    location: parsedData.location || "N/A", 
                    source: parsedData.source || "N/A",      // <<< NEW
                    jobUrl: parsedData.jobUrl || "N/A",      // Changed from linkToJobPosting
                    notes: parsedData.notes || "N/A"         // <<< NEW
                });
                Logger.log(`[${FUNC_NAME} WARN] LLM output was a single object, parsed as one job.`);
            } else { 
                Logger.log(`[${FUNC_NAME} WARN] LLM output not a JSON array or a single parsable job object. Output (start): ${jsonStringFromLLM.substring(0, 200)}`); 
            }
        } catch (jsonError) {
            Logger.log(`[${FUNC_NAME} ERROR] Failed to parse JSON string from LLM: ${jsonError}. String (start): ${jsonStringFromLLM.substring(0, 500)}`);
        }
        Logger.log(`[${FUNC_NAME} INFO] Successfully parsed ${jobListings.length} job listings from Gemini response.`);
        return jobListings;
    } catch (e) {
        Logger.log(`[${FUNC_NAME} ERROR] Outer error during parsing Gemini response for leads: ${e.toString()}. API Resp Data (partial): ${JSON.stringify(apiResponseData).substring(0, 300)}`);
        return jobListings; // Return empty array on error
    }
}
