// GeminiService.gs
// This file provides functions for interacting with Google's Gemini API.
// It handles API key retrieval from PropertiesService and makes HTTP requests
// to the Gemini API endpoint.

/** 
 * @const {string} The name of the script property used to store the Gemini API Key.
 * This constant is specific to this service for fetching its API key.
 */
const GEMINI_API_KEY_PROPERTY_NAME = 'GEMINI_API_KEY';

// Assumes a constant like GEMINI_DEFAULT_MODEL is defined in 'Constants.gs' if you want a configurable default.
// Example in Constants.gs:
// const GEMINI_DEFAULT_MODEL = "gemini-1.5-flash-latest";

/**
 * Calls the Gemini API with the provided prompt text and model.
 *
 * @param {string} promptText The complete prompt to send to the Gemini API.
 * @param {string} [modelName=GEMINI_DEFAULT_MODEL] Optional. The specific Gemini model to use.
 *        Defaults to the value of GEMINI_DEFAULT_MODEL from Constants.gs (e.g., "gemini-1.5-flash-latest").
 * @return {string|null} The text content of the Gemini API response on success,
 *                       or an "ERROR: ..." string on failure or if the API key is missing.
 */
function callGemini(promptText, modelName = (typeof GEMINI_DEFAULT_MODEL !== 'undefined' ? GEMINI_DEFAULT_MODEL : "gemini-1.5-flash-latest")) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY_NAME);

  if (!apiKey) {
    Logger.log(`ERROR (callGemini): Gemini API Key not found. Please set property: '${GEMINI_API_KEY_PROPERTY_NAME}'`);
    // Potentially show a UI alert if called from a user-facing function
    // try { SpreadsheetApp.getUi().alert("API Key Error", "Gemini API Key is missing."); } catch(e) {}
    return "ERROR: Gemini API Key Missing";
  }

  const effectiveModelName = modelName; // Use the provided or defaulted modelName
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${effectiveModelName}:generateContent?key=${apiKey}`;
  
  const payload = {
    "contents": [{"parts": [{"text": promptText}]}],
    "generationConfig": {
      "temperature": 0.2,       // Lower for more deterministic, factual output
      "maxOutputTokens": 4096,  // Max tokens the model can generate for this request
      "topP": 0.95,             // Nucleus sampling parameter
      "topK": 40,               // Top-k sampling parameter
      "responseMimeType": "application/json" // Requesting Gemini to return JSON
    },
    "safetySettings": [ // Standard safety settings to block harmful content
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // Allows handling of non-200 responses in code
  };

  Logger.log(`callGemini: Calling Gemini API (Model: ${effectiveModelName}). Prompt length: ${promptText.length}`);
  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
        const generatedText = jsonResponse.candidates[0].content.parts[0].text;
        Logger.log(`callGemini: Success from Gemini. Output length: ${generatedText.length}`);
        return generatedText;
      } else if (jsonResponse.promptFeedback?.blockReason) {
        Logger.log(`ERROR (callGemini): Prompt blocked by Gemini. Reason: ${jsonResponse.promptFeedback.blockReason}. Details: ${JSON.stringify(jsonResponse.promptFeedback.safetyRatings)}`);
        return `ERROR: Prompt Blocked by Gemini - ${jsonResponse.promptFeedback.blockReason}`;
      } else {
        Logger.log(`ERROR (callGemini): Gemini response structure unexpected (HTTP 200). Full Response: ${responseBody.substring(0, 1000)}...`);
        return "ERROR: Unexpected Gemini API Response Structure";
      }
    } else {
      Logger.log(`ERROR (callGemini): Gemini API call failed. HTTP Code: ${responseCode}. Body: ${responseBody.substring(0, 1000)}...`);
      return `ERROR: Gemini API Call Failed - HTTP ${responseCode}`;
    }
  } catch (e) {
    Logger.log(`EXCEPTION (callGemini): during Gemini API call: ${e.toString()}\nStack: ${e.stack}`);
    return `ERROR: Exception during Gemini API call - ${e.message}`;
  }
}

/**
 * UI-triggered function to allow users to set the Gemini API Key.
 * The key is stored in Script Properties. This should be run once from the Apps Script editor.
 */
function SET_GEMINI_API_KEY_UI() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi(); // Or DocumentApp.getUi(), FormApp.getUi() depending on context
  } catch (e) {
    Logger.log("SET_GEMINI_API_KEY_UI: Could not get UI service. This function must be run from an environment with UI support (e.g., bound script editor).");
    return;
  }
  
  const result = ui.prompt(
      'Set Gemini API Key',
      'Please enter your Gemini API Key (leave blank to clear):', // Added clear instruction
      ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText().trim(); // Trim whitespace
    const properties = PropertiesService.getScriptProperties();
    if (apiKey !== "") {
      properties.setProperty(GEMINI_API_KEY_PROPERTY_NAME, apiKey);
      ui.alert('API Key Set', `Gemini API Key has been stored for this script project.`);
      Logger.log(`Gemini API Key updated via UI.`);
    } else {
      properties.deleteProperty(GEMINI_API_KEY_PROPERTY_NAME); // Allow clearing the key
      ui.alert('API Key Cleared', 'Gemini API Key has been cleared from script properties.');
      Logger.log(`Gemini API Key cleared via UI.`);
    }
  } else {
    ui.alert('Action Cancelled', 'Gemini API Key was not changed.');
  }
}
