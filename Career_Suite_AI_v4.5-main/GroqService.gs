// GroqService.gs
// This file provides functions for interacting with the Groq API.
// It handles API key retrieval from PropertiesService and makes HTTP requests
// to the Groq API endpoint (OpenAI-compatible).

/** 
 * @const {string} The name of the script property used to store the Groq API Key.
 * This is specific to this service for fetching its API key.
 */
const GROQ_API_KEY_PROPERTY_NAME = 'GROQ_API_KEY';

// Assumes a constant like GROQ_DEFAULT_MODEL_NAME is defined in 'Constants.gs'.
// Example from Constants.gs (ensure it matches the name you used, e.g., GROQ_MODEL_NAME):
// const GROQ_DEFAULT_MODEL_NAME = "gemma2-9b-it"; 

// You could also define a default system prompt in Constants.gs if desired:
// const GROQ_DEFAULT_SYSTEM_PROMPT = "You are a helpful AI assistant...";


/**
 * Calls the Groq API with the provided prompt text using an OpenAI-compatible chat completions endpoint.
 *
 * @param {string} promptText The user's content/message for the prompt.
 * @param {string} [modelName=(GROQ_DEFAULT_MODEL_NAME || "gemma2-9b-it")] The specific Groq model to use. 
 *        Defaults to GROQ_DEFAULT_MODEL_NAME from Constants.gs, or "gemma2-9b-it" if not defined.
 * @param {number} [temperature=0.2] Optional. The temperature for generation (0.0 to 2.0).
 * @param {number} [maxTokens=2048] Optional. Maximum number of tokens for the completion.
 * @param {string} [systemContent="You are a helpful..."] Optional. A system message to guide the assistant.
 * @return {string|null} The text content of the Groq API response on success, 
 *                       or an "ERROR: ..." string on failure or if API key is missing.
 */
function callGroq(
  promptText,
  modelName = (typeof GROQ_DEFAULT_MODEL_NAME !== 'undefined' ? GROQ_DEFAULT_MODEL_NAME : "gemma2-9b-it"), // Use global default
  temperature = 0.2,
  maxTokens = 2048,
  systemContent = "You are a helpful and meticulous AI assistant. Respond ONLY with the requested format (e.g., JSON). Do not include any explanatory text or markdown formatting before or after the JSON output." // Default system prompt can remain here or be moved to Constants.gs
) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty(GROQ_API_KEY_PROPERTY_NAME);

  if (!apiKey) {
    Logger.log(`ERROR (callGroq): Groq API Key not found. Please set property: '${GROQ_API_KEY_PROPERTY_NAME}' in Script Properties.`);
    return `ERROR: Groq API Key Missing`;
  }

  const API_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions";
  
  let messages = [];
  if (systemContent && systemContent.trim() !== "") {
    messages.push({"role": "system", "content": systemContent});
  }
  messages.push({"role": "user", "content": promptText});

  const payload = {
    "messages": messages,
    "model": modelName, 
    "temperature": temperature,
    "max_tokens": maxTokens,
    "top_p": 1,          // Default value, can be adjusted if needed
    "stream": false,     // Must be false for standard Apps Script UrlFetch to get full response
    "stop": null         // No specific stop sequences by default
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // Allows handling HTTP errors in code
  };

  Logger.log(`callGroq: Calling Groq API (Model: ${modelName}, Temp: ${temperature}). User Prompt Len: ${promptText.length}`);
  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.choices && jsonResponse.choices[0]?.message?.content) {
        const generatedText = jsonResponse.choices[0].message.content;
        Logger.log(`callGroq: Success from Groq. Output length: ${generatedText.length}`);
        return generatedText.trim();
      } else if (jsonResponse.error) { 
        Logger.log(`ERROR (callGroq): Groq API (HTTP 200) returned an error object: ${JSON.stringify(jsonResponse.error)}`);
        return `ERROR: Groq API Error - ${jsonResponse.error.message || JSON.stringify(jsonResponse.error)}`;
      } else {
        Logger.log(`ERROR (callGroq): Groq response structure unexpected (HTTP 200). Full Response (first 1000 chars): ${responseBody.substring(0,1000)}`);
        return "ERROR: Unexpected API Response Structure from Groq (HTTP 200)";
      }
    } else { 
      Logger.log(`ERROR (callGroq): Groq API call failed. HTTP Code: ${responseCode}. Body (first 500 chars): ${responseBody.substring(0, 500)}...`);
      let errorMessage = `ERROR: API Call Failed (Groq) - HTTP ${responseCode}`;
      try { 
        const errorJson = JSON.parse(responseBody); 
        if (errorJson.error && errorJson.error.message) {
          errorMessage += `: ${errorJson.error.message}`;
        }
      } catch (e) { /* Ignore parse error of non-JSON error body */ }
      return errorMessage;
    }
  } catch (e) {
    Logger.log(`EXCEPTION (callGroq): during Groq API call: ${e.toString()}\nStack: ${e.stack}`);
    return `ERROR: Exception during Groq API call - ${e.message}`;
  }
}

/**
 * UI-triggered function to allow users to set the Groq API Key.
 * The key is stored in Script Properties. This should be run once from the Apps Script editor.
 */
function SET_GROQ_API_KEY_UI() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi(); // Or DocumentApp.getUi(), etc., depending on where it's run
  } catch (e) {
    Logger.log("SET_GROQ_API_KEY_UI: Could not get UI service. Run from an environment with UI support.");
    return;
  }
  
  const result = ui.prompt(
      'Set Groq API Key',
      'Please enter your Groq API Key (leave blank to clear):',
      ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    const properties = PropertiesService.getScriptProperties();
    if (apiKey !== "") {
      properties.setProperty(GROQ_API_KEY_PROPERTY_NAME, apiKey);
      ui.alert('API Key Set', 'Groq API Key has been stored for this script project.');
      Logger.log("Groq API Key updated via UI.");
    } else {
      properties.deleteProperty(GROQ_API_KEY_PROPERTY_NAME);
      ui.alert('API Key Cleared', 'Groq API Key has been cleared from script properties.');
      Logger.log("Groq API Key cleared via UI.");
    }
  } else {
    ui.alert('Action Cancelled', 'Groq API Key was not changed.');
  }
}
