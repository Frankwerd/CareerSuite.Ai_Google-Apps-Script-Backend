// File: AdminUtils.gs
// Description: Contains administrative utility functions for project setup and configuration,
// such as managing API keys stored in UserProperties.

// In AdminUtils.gs

/**
 * Provides a UI prompt to set the shared Gemini API Key in UserProperties.
 * Uses the GEMINI_API_KEY_PROPERTY constant from Config.gs.
 */
function setSharedGeminiApiKey_UI() { // Renamed for clarity
  let ui;
  try {
    // Try to get UI from SpreadsheetApp first, then fallback to Browser.inputBox for editor context
    ui = SpreadsheetApp.getUi();
    const currentKey = PropertiesService.getUserProperties().getProperty(GEMINI_API_KEY_PROPERTY); // From Config.gs
    const promptMessage = `Enter the shared Google AI Gemini API Key for all modules.\nThis will be stored in UserProperties under the key: "${GEMINI_API_KEY_PROPERTY}".\n${currentKey ? '(This will overwrite an existing key)' : '(No key currently set)'}`;
    const response = ui.prompt('Set Shared Gemini API Key', promptMessage, ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
      const apiKey = response.getResponseText().trim();
      if (apiKey && apiKey.startsWith("AIza")) { // Basic validation
        PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey); // From Config.gs
        ui.alert('API Key Saved', `The Gemini API Key has been saved successfully for the project under property: "${GEMINI_API_KEY_PROPERTY}".`);
      } else if (apiKey) {
        ui.alert('API Key Not Saved', 'The entered key does not appear to be a valid Gemini API key (should start with "AIza"). Please try again.', ui.ButtonSet.OK);
      } else {
        ui.alert('API Key Not Saved', 'No API key was entered.', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('API Key Setup Cancelled', 'The API key setup process was cancelled.', ui.ButtonSet.OK);
    }
  } catch (e) {
    // Fallback if SpreadsheetApp.getUi() fails (e.g., run purely from script editor without sheet open)
    Logger.log('setSharedGeminiApiKey_UI: Spreadsheet UI context error: ' + e.message + ". Attempting Browser.inputBox.");
    try {
        const currentKeyInfo = PropertiesService.getUserProperties().getProperty(GEMINI_API_KEY_PROPERTY) ? "(An existing key will be overwritten)" : "(No key currently set)";
        const apiKey = Browser.inputBox(`Set Shared Gemini API Key`, `Enter the shared Gemini API Key for all modules. ${currentKeyInfo} (Stored as property: ${GEMINI_API_KEY_PROPERTY})`, Browser.Buttons.OK_CANCEL);
        if (apiKey !== 'cancel' && apiKey !== null) {
            if (apiKey.trim() && apiKey.trim().startsWith("AIza")) {
                PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey.trim());
                Browser.msgBox('API Key Saved', `The Gemini API Key has been saved successfully under property: "${GEMINI_API_KEY_PROPERTY}".`, Browser.Buttons.OK);
            } else if (apiKey.trim()){
                Browser.msgBox('API Key Not Saved', 'The entered key does not appear to be a valid Gemini API key (should start with "AIza"). Please try again.', Browser.Buttons.OK);
            } else {
                Browser.msgBox('API Key Not Saved', 'No API key was entered.', Browser.Buttons.OK);
            }
        } else {
            Browser.msgBox('API Key Setup Cancelled', 'The API key setup process was cancelled.', Browser.Buttons.OK);
        }
    } catch (e2) {
        Logger.log('setSharedGeminiApiKey_UI: Browser.inputBox also failed: ' + e2.message);
        // No UI possible, just log.
    }
  }
}

/**
 * TEMPORARY: Manually sets the shared Gemini API Key in UserProperties.
 * Edit YOUR_GEMINI_KEY_HERE in the code before running.
 * REMOVE OR CLEAR THE KEY FROM CODE AFTER RUNNING FOR SECURITY.
 */
function TEMPORARY_manualSetSharedGeminiApiKey() { // Renamed
  const YOUR_GEMINI_KEY_HERE = 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; // <<< EDIT THIS LINE WITH YOUR KEY
  const propertyName = GEMINI_API_KEY_PROPERTY; // From Config.gs

  if (YOUR_GEMINI_KEY_HERE === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || YOUR_GEMINI_KEY_HERE.trim() === '') {
    const msg = `ERROR: Gemini API Key not set in TEMPORARY_manualSetSharedGeminiApiKey function. Edit the script code first with your key (variable YOUR_GEMINI_KEY_HERE). It should be stored under UserProperty: "${propertyName}".`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert('Action Required', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    catch(e) { try { Browser.msgBox('Action Required', msg, Browser.Buttons.OK); } catch(e2) {} }
    return;
  }
  PropertiesService.getUserProperties().setProperty(propertyName, YOUR_GEMINI_KEY_HERE);
  const successMsg = `UserProperty "${propertyName}" has been MANUALLY SET with the hardcoded Gemini API Key. IMPORTANT: For security, now remove or comment out the TEMPORARY_manualSetSharedGeminiApiKey function, or at least clear the YOUR_GEMINI_KEY_HERE variable in the code.`;
  Logger.log(successMsg);
  try { SpreadsheetApp.getUi().alert('API Key Manually Set', successMsg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { try { Browser.msgBox('API Key Manually Set', successMsg, Browser.Buttons.OK); } catch(e2) {} }
}

/**
 * Displays all UserProperties set for this script project to the logs.
 * Sensitive values like API keys are partially masked.
 */
function showAllUserProperties() { // Renamed
  const userProps = PropertiesService.getUserProperties().getProperties();
  let logOutput = "Current UserProperties for this script project:\n";
  if (Object.keys(userProps).length === 0) {
    logOutput += "  (No UserProperties are currently set for this project)\n";
  } else {
    for (const key in userProps) {
      let value = userProps[key];
      // Mask sensitive values - adjust keywords if needed
      if (key.toLowerCase().includes('api') || key.toLowerCase().includes('key') || key.toLowerCase().includes('secret')) {
        if (value && typeof value === 'string' && value.length > 10) {
          value = value.substring(0, 4) + "..." + value.substring(value.length - 4);
        } else if (value && typeof value === 'string') {
            value = "**** (short value)";
        }
      }
      logOutput += `  ${key}: ${value}\n`;
    }
  }
  Logger.log(logOutput);
  const alertMsg = "Current UserProperties have been logged. Check Apps Script logs (View > Logs or Executions) to see them. Sensitive values are partially masked in the log.";
  try { SpreadsheetApp.getUi().alert("User Properties Logged", alertMsg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { try { Browser.msgBox("User Properties Logged", alertMsg, Browser.Buttons.OK); } catch(e2) {} }
}
