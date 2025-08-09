// =====================================================================================================================
// --- GLOBAL CONFIGURATION & UI ---
// This file handles fetching configuration from Script Properties and setting up the user interface menus.
// =====================================================================================================================

/**
 * Fetches the user-defined script properties. This is the central place to get configuration.
 * @returns {{course_id: string|null, gemini_api_key: string|null}} An object containing the configuration.
 */
function getConfiguration() {
  const properties = PropertiesService.getScriptProperties();
  return {
    course_id: properties.getProperty('course_id'),
    gemini_api_key: properties.getProperty('gemini_api_key')
  };
}

/**
 * An Apps Script special function that runs automatically when the spreadsheet is opened.
 * It creates the custom menus for all script features.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Create a single, unified menu for all tools.
  ui.createMenu('Classroom Tools')
    .addSubMenu(ui.createMenu('Command Center')
      .addItem('1. Create Student Sheets & Update List', 'createSheetsFromClassroom')
      .addItem('2. Setup Command Center & Populate Scores', 'createCommandCenter')
      .addSeparator()
      .addItem('Setup Rubric on This Sheet (Utility)', 'setupRubric')
      .addSeparator()
      .addItem('Select Active Course', 'selectAndSetCourseId'))
    .addSeparator()
    // The AI sidebar is now named "Commander Assistant" for direct one-click access.
    .addItem('Commander Assistant', 'showAiCopilotSidebar')
    .addToUi();
}

/**
 * Creates and shows the HTML sidebar for the AI Copilot.
 * It first checks if the necessary configuration is in place.
 */
function showAiCopilotSidebar() {
  // Pass 'true' to ensure the Gemini API key is also checked.
  if (!checkConfiguration(true)) return; 
  
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Commander Assistant'); // Also updated the sidebar title for consistency
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Checks if the required configuration (course_id and optionally gemini_api_key) is set.
 * This function prevents other functions from running without the necessary setup.
 * @param {boolean} checkGeminiKey - If true, also checks for the gemini_api_key.
 * @returns {boolean} True if the configuration is valid, false otherwise.
 */
function checkConfiguration(checkGeminiKey = false) {
    const config = getConfiguration();
    const ui = SpreadsheetApp.getUi();

    if (!config.course_id) {
        ui.alert('Configuration Missing', 'No Course ID has been set. Please use "Command Center > Select Active Course" to set it.', ui.ButtonSet.OK);
        return false;
    }

    if (checkGeminiKey && !config.gemini_api_key) {
        // Reverted the alert message to point to the Script Properties.
        ui.alert('Configuration Missing', 'The Gemini API Key is not set. Please add it via Extensions > Apps Script > Project Settings > Script Properties.', ui.ButtonSet.OK);
        return false;
    }
    
    return true;
}
