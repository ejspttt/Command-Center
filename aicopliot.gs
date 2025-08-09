// =====================================================================================================================
// --- AI COPILOT CORE LOGIC ---
// This file contains all the logic for interpreting and executing natural language commands via the Gemini API.
// =====================================================================================================================


/**
 * The primary function called from the sidebar UI. It orchestrates the entire process.
 * @param {string} command The natural language command from the user.
 * @return {string} A result string to be displayed in the debug console.
 */
function processCommand(command) {
  Logger.log(`[DEBUG] Received command: "${command}"`);
  if (!command) {
    return "[DEBUG] No command entered.";
  }
  try {
    const structuredData = getStructuredDataFromGemini(command);
    Logger.log(`[DEBUG] Received structured data from Gemini: ${JSON.stringify(structuredData, null, 2)}`);

    // Handle multi-step commands (e.g., "create 3 assignments").
    if (structuredData && Array.isArray(structuredData.apiCalls)) {
      Logger.log(`[DEBUG] Multi-step command detected with ${structuredData.apiCalls.length} steps.`);
      let results = [];
      for (const apiCall of structuredData.apiCalls) {
        const result = executeClassroomAction({ apiCall });
        results.push(result);
      }
      return results.join('\n-------------------\n');
    }
    
    // Handle single-step commands.
    if (structuredData && structuredData.apiCall) {
      Logger.log(`[DEBUG] Single-step command detected.`);
      return executeClassroomAction(structuredData);
    }

    Logger.log("[DEBUG] Could not determine a valid action from the command.");
    return "[DEBUG] Gemini did not return a valid action. Please check the prompt or Gemini response.";

  } catch (e) {
    Logger.log(`[DEBUG] CRITICAL ERROR in processCommand: ${e.toString()}\nStack: ${e.stack}`);
    return `[DEBUG] An error occurred: ${e.message}`;
  }
}

/**
 * Executes a specific Google Classroom API call based on the structured data from Gemini.
 * @param {Object} data A wrapper object containing the 'apiCall' object from Gemini.
 * @return {string} A result message for the debug console.
 */
function executeClassroomAction(data) {
  const config = getConfiguration();
  const { resource, method, pre_flight, summary } = data.apiCall;
  let params = data.apiCall.params;

  try {
    // --- PRE-FLIGHT LOGIC ---
    if (pre_flight) {
      Logger.log(`[DEBUG] Pre-flight action detected. Finding item first.`);
      const { resource: preFlightResource, method: preFlightMethod, params: preFlightParams, filter, select } = pre_flight;
      
      const preFlightResourceParts = preFlightResource.split('.');
      let preFlightApiObject = Classroom;
      for (const part of preFlightResourceParts) {
        preFlightApiObject = preFlightApiObject[part];
      }
      
      const listResult = preFlightApiObject[preFlightMethod](...preFlightParams);
      const resultKey = Object.keys(listResult).find(key => Array.isArray(listResult[key]));
      let items = resultKey ? listResult[resultKey] : [];

      if (!items || items.length === 0) {
        return `[DEBUG] PRE-FLIGHT FAILED: Could not find any items in '${preFlightResource}' to perform the action on.`;
      }
      
      if (filter) {
        items = items.filter(item => 
          item && item[filter.key] && typeof item[filter.key] === 'string' && 
          item[filter.key].toLowerCase().includes(filter.value.toLowerCase())
        );
      }
      
      if (items.length === 0) return `[DEBUG] PRE-FLIGHT FAILED: Could not find an item matching your criteria.`;
      
      const targetItem = items[0]; 
      const selectedValue = targetItem[select];
      
      params = [config.course_id, selectedValue];
      Logger.log(`[DEBUG] Pre-flight successful. Found item ID: ${selectedValue}. New params: ${JSON.stringify(params)}`);
    }

    // --- EXECUTION LOGIC ---
    const resourceParts = resource.split('.');
    let apiObject = Classroom;
    for (const part of resourceParts) {
      apiObject = apiObject[part];
    }

    if (typeof apiObject[method] !== 'function') {
      const errorMessage = `[VALIDATION FAILED] The AI returned an invalid method name: '${method}'. This is not a function on the '${resource}' resource.`;
      Logger.log(errorMessage);
      return errorMessage;
    }

    Logger.log(`[DEBUG] SKIPPING CONFIRMATION. Intended action: ${summary}`);
    Logger.log(`[DEBUG] Executing: ${resource}.${method} with params: ${JSON.stringify(params)}`);

    let result;
    if (method.toLowerCase().includes('create')) {
        const resourceBody = params[1];
        const courseIdParam = params[0];
        result = apiObject[method](resourceBody, courseIdParam);
    } else {
        result = apiObject[method](...params);
    }

    Logger.log(`[DEBUG] API call successful. Result: ${JSON.stringify(result).substring(0, 500)}...`);

    if (method.toLowerCase().includes('list')) {
      const resultKey = Object.keys(result).find(key => Array.isArray(result[key]));
      const items = resultKey ? result[resultKey] : [];
      if (items.length === 0) return `[DEBUG] SUCCESS: No items found.`;
      
      const names = items.map(item => item.profile ? item.profile.name.fullName : item.title || 'Unknown Item');
      return `[DEBUG] SUCCESS: Found ${items.length} items:\n- ${names.join('\n- ')}`;
    }

    return `[DEBUG] SUCCESS: Action "${summary}" completed.`;

  } catch (e) {
    Logger.log(`[DEBUG] API call FAILED: ${e.toString()}`);
    return `[DEBUG] API call FAILED: ${e.message}`;
  }
}


// =====================================================================================================================
// --- GEMINI API INTERACTION ---
// =====================================================================================================================

/**
 * Constructs a detailed system prompt and sends the user's command to the Gemini API.
 * @param {string} prompt The user's natural language command.
 * @return {Object|null} The parsed JSON object from Gemini's response.
 */
function getStructuredDataFromGemini(prompt) {
  const config = getConfiguration();
  const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const todayString = new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD format

  const geminiApiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${config.gemini_api_key}`;

  const systemPrompt = `You are an expert assistant that translates natural language commands into a flexible JSON object representing Google Classroom API calls. The course ID is '${config.course_id}'. The user's timezone is ${timezone}. The current date is ${todayString}.
  RULES:
  1. For multi-step commands (e.g., "create 3 assignments"), respond with a root object containing a single key "apiCalls", which is an ARRAY of apiCall objects. For single commands, use a single "apiCall" object.
  2. For 'create' methods, the 'params' array MUST contain two elements: the courseId ('${config.course_id}') as the first element, and the resource body object as the second.
  3. If the resource body contains a 'dueDate', it MUST also contain a 'dueTime' object. If the user does not specify a time, default to the end of the day: {"hours": 23, "minutes": 59}.
  4. When creating 'CourseWork', the resource body MUST include a "workType" field. Default to "ASSIGNMENT" unless the user specifies "question".
  5. When listing 'CourseWork' (e.g., in a pre_flight call), the parameters MUST include '{"courseWorkStates": ["DRAFT", "PUBLISHED"]}' to ensure draft assignments are found.
  6. For deletion commands, the method name MUST be "remove", not "delete".
  7. Do NOT add a 'state' property to the resource body unless the user explicitly asks to "publish" or "post" it. The default is 'DRAFT'.
  8. For commands that require finding an item before acting (e.g., "delete assignment 'test'"), define a "pre_flight" object to describe the search.
  9. Respond with ONLY a valid JSON object.
  Example: "delete assignment test" -> { "apiCall": { "resource": "Courses.CourseWork", "method": "remove", "params": [], "summary": "Delete assignment named 'test'.", "pre_flight": { "resource": "Courses.CourseWork", "method": "list", "params": ["${config.course_id}", {"courseWorkStates": ["DRAFT", "PUBLISHED"]}], "filter": {"key": "title", "value": "test"}, "select": "id" } } }`;

  const requestBody = {
    "contents": [{ "parts": [{ "text": systemPrompt + "\n\nUser Command: " + prompt }] }],
    "generationConfig": { "responseMimeType": "application/json" }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true
  };

  Logger.log(`[DEBUG] Sending to Gemini: ${JSON.stringify(requestBody)}`);
  const response = UrlFetchApp.fetch(geminiApiEndpoint, options);
  const responseBody = response.getContentText();
  const responseCode = response.getResponseCode();
  Logger.log(`[DEBUG] Gemini Response Code: ${responseCode}`);
  Logger.log(`[DEBUG] Gemini Raw Response Body: ${responseBody}`);

  if (responseCode !== 200) {
      throw new Error(`Gemini API returned error code ${responseCode}`);
  }

  try {
      const parsedResponse = JSON.parse(responseBody);
      let jsonText = parsedResponse.candidates[0].content.parts[0].text;
      jsonText = jsonText.replace(/^```json\s*/, '').replace(/```$/, '');
      return JSON.parse(jsonText);
  } catch(e) {
      Logger.log(`[DEBUG] Failed to parse JSON from Gemini response. Error: ${e.message}`);
      throw new Error(`Could not parse JSON from Gemini's response. Raw response was:\n\n${responseBody}`);
  }
}

/**
 * A dummy function to force the inclusion of all necessary OAuth scopes.
 * This function is never called, but its presence allows Apps Script's
 * static analysis to detect the required permissions without editing the manifest.
 */
function forceScopes() {
  Classroom.Courses.list();
  Classroom.Courses.Students.list('');
  Classroom.Courses.CourseWork.list('');
  Classroom.Courses.CourseWork.create({}, '');
  Classroom.Courses.CourseWork.remove('', '');
  Classroom.Courses.Announcements.create({}, '');
}
