// =====================================================================================================================
// --- RUBRIC AND COMMAND CENTER FUNCTIONS ---
// This file contains all the logic for creating student sheets, the command center, and managing rubrics.
// =====================================================================================================================

// =================================================================
// --- AUTOMATIC TRIGGER & HELPER FUNCTIONS ---
// =================================================================

/**
 * An simple trigger that runs automatically when a user edits the spreadsheet.
 * This checks if a student sheet was edited and updates their score.
 * @param {Object} e The event object from the edit.
 */
function onEdit(e) {
  const editedSheet = e.source.getActiveSheet();
  const editedSheetName = editedSheet.getName();

  if (editedSheetName === "Command Center") return;

  updateScoreForStudent(editedSheetName);
}

/**
 * Finds a specific student's score and updates their row in the Command Center.
 * @param {string} studentName The name of the student to update.
 */
function updateScoreForStudent(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const commandCenterSheet = ss.getSheetByName("Command Center");
  if (!commandCenterSheet) return;

  const studentNames = commandCenterSheet.getRange("A:A").getValues();
  const studentRowIndex =
    studentNames.findIndex((row) => row[0] === studentName) + 1;

  if (studentRowIndex === 0) return;

  const studentSheet = ss.getSheetByName(studentName);
  let score = "Not Found";
  let lastUpdated = "";

  if (studentSheet) {
    const data = studentSheet.getDataRange().getValues();
    if (data.length > 0) {
      const headers = data[0];
      const averageColIndex = headers.findIndex(
        (header) => header.toString().trim() === "Overall Average"
      );
      const totalRowIndex = data.findIndex((row) =>
        row[0].toString().includes("Weekly Total:")
      );

      if (averageColIndex !== -1 && totalRowIndex !== -1) {
        score = data[totalRowIndex][averageColIndex];
        lastUpdated = new Date();
      }
    }
  }

  const scoreCell = commandCenterSheet.getRange(studentRowIndex, 2);
  const timeCell = commandCenterSheet.getRange(studentRowIndex, 3);
  const rowRange = commandCenterSheet.getRange(studentRowIndex, 1, 1, 3);

  scoreCell.setValue(score);
  timeCell.setValue(lastUpdated).setNumberFormat("M/d/yyyy h:mm am/pm");

  if (score === "Not Found") {
    rowRange.setBackground("#f4cccc");
  } else {
    rowRange.setBackground(null);
  }
}

// =================================================================
// --- INITIAL SETUP & STUDENT LIST FUNCTIONS ---
// =================================================================

/**
 * Fetches the student roster, creates sheets for them, and populates the Command Center.
 */
function createSheetsFromClassroom() {
  if (!checkConfiguration()) return;
  const config = getConfiguration();
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const students = getStudentRoster(config.course_id);
  if (!students) return;

  if (students.length === 0) {
    ui.alert(
      "No students found in the specified course. Please check the Course ID."
    );
    return;
  }

  const allSheetNames = spreadsheet.getSheets().map((sheet) => sheet.getName());

  let createdCount = 0;
  students.forEach((student) => {
    const studentName = student.profile.name.fullName;
    if (!allSheetNames.includes(studentName)) {
      spreadsheet.insertSheet(studentName);
      createdCount++;
    }
  });

  const studentNamesForSheet = students.map((student) => [
    student.profile.name.fullName,
  ]);
  const targetSheet =
    spreadsheet.getSheetByName("Command Center") || spreadsheet.getSheets()[0];

  targetSheet.getRange(2, 1, targetSheet.getMaxRows(), 1).clearContent();
  targetSheet
    .getRange(2, 1, studentNamesForSheet.length, 1)
    .setValues(studentNamesForSheet);

  const summaryMessage =
    `Process Complete. New sheets created: ${createdCount}.\n\n` +
    `The student list in your Command Center has been updated with ${studentNamesForSheet.length} names.`;
  ui.alert(summaryMessage);
}

/**
 * Sets up the Command Center sheet and populates initial scores.
 */
function createCommandCenter() {
  if (!checkConfiguration()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const commandCenterSheet = ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();

  try {
    commandCenterSheet.setName("Command Center");
    commandCenterSheet.setFrozenRows(1);

    const headers = [["Student Name", "Overall Score", "Last Updated"]];
    const headerRange = commandCenterSheet.getRange(1, 1, 1, 3);

    headerRange
      .setValues(headers)
      .setFontWeight("bold")
      .setBackground("#d9ead3");

    const studentListRange = commandCenterSheet.getRange(
      2,
      1,
      commandCenterSheet.getLastRow() - 1,
      1
    );
    const studentNames = studentListRange
      .getValues()
      .map((row) => row[0])
      .filter(String);

    if (studentNames.length === 0) {
      ui.alert(
        'No student names found in Column A. Please run "1. Create Student Sheets & Update List" first.'
      );
      return;
    }

    studentNames.forEach((name) => updateScoreForStudent(name));

    commandCenterSheet.autoResizeColumn(1);
    commandCenterSheet.autoResizeColumn(2);
    commandCenterSheet.autoResizeColumn(3);

    ui.alert(
      "Success!",
      "Command Center has been set up and all student scores have been populated.",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Error processing Command Center: " + e.message);
    ui.alert(
      "An error occurred while processing the Command Center. Please check the logs for details."
    );
  }
}

// =================================================================
// --- RUBRIC HELPER FUNCTIONS ---
// =================================================================

/**
 * Main function to set up the rubric on the active sheet.
 */
function setupRubric() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Setup Rubric",
    "Enter the total number of weeks for the rubric:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() != ui.Button.OK) return;

  const desiredWeeks = parseInt(response.getResponseText(), 10);
  if (isNaN(desiredWeeks) || desiredWeeks < 0) {
    ui.alert(
      "Invalid Input",
      "Please enter a valid, non-negative number.",
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    if (sheet.getLastRow() === 0) {
      const defaultCategories = [
        ["Category"],
        ["Professionalism"],
        ["Safety"],
        ["Clean Up"],
        ["Participation"],
        ["Weekly Total:"],
      ];
      sheet
        .getRange(1, 1, defaultCategories.length, 1)
        .setValues(defaultCategories);
    }

    const columnAValues = sheet.getRange("A:A").getValues();
    let totalRow =
      columnAValues.findIndex((row) =>
        row[0].toString().includes("Weekly Total:")
      ) + 1;

    if (totalRow === 0) {
      ui.alert(
        "Error",
        "Could not find the 'Weekly Total:' row. Please ensure this text exists in Column A.",
        ui.ButtonSet.OK
      );
      return;
    }

    const startDataRow = 2;
    const endDataRow = totalRow - 1;

    if (startDataRow >= endDataRow) {
      ui.alert(
        "Error",
        "No categories found. Please add category names above the 'Weekly Total:' row.",
        ui.ButtonSet.OK
      );
      return;
    }

    const lastColCheck = sheet.getLastColumn();
    if (sheet.getRange(1, lastColCheck).getValue() === "Overall Average") {
      sheet.deleteColumn(lastColCheck);
    }

    const currentWeeks = Math.floor((sheet.getLastColumn() - 1) / 2);

    if (desiredWeeks > currentWeeks) {
      const weeksToAdd = desiredWeeks - currentWeeks;
      for (let i = 0; i < weeksToAdd; i++) {
        const newWeekNum = currentWeeks + i + 1;
        const lastCol = sheet.getLastColumn();
        sheet.insertColumnsAfter(lastCol, 2);
        sheet.getRange(1, lastCol + 1).setValue(`Week ${newWeekNum}`);
        sheet.getRange(1, lastCol + 2).setValue(`Notes`);
      }
    } else if (desiredWeeks < currentWeeks) {
      const weeksToRemove = currentWeeks - desiredWeeks;
      const startDeleteColumn = desiredWeeks * 2 + 2;
      sheet.deleteColumns(startDeleteColumn, weeksToRemove * 2);
    }

    if (sheet.getLastColumn() > 1) {
      sheet.getRange(totalRow, 2, 1, sheet.getLastColumn()).clearContent();
    }

    const weeklyTotalCells = [];
    for (let week = 1; week <= desiredWeeks; week++) {
      const scoreColumn = week * 2;
      const formula = `=SUM(${sheet
        .getRange(startDataRow, scoreColumn)
        .getA1Notation()}:${sheet
        .getRange(endDataRow, scoreColumn)
        .getA1Notation()})`;
      sheet.getRange(totalRow, scoreColumn).setFormula(formula);
      weeklyTotalCells.push(
        sheet.getRange(totalRow, scoreColumn).getA1Notation()
      );
    }

    const overallTotalColumn = desiredWeeks * 2 + 2;
    sheet.getRange(1, overallTotalColumn).setValue("Overall Average");

    if (weeklyTotalCells.length > 0) {
      const averageFormula = `=AVERAGE(${weeklyTotalCells.join(",")})`;
      sheet.getRange(totalRow, overallTotalColumn).setFormula(averageFormula);
    }

    applyRubricFormatting(sheet, totalRow, endDataRow, startDataRow);
    ui.alert(
      "Success!",
      `The rubric has been set up for ${desiredWeeks} weeks.`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      "Error",
      "Could not set up the rubric. Error: " + e.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Applies standard formatting to the rubric for better readability.
 */
function applyRubricFormatting(sheet, totalRow, endDataRow, startDataRow) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headerColor = "#d9ead3";
  const totalColor = "#fce5cd";
  const stripeColor = "#f3f3f3";
  const scoreFillColor = "#b6d7a8";

  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearFormat();
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  sheet
    .getRange(1, 1, 1, lastCol)
    .setBackground(headerColor)
    .setFontWeight("bold");
  sheet
    .getRange(totalRow, 1, 1, lastCol)
    .setBackground(totalColor)
    .setFontWeight("bold");

  for (let i = startDataRow; i <= endDataRow; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, lastCol).setBackground(stripeColor);
    }
  }

  sheet.clearConditionalFormatRules();
  const scoreRange = sheet.getRange(
    startDataRow,
    2,
    endDataRow - startDataRow + 1,
    lastCol - 2
  );
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellNotEmpty()
    .setBackground(scoreFillColor)
    .setRanges([scoreRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// =================================================================
// --- UTILITY AND CONFIGURATION FUNCTIONS ---
// =================================================================

/**
 * Fetches all Classroom courses and displays them in a custom HTML dialog for selection.
 */
function selectAndSetCourseId() {
  try {
    // Calling list without any parameters fetches courses of all states.
    const response = Classroom.Courses.list();
    const courses = response.courses;
    const ui = SpreadsheetApp.getUi();

    if (!courses || courses.length === 0) {
      ui.alert(
        "No Courses Found",
        "No courses were found for your account.\n\nPlease ensure:\n1. The Classroom API is enabled for this script.\n2. You have authorized the script with the correct Google account.",
        ui.ButtonSet.OK
      );
      return;
    }

    // Pass the course data to the HTML template to be rendered.
    const htmlTemplate = HtmlService.createTemplateFromFile("SelectCourse");
    htmlTemplate.courses = courses;
    htmlTemplate.current_id = getConfiguration().course_id;

    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(350);

    ui.showModalDialog(htmlOutput, "Select an Active Course");
  } catch (err) {
    Logger.log("Failed to fetch courses with error: %s", err.message);
    const ui = SpreadsheetApp.getUi();
    if (err.message.includes("Classroom is not defined")) {
      ui.alert(
        "Error: The Classroom API is not enabled.",
        'Please go to the Script Editor, click on "Services +", find "Google Classroom API", and click "Add".',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "Failed to fetch courses. Please check your permissions. Error: " +
          err.message
      );
    }
  }
}

/**
 * Saves the selected course ID to script properties. This function is called
 * from the client-side script in the 'SelectCourse.html' dialog.
 * @param {string} courseId The ID of the course selected by the user.
 */
function saveSelectedCourseId(courseId) {
  if (courseId) {
    const ui = SpreadsheetApp.getUi();
    PropertiesService.getScriptProperties().setProperty("course_id", courseId);
    // FIXED: Added the third parameter (ui.ButtonSet.OK) to the alert call.
    ui.alert(
      "Success",
      `The active Course ID has been set to: ${courseId}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Fetches the student roster from a specific Google Classroom course.
 * @param {string} courseId The ID of the course to get the roster from.
 * @returns {Array|null} An array of student objects, or null if an error occurs.
 */
function getStudentRoster(courseId) {
  try {
    const response = Classroom.Courses.Students.list(courseId);
    return response.students;
  } catch (e) {
    Logger.log("Error fetching student roster: " + e.message);
    if (e.message.includes("Classroom is not defined")) {
      SpreadsheetApp.getUi().alert(
        "Error: The Classroom API is not enabled.",
        'Please go to the Script Editor, click on "Services +", find "Google Classroom API", and click "Add".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        "An error occurred while fetching the student roster. Please ensure your Course ID is correct and that you have permission to access it. Check logs for details."
      );
    }
    return null;
  }
}
