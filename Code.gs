/**
 * AI English Speech Lab - Google Apps Script Backend API
 * Connected to Google Sheet: EchoScore-Sheet
 * Sheet URL: https://docs.google.com/spreadsheets/d/12bc3j6bjRtwhXh5JIha3fpVopdLxCKVDjRT-sJx-1AY/edit
 */

const SPREADSHEET_ID = "12bc3j6bjRtwhXh5JIha3fpVopdLxCKVDjRT-sJx-1AY";

/**
 * Automatically creates and formats the 9 standard columns if the sheet is empty
 */
function setupSheetHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "timestamp",
      "className",
      "studentId",
      "studentName",
      "attemptCount",
      "targetSentence",
      "completion",
      "accuracy",
      "transcript"
    ]);
    
    // Style the header row with Indigo background and bold white text
    const headerRange = sheet.getRange(1, 1, 1, 9);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4f46e5");
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
}

/**
 * 1. GET Request - Health check & auto-initializes sheet headers
 */
function doGet(e) {
  try {
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch(err) {}
    
    if (!ss && SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    const sheet = ss.getActiveSheet();
    setupSheetHeaders(sheet);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "AI Pronunciation Assessment API is running! Headers verified.",
      sheetName: sheet.getName(),
      rowCount: sheet.getLastRow()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2. POST Request - Appends student practice result into the 9 columns
 */
function doPost(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      params = e.parameter;
    } else {
      params = {};
    }

    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch(err) {}
    
    if (!ss && SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    const sheet = ss.getActiveSheet();
    setupSheetHeaders(sheet);

    const completionStr = (params.completion !== undefined ? params.completion + "%" : "0%");
    const accuracyStr = (params.accuracy !== undefined ? params.accuracy + "%" : "0%");

    // Append row with 9 columns in strict order
    sheet.appendRow([
      new Date(),                      // A: timestamp
      params.className || "",          // B: className
      params.studentId || "",          // C: studentId
      params.studentName || "",        // D: studentName
      params.attemptCount || 0,        // E: attemptCount
      params.targetSentence || "",     // F: targetSentence
      completionStr,                   // G: completion
      accuracyStr,                     // H: accuracy
      params.transcript || ""          // I: transcript
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Practice result saved successfully"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
