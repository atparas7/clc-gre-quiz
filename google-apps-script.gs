/**
 * CLC GRE Practice — Results Logger
 *
 * SETUP:
 * 1. Create a new Google Sheet (sheets.google.com -> Blank).
 * 2. In that Sheet, go to Extensions -> Apps Script.
 * 3. Delete anything in the editor and paste this whole file.
 * 4. Click the disk icon (Save).
 * 5. Click "Deploy" -> "New deployment".
 *    - Click the gear icon next to "Select type" -> choose "Web app".
 *    - Description: anything, e.g. "CLC GRE results".
 *    - Execute as: Me.
 *    - Who has access: Anyone.
 *    - Click "Deploy".
 * 6. Authorize the script when Google asks (it's your own script, this is safe).
 * 7. Copy the "Web app URL" it gives you — it looks like:
 *    https://script.google.com/macros/s/AKfycb.../exec
 * 8. Paste that URL into every test .html file, replacing
 *    "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE"
 *    (look for the line: const REPORT_WEBHOOK_URL = "...";)
 *
 * Every time a student finishes a test, a new row will appear automatically
 * in this Sheet with: timestamp, name, section, test name, score, max score.
 */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Add header row once, if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Student Name", "Section", "Test", "Score", "Max Score", "Client time"]);
  }

  var data = e.parameter || {};
  sheet.appendRow([
    new Date(),
    data.name || "",
    data.section || "",
    data.test || "",
    data.score || "",
    data.maxScore || "",
    data.when || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Optional: lets you open the Web App URL directly in a browser to check it's alive
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "CLC GRE results logger is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}
