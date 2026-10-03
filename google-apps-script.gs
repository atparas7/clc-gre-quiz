/**
 * CLC GRE Practice — Results Logger (v3: scores + GRE 130–170 scores + every answer + per-student progress)
 *
 * WHAT YOU GET IN THE GOOGLE SHEET
 *   • "Results"  — one row per finished test: time, student, class, test, score, max, %, Section 2 level,
 *                  and for the full mock exams and Verbal adaptive tests the estimated GRE score (130–170) and percentile.
 *   • "Answers"  — one row per question of every finished test: student, test, question, the student's answer,
 *                  the correct answer, Correct/Wrong. Filter by student or test to see exactly what went wrong.
 *   • "Progress" — updates by itself: for every student, tests taken, average %, best %, last test date, best GRE Quant and best GRE Verbal score.
 *
 * FIRST-TIME SETUP
 *   1. Create a Google Sheet (sheets.google.com -> Blank).
 *   2. Extensions -> Apps Script. Delete everything in the editor and paste this whole file. Save.
 *   3. Deploy -> New deployment -> gear icon -> "Web app".
 *      Execute as: Me.   Who has access: Anyone.   Deploy, and authorize when Google asks.
 *   4. Copy the Web app URL (https://script.google.com/macros/s/AKfycb.../exec) and put it in every test .html file
 *      on the line:  const REPORT_WEBHOOK_URL = "...";
 *
 * UPDATING FROM THE OLD VERSION (keeps the SAME URL — no change needed in the test files)
 *   1. Open the Sheet -> Extensions -> Apps Script.
 *   2. Replace all the code with this file. Save.
 *   3. Deploy -> Manage deployments -> pencil icon (Edit) on your existing deployment ->
 *      Version: "New version" -> Deploy.
 *   Your existing rows stay where they are; the first sheet keeps being the "Results" sheet.
 */

var RESULTS_HEADER = ["Timestamp", "Student Name", "Section", "Test", "Score", "Max Score", "Client time", "Percent", "Section 2 level", "GRE Quant (130-170)", "GRE Verbal (130-170)", "Percentile"];
var ANSWERS_HEADER = ["Timestamp", "Student Name", "Section", "Test", "Part", "Question #", "Question", "Student answer", "Correct answer", "Result"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // several students may finish at the same moment
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = (e && e.parameter) || {};
    var now = new Date();

    var results = ss.getSheets()[0];          // the first sheet, as in the original version
    ensureHeader(results, RESULTS_HEADER);
    var pct = data.percent !== undefined && data.percent !== "" ? Number(data.percent) / 100
            : (Number(data.maxScore) ? Number(data.score) / Number(data.maxScore) : "");
    results.appendRow([now, data.name || "", data.section || "", data.test || "",
                       num(data.score), num(data.maxScore), data.when || "", pct, data.level || "",
                       data.greMeasure === "Verbal" ? "" : num(data.greScore), data.greMeasure === "Verbal" ? num(data.greScore) : "", num(data.percentile)]);
    results.getRange(results.getLastRow(), 8).setNumberFormat("0%");

    if (data.answers) {
      var rows = [];
      try { rows = JSON.parse(data.answers); } catch (err) { rows = []; }
      if (rows.length) {
        var answers = ss.getSheetByName("Answers") || ss.insertSheet("Answers");
        ensureHeader(answers, ANSWERS_HEADER);
        var out = rows.map(function (r) {
          return [now, data.name || "", data.section || "", data.test || "", r.s || "", r.n || "", r.q || "",
                  r.y || "", r.c || "", r.ok ? "Correct" : "Wrong"];
        });
        answers.getRange(answers.getLastRow() + 1, 1, out.length, ANSWERS_HEADER.length).setValues(out);
      }
    }
    ensureProgressSheet(ss, results);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "ok" })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "CLC GRE results logger v2 is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function num(v) { var n = Number(v); return v === undefined || v === "" || isNaN(n) ? (v || "") : n; }

function ensureHeader(sheet, header) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
  } else {                                     // upgrade an older header row (adds the new columns at the end)
    var have = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    for (var i = 0; i < header.length; i++) if (!have[i]) sheet.getRange(1, i + 1).setValue(header[i]);
  }
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

var PROGRESS_QUERY = "select B, C, count(H), avg(H), max(H), max(J), max(K), max(A) where B is not null and H is not null group by B, C order by B " +
  "label B 'Student', C 'Class', count(H) 'Tests taken', avg(H) 'Average %', max(H) 'Best %', max(J) 'Best GRE Quant', max(K) 'Best GRE Verbal', max(A) 'Last test'";

function ensureProgressSheet(ss, results) {
  var existing = ss.getSheetByName("Progress");
  if (existing) {                              // upgrade a v2 Progress sheet so it also shows the best GRE scores
    if (String(existing.getRange("A1").getFormula()).indexOf("max(K)") === -1) setProgressFormula(existing, results);
    return;
  }
  // one time: give rows saved by the old version a Percent too, so they count in the summary
  var last = results.getLastRow();
  if (last > 1) {
    var vals = results.getRange(2, 5, last - 1, 4).getValues();          // Score, Max Score, Client time, Percent
    var fill = vals.map(function (r) { return [r[3] !== "" ? r[3] : (Number(r[1]) ? Number(r[0]) / Number(r[1]) : "")]; });
    results.getRange(2, 8, last - 1, 1).setValues(fill).setNumberFormat("0%");
  }
  setProgressFormula(ss.insertSheet("Progress"), results);
}

function setProgressFormula(p, results) {
  var src = "'" + results.getName().replace(/'/g, "''") + "'!A:K";
  p.getRange("A1").setFormula("=QUERY(" + src + ", \"" + PROGRESS_QUERY + "\", 1)");
  p.getRange("D:E").setNumberFormat("0%");
  p.getRange("F:G").setNumberFormat("0");
  p.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm");
  p.setFrozenRows(1);
}
