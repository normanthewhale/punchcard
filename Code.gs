// ============================================================
//  Punchcard – Google Apps Script Backend
// ============================================================

const SHEET_NAME_LOGS     = "TimeLogs";
const SHEET_NAME_PROJECTS = "Projects";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var output;
  try {
    var params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter || {};
    }

    var action = params.action;

    if      (action === "getProjects") output = getProjects();
    else if (action === "logTime")     output = logTime(params);
    else if (action === "getLogs")     output = getLogs(params);
    else                               output = { error: "Unknown action: " + action };
  } catch(err) {
    output = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Get project list (now includes Category column) ───────────────────────────
function getProjects() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_PROJECTS);
  if (!sheet) return { projects: [] };

  var data     = sheet.getDataRange().getValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) {
    var name     = data[i][0];
    var category = data[i][1] || "Other";
    if (name && name.toString().trim() !== "") {
      projects.push({
        name:     name.toString().trim(),
        category: category.toString().trim()
      });
    }
  }
  return { projects: projects };
}

// ── Log a time entry ──────────────────────────────────────────────────────────
function logTime(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_LOGS);
    sheet.appendRow(["ID","Worker Name","Category","Project","Date","Start Time","End Time","Duration (hrs)","Description","Submitted At"]);
    sheet.getRange(1,1,1,10).setFontWeight("bold").setBackground("#e8f0fe");
    sheet.setFrozenRows(1);
  }

  var id          = Utilities.getUuid();
  var tz          = Session.getScriptTimeZone();
  var now         = new Date();
  var startDT     = new Date(params.startTime);
  var endDT       = new Date(params.endTime);
  var durationHrs = ((endDT - startDT) / 3600000).toFixed(2);

  // Pre-format the next row's date/time columns as plain text so Sheets
  // doesn't auto-convert the formatted strings back into date/time cell types.
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 5).setNumberFormat("@");   // Date
  sheet.getRange(nextRow, 6).setNumberFormat("@");   // Start Time
  sheet.getRange(nextRow, 7).setNumberFormat("@");   // End Time
  sheet.getRange(nextRow, 8).setNumberFormat("0.00"); // Duration (hrs)

  sheet.appendRow([
    id,
    params.workerName   || "",
    params.category     || "",
    params.project      || "",
    Utilities.formatDate(startDT, tz, "yyyy-MM-dd"),
    Utilities.formatDate(startDT, tz, "h:mm a"),
    Utilities.formatDate(endDT,   tz, "h:mm a"),
    parseFloat(durationHrs),
    params.description  || "",
    Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss")
  ]);

  return { success: true, id: id, duration: durationHrs };
}

// ── Get logs ──────────────────────────────────────────────────────────────────
function getLogs(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) return { logs: [] };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(String);
  var tz      = Session.getScriptTimeZone();
  var logs    = [];

  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var val    = data[i][j];
      var header = headers[j];
      // Sheets auto-converts date/time strings to Date objects on read.
      // Convert them back to the string format the frontend expects.
      if (val instanceof Date) {
        if (header === "Date") {
          val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else if (header === "Start Time" || header === "End Time") {
          // Time-only cells land on the 1899 epoch — read UTC to avoid TZ shift
          if (val.getUTCFullYear() <= 1900) {
            var h   = val.getUTCHours();
            var m   = val.getUTCMinutes();
            var ap  = h >= 12 ? "PM" : "AM";
            var h12 = h % 12 || 12;
            val = h12 + ":" + (m < 10 ? "0" + m : String(m)) + " " + ap;
          } else {
            val = Utilities.formatDate(val, tz, "h:mm a");
          }
        } else if (header === "Duration (hrs)") {
          // Time-formatted duration: extract hours from the UTC time component
          val = parseFloat((val.getUTCHours() + val.getUTCMinutes() / 60 + val.getUTCSeconds() / 3600).toFixed(2));
        } else {
          val = Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss");
        }
      }
      row[header] = val;
    }
    logs.push(row);
  }
  return { logs: logs.reverse() };
}

// ── First-time setup ──────────────────────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var projSheet = ss.getSheetByName(SHEET_NAME_PROJECTS);
  if (!projSheet) {
    projSheet = ss.insertSheet(SHEET_NAME_PROJECTS);
    projSheet.appendRow(["Project Name", "Category"]);
    projSheet.appendRow(["123 Oak Street",  "New Construction"]);
    projSheet.appendRow(["456 Maple Drive", "New Construction"]);
    projSheet.appendRow(["789 Pine Court",  "Rental"]);
    projSheet.getRange(1,1,1,2).setFontWeight("bold").setBackground("#e8f0fe");
    projSheet.setFrozenRows(1);
    projSheet.setColumnWidth(1, 220);
    projSheet.setColumnWidth(2, 160);

    // Add dropdown validation for Category column
    var categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["New Construction", "Rental", "Other"], true)
      .build();
    projSheet.getRange("B2:B1000").setDataValidation(categoryRule);
  }

  var logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOGS);
    logSheet.appendRow(["ID","Worker Name","Category","Project","Date","Start Time","End Time","Duration (hrs)","Description","Submitted At"]);
    logSheet.getRange(1,1,1,10).setFontWeight("bold").setBackground("#e8f0fe");
    logSheet.setFrozenRows(1);
    var widths = [220,140,140,180,100,90,90,110,300,160];
    widths.forEach(function(w,i){ logSheet.setColumnWidth(i+1, w); });
  }

  SpreadsheetApp.getUi().alert("✅ Setup complete!\n\nProjects tab now has a Category column with a dropdown.\nAdd your properties and assign each a category.");
}

// ── Spreadsheet menu ──────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⏱ Punchcard")
    .addItem("Set up sheets", "setupSheets")
    .addToUi();
}