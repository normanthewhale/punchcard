// ============================================================
//  Punchcard – Google Apps Script Backend
//  Paste this entire file into your Apps Script project.
// ============================================================

const SHEET_NAME_LOGS     = "TimeLogs";
const SHEET_NAME_PROJECTS = "Projects";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var output;
  try {
    // GET requests use e.parameter (query string key/value pairs)
    // POST requests use e.postData.contents (JSON body)
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

// ── Get project list ──────────────────────────────────────────────────────────
function getProjects() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_PROJECTS);
  if (!sheet) return { projects: [] };

  var data     = sheet.getDataRange().getValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    if (name && name.toString().trim() !== "") {
      projects.push(name.toString().trim());
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
    sheet.appendRow(["ID","Worker Name","Project","Date","Start Time","End Time","Duration (hrs)","Description","Submitted At"]);
    sheet.getRange(1,1,1,9).setFontWeight("bold").setBackground("#e8f0fe");
    sheet.setFrozenRows(1);
  }

  var id          = Utilities.getUuid();
  var now         = new Date();
  var startDT     = new Date(params.startTime);
  var endDT       = new Date(params.endTime);
  var durationHrs = ((endDT - startDT) / 3600000).toFixed(2);

  sheet.appendRow([
    id,
    params.workerName   || "",
    params.project      || "",
    Utilities.formatDate(startDT, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    Utilities.formatDate(startDT, Session.getScriptTimeZone(), "h:mm a"),
    Utilities.formatDate(endDT,   Session.getScriptTimeZone(), "h:mm a"),
    parseFloat(durationHrs),
    params.description  || "",
    Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
  ]);

  return { success: true, id: id, duration: durationHrs };
}

// ── Get logs ──────────────────────────────────────────────────────────────────
function getLogs(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) return { logs: [] };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var logs    = [];

  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    if (!params.workerName || row["Worker Name"] === params.workerName) {
      logs.push(row);
    }
  }
  return { logs: logs.reverse() };
}

// ── First-time setup ──────────────────────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var projSheet = ss.getSheetByName(SHEET_NAME_PROJECTS);
  if (!projSheet) {
    projSheet = ss.insertSheet(SHEET_NAME_PROJECTS);
    projSheet.appendRow(["Project Name"]);
    projSheet.appendRow(["123 Oak Street"]);
    projSheet.appendRow(["456 Maple Drive"]);
    projSheet.appendRow(["789 Pine Court"]);
    projSheet.getRange(1,1).setFontWeight("bold").setBackground("#e8f0fe");
    projSheet.setFrozenRows(1);
    projSheet.setColumnWidth(1, 220);
  }

  var logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOGS);
    logSheet.appendRow(["ID","Worker Name","Project","Date","Start Time","End Time","Duration (hrs)","Description","Submitted At"]);
    logSheet.getRange(1,1,1,9).setFontWeight("bold").setBackground("#e8f0fe");
    logSheet.setFrozenRows(1);
    var widths = [220,140,180,100,90,90,110,300,160];
    widths.forEach(function(w,i){ logSheet.setColumnWidth(i+1, w); });
  }

  SpreadsheetApp.getUi().alert("✅ Setup complete! Both sheets are ready.\n\nAdd your house/project names to the 'Projects' tab, then deploy as a Web App.");
}

// ── Spreadsheet menu ──────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⏱ Punchcard")
    .addItem("Set up sheets", "setupSheets")
    .addToUi();
}
