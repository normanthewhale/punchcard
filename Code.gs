// ============================================================
//  Punchcard – Google Apps Script Backend 
// ============================================================
// Auto-deploy via GitHub Actions ✓
const SHEET_NAME_LOGS     = "TimeLogs";
const SHEET_NAME_PROJECTS = "Projects";
const SHEET_NAME_WORKERS  = "Workers";

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

    if      (action === "getProjects")    output = getProjects();
    else if (action === "logTime")        output = logTime(params);
    else if (action === "getLogs")        output = getLogs(params);
    else if (action === "editLog")        output = editLog(params);
    else if (action === "saveWorkerRate") output = saveWorkerRate(params);
    else if (action === "getWorkerRates") output = getWorkerRates(params);
    else if (action === "getWorkers")     output = getWorkers(params);
    else                                  output = { error: "Unknown action: " + action };
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
  if (!sheet) return { logs: [], isAdmin: false };

  // ── Admin access check ────────────────────────────────────────────────────
  var isAdmin   = false;
  var adminCode = (params.adminCode || "").toString().trim();
  if (adminCode) {
    var storedCode = PropertiesService.getScriptProperties().getProperty("ADMIN_CODE");
    isAdmin = !!(storedCode && adminCode === storedCode);
  }

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
          val = Utilities.formatDate(val, tz, "h:mm a");
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

  // ── Filter by worker name for non-admin users ─────────────────────────────
  if (!isAdmin) {
    var workerName = (params.workerName || "").toString().trim();
    if (workerName) {
      logs = logs.filter(function(log) {
        return (log["Worker Name"] || "").toString().trim().toLowerCase() === workerName.toLowerCase();
      });
    } else {
      logs = []; // No name provided → return nothing (don't expose all records)
    }
  }

  return { logs: logs.reverse(), isAdmin: isAdmin };
}

// ── Save / update a worker's rate ────────────────────────────────────────────
function saveWorkerRate(params) {
  var name = (params.workerName || "").toString().trim();
  if (!name) return { error: "Name required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_WORKERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_WORKERS);
    sheet.appendRow(["Name", "Rate Type", "Rate Amount"]);
    sheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#e8f0fe");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
  }

  var data    = sheet.getDataRange().getValues();
  var nameCol = 0; // column A
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim().toLowerCase() === name.toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }

  var rateType   = (params.rateType   || "hourly").toString();
  var rateAmount = parseFloat(params.rateAmount) || 0;

  if (rowIndex === -1) {
    sheet.appendRow([name, rateType, rateAmount]);
  } else {
    sheet.getRange(rowIndex, 2).setValue(rateType);
    sheet.getRange(rowIndex, 3).setValue(rateAmount);
  }

  return { success: true };
}

// ── Get worker names list (admin only) ───────────────────────────────────────
function getWorkers(params) {
  var adminCode  = (params.adminCode || "").toString().trim();
  var storedCode = PropertiesService.getScriptProperties().getProperty("ADMIN_CODE");
  if (!storedCode || adminCode !== storedCode) return { error: "Unauthorized" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_WORKERS);
  if (!sheet) return { workers: [] };

  var data    = sheet.getDataRange().getValues();
  var workers = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    if (name) workers.push(name);
  }
  return { workers: workers };
}

// ── Get all worker rates (admin only) ─────────────────────────────────────────
function getWorkerRates(params) {
  // Require admin code to access pay rates
  var adminCode  = (params.adminCode || "").toString().trim();
  var storedCode = PropertiesService.getScriptProperties().getProperty("ADMIN_CODE");
  if (!storedCode || adminCode !== storedCode) return { error: "Unauthorized" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_WORKERS);
  if (!sheet) return { rates: {} };

  var data  = sheet.getDataRange().getValues();
  var rates = {};
  for (var i = 1; i < data.length; i++) {
    var wName = String(data[i][0]).trim();
    if (wName) {
      rates[wName] = {
        rateType:   String(data[i][1] || "hourly"),
        rateAmount: parseFloat(data[i][2]) || 0
      };
    }
  }
  return { rates: rates };
}

// ── Edit a log entry ─────────────────────────────────────────────────────────
function editLog(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) return { error: "No TimeLogs sheet found" };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(String);

  var idCol    = headers.indexOf("ID");
  var startCol = headers.indexOf("Start Time") + 1;
  var endCol   = headers.indexOf("End Time")   + 1;
  var descCol  = headers.indexOf("Description")+ 1;
  var durCol   = headers.indexOf("Duration (hrs)") + 1;

  if (idCol === -1) return { error: "ID column not found" };

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(params.id)) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { error: "Log not found" };

  // Convert "HH:MM" (from time input) → "h:mm AM/PM"
  function toAmPm(hhmm) {
    var p = hhmm.split(":"), h = parseInt(p[0]), m = parseInt(p[1]);
    var ap = h >= 12 ? "PM" : "AM", h12 = h % 12 || 12;
    return h12 + ":" + (m < 10 ? "0" + m : String(m)) + " " + ap;
  }

  if (params.startTime) sheet.getRange(rowIndex, startCol).setNumberFormat("@").setValue(toAmPm(params.startTime));
  if (params.endTime)   sheet.getRange(rowIndex, endCol).setNumberFormat("@").setValue(toAmPm(params.endTime));
  if (params.description !== undefined) sheet.getRange(rowIndex, descCol).setValue(params.description);

  // Recalculate duration when both times provided
  if (params.startTime && params.endTime) {
    var sp = params.startTime.split(":").map(Number);
    var ep = params.endTime.split(":").map(Number);
    var diffHrs = ((ep[0] * 60 + ep[1]) - (sp[0] * 60 + sp[1])) / 60;
    if (diffHrs < 0) diffHrs += 24;
    sheet.getRange(rowIndex, durCol).setNumberFormat("0.00").setValue(parseFloat(diffHrs.toFixed(2)));
  }

  return { success: true };
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

  var workerSheet = ss.getSheetByName(SHEET_NAME_WORKERS);
  if (!workerSheet) {
    workerSheet = ss.insertSheet(SHEET_NAME_WORKERS);
    workerSheet.appendRow(["Name", "Rate Type", "Rate Amount"]);
    workerSheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#e8f0fe");
    workerSheet.setFrozenRows(1);
    workerSheet.setColumnWidth(1, 180);
    workerSheet.setColumnWidth(2, 100);
    workerSheet.setColumnWidth(3, 120);
  }

  SpreadsheetApp.getUi().alert("✅ Setup complete!\n\nProjects tab now has a Category column with a dropdown.\nAdd your properties and assign each a category.");
}

// ── Set admin passcode (run from Sheets menu) ─────────────────────────────────
function setAdminCode() {
  var ui      = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty("ADMIN_CODE");
  var prompt  = current
    ? "A passcode is already set. Enter a new one to replace it, or leave blank to clear it:"
    : "Enter a passcode for admin access in the app (minimum 4 characters):";

  var result = ui.prompt("Set Admin Passcode", prompt, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var code = result.getResponseText().trim();
  if (code === "") {
    PropertiesService.getScriptProperties().deleteProperty("ADMIN_CODE");
    ui.alert("✅ Admin passcode cleared. All users will now see only their own data.");
  } else if (code.length < 4) {
    ui.alert("❌ Passcode must be at least 4 characters. Nothing was changed.");
  } else {
    PropertiesService.getScriptProperties().setProperty("ADMIN_CODE", code);
    ui.alert("✅ Admin passcode saved.\n\nShare this code only with people who should see all workers' data in the app.");
  }
}

// ── Fix column structure (run once if sheet predates the Category+Project feature) ─
function migrateAddProjectColumn() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("❌ No TimeLogs sheet found. Run 'Set up sheets' first.");
    return;
  }

  var lastCol = logSheet.getLastColumn();
  var headers = logSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  var hasCategory = headers.indexOf("Category") !== -1;
  var hasProject  = headers.indexOf("Project")  !== -1;

  if (hasCategory && hasProject) {
    SpreadsheetApp.getUi().alert("✅ Both Category and Project columns already exist — no migration needed.");
    return;
  }

  if (hasProject && !hasCategory) {
    // Column C is labeled "Project" but holds Category data.
    // Fix: rename C → "Category", insert blank "Project" as new column D.
    var projectCol = headers.indexOf("Project") + 1; // 1-based
    logSheet.getRange(1, projectCol).setValue("Category");
    logSheet.insertColumnAfter(projectCol);
    var newCol = logSheet.getRange(1, projectCol + 1);
    newCol.setValue("Project");
    newCol.setFontWeight("bold").setBackground("#e8f0fe");
    logSheet.setColumnWidth(projectCol + 1, 180);

    SpreadsheetApp.getUi().alert(
      "✅ Migration complete!\n\n" +
      "• Column C renamed from \"Project\" → \"Category\" (data was already correct)\n" +
      "• New blank \"Project\" column inserted as column D\n\n" +
      "Existing rows will have a blank Project cell.\n" +
      "Please delete any test entries with bad/shifted data, then log a fresh entry."
    );
    return;
  }

  if (hasCategory && !hasProject) {
    // Category exists but Project is missing — insert Project after Category.
    var categoryCol = headers.indexOf("Category") + 1;
    logSheet.insertColumnAfter(categoryCol);
    var newCol2 = logSheet.getRange(1, categoryCol + 1);
    newCol2.setValue("Project");
    newCol2.setFontWeight("bold").setBackground("#e8f0fe");
    logSheet.setColumnWidth(categoryCol + 1, 180);

    SpreadsheetApp.getUi().alert(
      "✅ \"Project\" column inserted after Category.\n\n" +
      "Please delete any test entries with bad/shifted data, then log a fresh entry."
    );
    return;
  }

  SpreadsheetApp.getUi().alert("❌ Could not determine correct column structure. Check your header row.");
}

// ── Spreadsheet menu ──────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⏱ Punchcard")
    .addItem("Set up sheets", "setupSheets")
    .addItem("Fix columns (add missing Project column)", "migrateAddProjectColumn")
    .addItem("Set admin passcode", "setAdminCode")
    .addToUi();
}