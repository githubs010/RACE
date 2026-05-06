// ============================================================
// Code.gs  —  Race Lap Time Panel  |  Google Apps Script
// ============================================================
// SETUP:
//  1. Open Google Sheets → Extensions → Apps Script
//  2. Paste this entire file as Code.gs
//  3. Deploy → New deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  4. Copy the Web App URL into the HTML file
// ============================================================

var SHEET_NAME  = "LapTimes";
var EVENT_SHEET = "EventConfig";

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME) {
      sheet.appendRow(["ID","Name","Color","LapTime","LapTimeMs","Event","Venue","Category","Timestamp"]);
      sheet.setFrozenRows(1);
    }
    if (name === EVENT_SHEET) {
      sheet.appendRow(["name","venue","date","cat","note","updatedAt"]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── Router ──────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action || "read";
  try {
    if (action === "read")      return response(readAll());
    if (action === "delete")    return response(deleteRow(e.parameter.id));
    if (action === "clearAll")  return response(clearAll());
    if (action === "readEvent") return response(readEvent());
    return response({ error: "Unknown action" });
  } catch(err) { return response({ error: err.toString() }); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "create";
    if (action === "create")    return response(createRow(data));
    if (action === "update")    return response(updateRow(data));
    if (action === "saveEvent") return response(saveEvent(data));
    return response({ error: "Unknown action" });
  } catch(err) { return response({ error: err.toString() }); }
}

// ── CREATE ───────────────────────────────────────────────────
function createRow(data) {
  var sheet = getOrCreateSheet(SHEET_NAME);
  var id = Utilities.getUuid();
  var ts = new Date().toISOString();
  sheet.appendRow([id, data.name, data.color, data.lapTime, data.lapTimeMs,
                   data.event||"", data.venue||"", data.category||"", ts]);
  return { success: true, id: id, timestamp: ts };
}

// ── READ ALL ─────────────────────────────────────────────────
function readAll() {
  var sheet = getOrCreateSheet(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { records: [] };
  var headers = rows[0];
  var records = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  records.sort(function(a, b) { return a.LapTimeMs - b.LapTimeMs; });
  return { records: records };
}

// ── UPDATE ───────────────────────────────────────────────────
function updateRow(data) {
  var sheet = getOrCreateSheet(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      if (data.name)      sheet.getRange(i+1,2).setValue(data.name);
      if (data.color)     sheet.getRange(i+1,3).setValue(data.color);
      if (data.lapTime)   sheet.getRange(i+1,4).setValue(data.lapTime);
      if (data.lapTimeMs) sheet.getRange(i+1,5).setValue(data.lapTimeMs);
      return { success: true };
    }
  }
  return { error: "Record not found" };
}

// ── DELETE ───────────────────────────────────────────────────
function deleteRow(id) {
  var sheet = getOrCreateSheet(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { error: "Record not found" };
}

// ── CLEAR ALL ────────────────────────────────────────────────
function clearAll() {
  var sheet = getOrCreateSheet(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { success: true };
}

// ── SAVE EVENT CONFIG ────────────────────────────────────────
function saveEvent(data) {
  var sheet = getOrCreateSheet(EVENT_SHEET);
  sheet.clearContents();
  sheet.appendRow(["name","venue","date","cat","note","updatedAt"]);
  sheet.appendRow([data.name||"", data.venue||"", data.date||"", data.cat||"", data.note||"", new Date().toISOString()]);
  return { success: true };
}

// ── READ EVENT CONFIG ────────────────────────────────────────
function readEvent() {
  var sheet = getOrCreateSheet(EVENT_SHEET);
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { event: {} };
  var headers = rows[0], row = rows[1], ev = {};
  headers.forEach(function(h,i){ ev[h] = row[i]; });
  return { event: ev };
}

// ── Helper ───────────────────────────────────────────────────
function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
