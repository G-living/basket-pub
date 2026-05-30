/**
 * TuIA — Early-access subscriber collector
 */

var SHEET_ID = '1PfEWVArQr9jFQYUEjgSrNQoFBpt6tfuQnueLQJEl8nE';
var SHEET_NAME = 'Sheet1';
var HEADERS = ['Timestamp', 'Email', 'Name', 'Locale', 'Source'];

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var email = (p.email || '').toString().trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ result: 'error', message: 'invalid_email' });
    }
    var sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      email,
      (p.name  || '').toString().trim(),
      (p.locale|| '').toString().trim(),
      (p.source|| '').toString().trim()
    ]);
    return json_({ result: 'success' });
  } catch (err) {
    return json_({ result: 'error', message: String(err) });
  }
}

function doGet() {
  return json_({ result: 'ok', service: 'tuia-early-access' });
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}