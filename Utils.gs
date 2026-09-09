/** @return {GoogleAppsScript.Spreadsheet.Spreadsheet} */
function getMailLensSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name) {
  var spreadsheet = getMailLensSpreadsheet_();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function normaliseEmail_(value) {
  var match = String(value || '').match(/<([^>]+)>/);
  var email = (match ? match[1] : String(value || '')).trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}

function displayName_(value, email) {
  var name = String(value || '').replace(/<[^>]+>/, '').replace(/["']/g, '').trim();
  return name || email;
}

function parseDate_(value) {
  if (!value) return null;
  var date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function safeSheetName_(name) {
  return String(name).replace(/[\\/?*\[\]:]/g, ' ').slice(0, 100);
}
