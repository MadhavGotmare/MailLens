/**
 * Gets the bound report spreadsheet, including when invoked by a time trigger
 * where Apps Script does not provide an active spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getMailLensSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getDocumentProperties().setProperty(MAILLENS.SPREADSHEET_ID_KEY, active.getId());
    return active;
  }
  var spreadsheetId = PropertiesService.getDocumentProperties().getProperty(MAILLENS.SPREADSHEET_ID_KEY);
  if (!spreadsheetId) {
    throw new Error('MailLens has no report spreadsheet. Start a scan from the bound Google Sheet first.');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function showMailLensToast_(message, seconds) {
  getMailLensSpreadsheet_().toast(message, MAILLENS.MENU_NAME, seconds);
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
