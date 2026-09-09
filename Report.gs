function ensureSendersSheet_() {
  var sheet = getOrCreateSheet_(MAILLENS.SENDERS_SHEET);
  var headers = ['Email', 'Name', 'Message Count', 'First Seen', 'Last Seen', 'Newsletter Messages', 'Unsubscribe Messages', 'Newsletter Likely'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('D:E').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

/** Merges a batch of aggregate records into the report without losing prior batches. */
function upsertSenders_(records) {
  var keys = Object.keys(records);
  if (!keys.length) return;
  var sheet = ensureSendersSheet_();
  var lastRow = sheet.getLastRow();
  var existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  var rowsByEmail = {};
  existing.forEach(function(row, index) { rowsByEmail[String(row[0]).toLowerCase()] = { row: index + 2, values: row }; });
  var updates = [], additions = [];
  keys.forEach(function(email) {
    var incoming = records[email], found = rowsByEmail[email];
    if (!found) {
      additions.push([email, incoming.name, incoming.count, incoming.firstSeen, incoming.lastSeen,
        incoming.newsletterCount, incoming.unsubscribeCount, incoming.newsletterCount > 0 ? 'Yes' : 'No']);
      return;
    }
    var current = found.values;
    updates.push({ row: found.row, values: [email, current[1] || incoming.name, Number(current[2]) + incoming.count,
      earliest_(parseDate_(current[3]), incoming.firstSeen), latest_(parseDate_(current[4]), incoming.lastSeen),
      Number(current[5]) + incoming.newsletterCount, Number(current[6]) + incoming.unsubscribeCount,
      (Number(current[5]) + incoming.newsletterCount) > 0 ? 'Yes' : 'No'] });
  });
  updates.forEach(function(update) { sheet.getRange(update.row, 1, 1, 8).setValues([update.values]); });
  if (additions.length) sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, 8).setValues(additions);
}
function earliest_(a, b) { return !a || (b && b < a) ? b : a; }
function latest_(a, b) { return !a || (b && b > a) ? b : a; }

/** Sorts and formats a durable sender report. */
function generateSendersReport() {
  var sheet = ensureSendersSheet_();
  if (sheet.getLastRow() > 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).sort({ column: 3, ascending: false });
  sheet.autoResizeColumns(1, 8);
  sheet.getRange(1, 1, 1, 8).setBackground('#0b57d0').setFontColor('#ffffff');
}
