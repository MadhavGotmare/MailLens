/** Starts a fresh, resumable Gmail scan. */
function startMailLensScan() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('MailLens scan', 'Optional Gmail search query:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var query = response.getResponseText().trim() || MAILLENS.DEFAULT_QUERY;
  // Persist the bound spreadsheet ID before creating a continuation trigger.
  getMailLensSpreadsheet_();
  resetMailLensScan_(false);
  saveScanState_({ query: query, offset: 0, processedThreads: 0, startedAt: new Date().toISOString() });
  continueMailLensScan();
}

/** Processes one bounded batch and schedules the next batch if needed. */
function continueMailLensScan() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    var state = loadScanState_();
    if (!state) return;
    var threads = GmailApp.search(state.query, state.offset, MAILLENS.DEFAULT_BATCH_SIZE);
    if (!threads.length) {
      finishMailLensScan_(state);
      return;
    }

    var aggregates = aggregateThreads_(threads);
    upsertSenders_(aggregates);
    state.offset += threads.length;
    state.processedThreads += threads.length;
    state.updatedAt = new Date().toISOString();
    saveScanState_(state);

    if (threads.length < MAILLENS.DEFAULT_BATCH_SIZE) {
      finishMailLensScan_(state);
    } else {
      scheduleContinuation_();
      showMailLensToast_('Processed ' + state.processedThreads + ' threads; scan will continue shortly.', 5);
    }
  } finally {
    lock.releaseLock();
  }
}

function aggregateThreads_(threads) {
  var records = {};
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      var from = message.getFrom();
      var email = normaliseEmail_(from);
      if (!email) return;
      var record = records[email] || (records[email] = {
        email: email, name: displayName_(from, email), count: 0, firstSeen: null, lastSeen: null,
        newsletterCount: 0, unsubscribeCount: 0
      });
      var date = message.getDate();
      record.count++;
      if (!record.firstSeen || date < record.firstSeen) record.firstSeen = date;
      if (!record.lastSeen || date > record.lastSeen) record.lastSeen = date;
      if (isNewsletterMessage_(message)) record.newsletterCount++;
      if (hasUnsubscribeLink_(message)) record.unsubscribeCount++;
    });
  });
  return records;
}

function isNewsletterMessage_(message) {
  return Boolean(message.getHeader('List-Unsubscribe') || message.getHeader('List-Id') ||
    /\b(newsletter|digest|weekly update)\b/i.test(message.getSubject()));
}

function hasUnsubscribeLink_(message) {
  return Boolean(message.getHeader('List-Unsubscribe') || /\bunsubscribe\b/i.test(message.getBody()));
}

function finishMailLensScan_(state) {
  clearContinuationTriggers_();
  PropertiesService.getDocumentProperties().deleteProperty(MAILLENS.STATE_KEY);
  generateSendersReport();
  refreshMailLensDashboard();
  showMailLensToast_('Scan complete: ' + state.processedThreads + ' threads processed.', 8);
}

function loadScanState_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(MAILLENS.STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}
function saveScanState_(state) {
  PropertiesService.getDocumentProperties().setProperty(MAILLENS.STATE_KEY, JSON.stringify(state));
}
function scheduleContinuation_() {
  clearContinuationTriggers_();
  ScriptApp.newTrigger(MAILLENS.TRIGGER_HANDLER).timeBased().after(60 * 1000).create();
}
function clearContinuationTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === MAILLENS.TRIGGER_HANDLER) ScriptApp.deleteTrigger(trigger);
  });
}
function resetMailLensScan() {
  resetMailLensScan_(true);
}
function resetMailLensScan_(notify) {
  clearContinuationTriggers_();
  PropertiesService.getDocumentProperties().deleteProperty(MAILLENS.STATE_KEY);
  if (notify) showMailLensToast_('Saved scan progress was cleared.', 5);
}
