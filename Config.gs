/** Application-wide constants. */
var MAILLENS = Object.freeze({
  MENU_NAME: 'MailLens',
  SENDERS_SHEET: 'Senders',
  DASHBOARD_SHEET: 'Dashboard',
  STATE_KEY: 'maillens.scan.state.v1',
  TRIGGER_HANDLER: 'continueMailLensScan',
  DEFAULT_QUERY: 'in:anywhere',
  DEFAULT_BATCH_SIZE: 100,
  MAX_BATCH_SIZE: 250,
  SENDER_HEADERS: ['Sender', 'From', 'Reply-To']
});
