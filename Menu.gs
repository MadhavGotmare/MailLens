function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MAILLENS.MENU_NAME)
    .addItem('Start scan', 'startMailLensScan')
    .addItem('Continue scan now', 'continueMailLensScan')
    .addSeparator()
    .addItem('Generate sender report', 'generateSendersReport')
    .addItem('Refresh dashboard', 'refreshMailLensDashboard')
    .addItem('Reset scan', 'resetMailLensScan')
    .addToUi();
}

function onInstall() {
  onOpen();
}
