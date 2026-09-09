/** Builds an executive dashboard from the normalized Senders report. */
function refreshMailLensDashboard() {
  var senders = ensureSendersSheet_();
  var dashboard = getOrCreateSheet_(MAILLENS.DASHBOARD_SHEET);
  dashboard.clear();
  dashboard.getCharts().forEach(function(chart) { dashboard.removeChart(chart); });

  var lastRow = senders.getLastRow();
  var data = lastRow > 1 ? senders.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  var totalMessages = data.reduce(function(sum, row) { return sum + Number(row[2] || 0); }, 0);
  var newsletters = data.filter(function(row) { return row[7] === 'Yes'; });
  var newsletterMessages = newsletters.reduce(function(sum, row) { return sum + Number(row[5] || 0); }, 0);

  dashboard.getRange('A1:D1').merge().setValue('MailLens Dashboard').setFontSize(18).setFontWeight('bold').setBackground('#0b57d0').setFontColor('#ffffff');
  dashboard.getRange('A3:B6').setValues([
    ['Metric', 'Value'],
    ['Unique senders', data.length],
    ['Messages scanned', totalMessages],
    ['Newsletter senders', newsletters.length]
  ]);
  dashboard.getRange('A3:B3').setFontWeight('bold').setBackground('#e8f0fe');
  dashboard.getRange('D3:F3').setValues([['Top senders', 'Messages', 'Newsletter?']]).setFontWeight('bold').setBackground('#e8f0fe');
  var top = data.slice().sort(function(a, b) { return Number(b[2]) - Number(a[2]); }).slice(0, 10)
    .map(function(row) { return [row[0], Number(row[2]), row[7]]; });
  if (top.length) dashboard.getRange(4, 4, top.length, 3).setValues(top);

  dashboard.getRange('A8:B10').setValues([
    ['Email category', 'Messages'],
    ['Newsletter', newsletterMessages],
    ['Other', Math.max(0, totalMessages - newsletterMessages)]
  ]);
  dashboard.getRange('A8:B8').setFontWeight('bold').setBackground('#e8f0fe');
  if (totalMessages) {
    var chart = dashboard.newChart().asPieChart().addRange(dashboard.getRange('A8:B10'))
      .setPosition(8, 4, 0, 0).setOption('title', 'Messages by category').build();
    dashboard.insertChart(chart);
  }
  dashboard.setFrozenRows(1);
  dashboard.autoResizeColumns(1, 6);
}
