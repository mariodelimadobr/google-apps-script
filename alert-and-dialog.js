const ui = SpreadsheetApp.getUi();

function menuItem1() {
  ui.alert('⚠️ Title Alert 1!', 'You clicked the first menu item!', ui.ButtonSet.OK);
}

function menuItem2() {
  ui.alert('⚠️ Title Alert 1!', 'You clicked the second menu item!', ui.ButtonSet.OK);
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, 'My custom dialog');
}
