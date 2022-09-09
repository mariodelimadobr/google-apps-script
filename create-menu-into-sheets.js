
/**
GOOGLE APPS SCRIPT 
CREATE MENU AND SUBMENU
*/

const ui = SpreadsheetApp.getUi();

function onOpen() {
  ui.createMenu('Custom Menu')
    .addItem('Show dialog', 'showDialog')
    .addItem('First item', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu')
      .addItem('Second item', 'menuItem2'))
    .addSeparator()
    .addItem('Remover Duplicadas', 'removeDuplicates')
    .addToUi();
}

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
