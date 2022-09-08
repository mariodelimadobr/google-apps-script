//@OnlyCurrentDoc
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Import CSV data 👉️")
    .addItem("Import from URL", "importCSVFromUrl")
    .addItem("Import from Drive", "importCSVFromDrive")
    .addToUi();
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert"); 
}

//Imports a CSV file at a URL into the Google Sheet
function importCSVFromUrl() {
  var url = promptUserForInput("Please enter the URL of the CSV file:");
  var contents = Utilities.parseCsv(UrlFetchApp.fetch(url));
  var sheetName = writeDataToSheet(contents);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
}
 
//Imports a CSV file in Google Drive into the Google Sheet
function importCSVFromDrive() {
  var fileName = promptUserForInput("Please enter the name of the CSV file to import from Google Drive:");
  var files = findFilesInDrive(fileName);
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
    return;
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.");
    return;
  }
  var file = files[0];
  var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheetName = writeDataToSheet(contents);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
}

//Prompts the user for input and returns their response
function promptUserForInput(promptText) {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt(promptText);
  var response = prompt.getResponseText();
  return response;
}

//Returns files in Google Drive that have a certain name.
function findFilesInDrive(filename) {
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}

//Inserts a new sheet and writes a 2D array of data in it
function writeDataToSheet(data) {
  var ss = SpreadsheetApp.getActive();
  sheet = ss.insertSheet();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return sheet.getName();
}
