function ConvertExcelToGoogleSheets(orginalFile, convertedFile) {
  let files = DriveApp.getFilesByName(orginalFile);
  let excelFile = (files.hasNext()) ? files.next() : null;
  let blob = excelFile.getBlob();
  let config = {
    title: convertedFile.toString() , //sets the title of the converted file
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, blob);
  Logger.log("ConvertExcelToGoogleSheets() is finished.");
}

function OpenExternalSpreadsheet(spreadsheetName, spreadsheetTab) {
  // Search for the file in Google Drive by name.
  let files = DriveApp.getFilesByName(spreadsheetName);

  if (files.hasNext()) {
    let file = files.next();
    let spreadsheet = SpreadsheetApp.open(file);
    let sheet = spreadsheet.getSheetByName(spreadsheetTab);

    if (sheet) {
      Logger.log("Sheet: \"" + spreadsheetName +  "\" is found, opened and activted tab: \"" + spreadsheetTab + "\".");
      return sheet;
    } else {
      Logger.log("Sheet not found in the opened spreadsheet.");
    }
  } else {
    Logger.log("Spreadsheet not found in Google Drive.");
  }
}

function DeleteFileIfExists(fileName) {
  let files = DriveApp.getFilesByName(fileName);

  while (files.hasNext()) {
    let file = files.next();
    file.setTrashed(true); // Move the file to the trash (effectively deleting it)
    Logger.log("File: \"" + fileName + "\" found and deleted.");
  }
}

function MoveColumn(sheet, sourceIndex, destinationIndex) {
  let sourceRange = sheet.getRange(1, sourceIndex, sheet.getLastRow(), 1);
  let sourceValues = sourceRange.getValues();
  sheet.deleteColumn(sourceIndex);
  sheet.insertColumnBefore(destinationIndex);

  let destinationRange = sheet.getRange(1, destinationIndex, sheet.getLastRow(), 1);
  destinationRange.setValues(sourceValues);
}

function AppendDataToSheet(sourceSheet, destinationSheet) {
  if (sourceSheet && destinationSheet) {
    let data = sourceSheet.getDataRange().getValues();
    let lastRow = destinationSheet.getLastRow();
    destinationSheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
    Logger.log("Data appended from sourceSheet to destinationSheet.");
  } else {
    Logger.log('Source or destination sheet not found');
  }
}
