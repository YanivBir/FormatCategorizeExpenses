function ConvertExcelToGoogleSheets(orginalFile, convertedFile) {
  let files = DriveApp.getFilesByName(orginalFile);
  let excelFile = (files.hasNext()) ? files.next() : null;
  let blob = excelFile.getBlob();
  let config = {
    title: convertedFile.toString(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, blob);
  Logger.log("ConvertExcelToGoogleSheets() is finished.");
}

function OpenExternalSpreadsheet(spreadsheetName, spreadsheetTab) {
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

function AppendDataToSheet(sourceSheet, targetSheet) {
 if (!sourceSheet || !targetSheet) {
    Logger.log('Sheets not found.');
    return;
  }

  let sourceData = sourceSheet.getDataRange().getValues();
  let sourceBackgrounds = sourceSheet.getDataRange().getBackgrounds();

  for (var i = 0; i < sourceData.length; i++) {
    targetSheet.appendRow(sourceData[i]);
    targetSheet.getRange(targetSheet.getLastRow(), 1, 1, sourceData[i].length).setBackgrounds([sourceBackgrounds[i]]);
  }

  Logger.log('Sheets appended successfully.');
}
