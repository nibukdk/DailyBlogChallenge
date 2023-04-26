function renderForm() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, "Copy Column With Reference");
}

//["Sheet1", "ID", "Name", "Sheet2", "ID", "Name"]
function setTargetColumns(data) {
  try {
    const sourceSheetName = data[0];
    const referenceColumnName = data[1];
    const sourceColumnName = data[2];
    const targetSheetName = data[3];
    const targetReferenceColumnName = data[4];
    const targetColumnName = data[5];

    // get sheet and data
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sheet.getSheetByName(sourceSheetName);

    const sourceSheetData = sourceSheet.getDataRange().getValues();
    const sourceSheetLastRow = sourceSheetData.length;
    const sourceSheetLastColumn = sourceSheetData[0].length;
    // get header column of source sheet
    const sourceSheetHeader = sourceSheet.getRange(1, 1, 1, sourceSheetLastColumn).getValues().flat();

    // find the index of the given column names
    const referenceColumnIndex = sourceSheetHeader.indexOf(referenceColumnName.trim());
    if (referenceColumnIndex === -1) throw "Reference Column Not Found";

    const sourceColumnIndex = sourceSheetHeader.indexOf(sourceColumnName.trim());
    if (sourceColumnIndex === -1) throw "Source Column Not Found"; // if the name is not found then throw error

    const sourceSheetData2 = [sourceSheet.getRange(2, referenceColumnIndex + 1, sourceSheetLastRow, 1).getValues().flat(), sourceSheet.getRange(2, sourceColumnIndex + 1, sourceSheetLastRow, 1).getValues().flat()]


    const targetSheet = sheet.getSheetByName(targetSheetName);
    const targetSheetData = targetSheet.getDataRange().getValues();
    const targetSheetLastRow = targetSheetData.length;


    const targetReferenceColumnIndex = sourceSheetHeader.indexOf(targetReferenceColumnName.trim());
    if (targetReferenceColumnIndex === -1) throw "Target Sheets Reference Column Not Found";

    const targetColumnIndex = sourceSheetHeader.indexOf(targetColumnName.trim());
    if (targetColumnIndex === -1) throw "Target Sheet's Target Column Not Found";

    const targetSheetRefData = targetSheet.getRange(2, targetReferenceColumnIndex + 1, targetSheetLastRow - 1, 1).getValues().flat();
    const targetSheetColData = targetSheet.getRange(2, targetColumnIndex + 1, targetSheetLastRow - 1, 1).getValues();


    for (let i = 0; i < sourceSheetData2[0].length; i++) {
      for (let j = 0; j < targetSheetRefData.length; j++) {
        if (targetSheetRefData[j] === sourceSheetData2[0][i]) {
          targetSheetColData[j] = [sourceSheetData2[1][i]];
          break;
        }
        continue;
      }
    }
    // set new values 
    targetSheet.getRange(2, targetColumnIndex + 1, targetSheetLastRow - 1, 1).setValues(targetSheetColData);

  } catch (e) {
    // alert error
    SpreadsheetApp.getUi().alert(`Error: ${e}`)
  }
}


/**
 * Menu creates menu UI in the document it's bound to.
 */
function createCustomMenu() {
  const menu = SpreadsheetApp.getUi().createMenu("Copy Columns");

  menu.addItem("Copy Column", "renderForm");
  menu.addToUi();
}


/**
 * OnOpen trigger that creates menu
 * @param {Dictionary} e
 */
function onOpen(e) {
  createCustomMenu();
}

/**
 * Word Counter is a script to count number of words written in Google Docs with Google Apps Script.
 * It is created in such a way that it only works with bound script.
 *
 * Created by: Nibesh Khadka.
 * linkedin: https://www.linkedin.com/in/nibesh-khadka/
 * website: https://nibeshkhadka.com
 */


