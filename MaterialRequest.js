//Globals
var masterProductionCopy = "" //Put master production sheet ID here
var materialRequest = "" //Put material request sheet ID here
var projectTracking = "" //Put project tracking sheet ID here
var projectTrackingNew = "" //Put project tracking new sheet ID here (for testing)

function dates() {

  var ss = SpreadsheetApp.openById(materialRequest);
  var sheet = ss.getSheetByName("Material Request");
  var dateRequested = sheet.getRange("DateRequested");
  var dateNeeded = sheet.getRange("DateNeeded");
  var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");

  dateRequested.getCell(1, 1).setValue(today);
  var twoWeeks = addDays(new Date(), 14);
  dateNeeded.getCell(1, 1).setValue(Utilities.formatDate(twoWeeks, "GMT-4", "MM/dd/YYYY"));
}

  function addDays(theDate, days) {
    return new Date(theDate.getTime() + days*24*60*60*1000);
}

function insertRow() {
  var ss = SpreadsheetApp.openById(materialRequest);
  var sheet = ss.getSheetByName("Material Request");
  var range = sheet.getRange("BlankRow");
  var rowNum = range.getRow();
  sheet.insertRowBefore(rowNum);
  var newRange = sheet.getRange("BlankRow");
  newRange.copyTo(range)
  }

function totalBoardFeet() {
  var ss = SpreadsheetApp.openById(materialRequest);
  var sheet = ss.getSheetByName("Material Request");
  var range = sheet.getRange("BlankRow");
  var rowNum = range.getRow();
  var boardFeet = sheet.getRange("TotalBoardFeet");
  range = sheet.getRange("I10:I"+(rowNum+1));
  Logger.log(rowNum);
  var total = 0;

  for(var i = 1; i < rowNum-9; i++){
    total = total + range.getCell(i,1).getValue();
  }

  boardFeet.getCell(1,1).setValue(total);
}

function submit() {
  var ss = SpreadsheetApp.openById(materialRequest);
  var sheet = ss.getSheetByName("Material Request");
  var range = sheet.getRange("Project");
  var projectName = range.getValue();
  if (projectName == ""){
    Browser.msgBox("No info in project field, form will not be submitted")
    return;
  }
  var order = sheet.getRange("Order").getValue();
    if (order == ""){
      var result = Browser.msgBox("There is no order number, continue?", Browser.Buttons.OK_CANCEL);
      if (result == "cancel"){
        return;
      }
    }
  var printSheet = ss.insertSheet(projectName);
  var lastRow = sheet.getLastRow();
  var printSheetRange = printSheet.getRange(1, 1, lastRow, 9);
  range = sheet.getRange(1, 1, lastRow, 9);
  range.copyTo(printSheetRange);

  for(var i = 1; i <= lastRow; i++){
    printSheet.setRowHeight(i, sheet.getRowHeight(i));
  }

  for(i = 1; i <= 9; i++){
  printSheet.setColumnWidth(i, sheet.getColumnWidth(i));
  }

  printSheet.showSheet()
  var url = emailSpreadsheetAsPDF(printSheet)
  copySheetsToFolder(printSheet, projectName);

 // sendDataToRoughmillList(sheet, url);
 // sendDataToCNCList(sheet, url);
  sendDataToProjectTracking(sheet, url);

  clearContents(sheet, range, lastRow);

  Browser.msgBox("Material Request Submitted Successfully");

}

function sendDataToCNCList(sheet, url) {
  var project = sheet.getRange("Project").getCell(1,1).getValue();
  var order = sheet.getRange("Order").getCell(1,1).getValue();
  var contact = sheet.getRange("Contact").getCell(1,1).getValue();
  var operator = sheet.getRange("Operator").getCell(1,1).getValue();

  var ss = SpreadsheetApp.openById(projectTracking);
  var destinationSheet = ss.getSheetByName("CNC");
  var lastRow = destinationSheet.getLastRow();
  var lastColumn = destinationSheet.getLastColumn();

  var link = '=HYPERLINK("' + url +'","View Request")';

  var masterProduction = SpreadsheetApp.openById(masterProductionCopy);
  var numSheets = masterProduction.getNumSheets();
  var masterSheets = masterProduction.getSheets();

  var bottom = 0;
  var masterRange = 0;

  destinationSheet.insertRowBefore(lastRow);
  var range = destinationSheet.getRange(lastRow, 1, 1, lastColumn);
  var rangeOld = destinationSheet.getRange(lastRow + 1, 1, 1, lastColumn);
  rangeOld.copyTo(range)
  var exit = false;
  var leave = false;
  var date = 0;
  var dim = "";
  var column = 0;
  var fullRange = 0;


  for(var i = 0; i < numSheets; i++){
    bottom = masterSheets[i].getLastRow();
    column = masterSheets[i].getLastColumn();

    masterRange = masterSheets[i].getRange(1, 1, bottom, column);
    var masterData = masterRange.getValues();


    for(var j = 0; j < bottom; j++){
      if (masterData[j][1] == order){

        for (var k = 0; k < column; k++){

          if (masterData[1][k] == "DIMENSIONS"){
            dim = masterData[j][k]
          }
          if (masterData[1][k] == "DUE"){
            date = masterData[j][k]
            exit == true
            break;
          }
        }
        if (exit == true){
          break;
        }
      }
    if (exit == true){
      break;
    }

  }
}
  Logger.log(date);

  range.getCell(1,1).setValue(project);
  range.getCell(1,2).setValue(order);
  range.getCell(1,3).setFormula(link);
  range.getCell(1,4).setValue(dim);
  range.getCell(1,5).setValue(contact);
  range.getCell(1,6).setValue(operator);
  range.getCell(1,9).setValue("Waiting on Material");

  if (date != 0){
    var dueToAssembly = addDays(date, -42);
    range.getCell(1, 7).setValue(Utilities.formatDate(dueToAssembly, "GMT-4", "MM/dd/YYYY"));
  }
  else {
    range.getCell(1, 7).setValue("");
  }



}

function sendDataToProjectTracking(sheet, url) {
  var dateRequest = sheet.getRange("DateRequested").getCell(1,1).getValue();
  var dateNeeded = sheet.getRange("DateNeeded").getCell(1,1).getValue();
  var project = sheet.getRange("Project").getCell(1,1).getValue();
  var order = sheet.getRange("Order").getCell(1,1).getValue();
  var contact = sheet.getRange("Contact").getCell(1,1).getValue();
  var operator = sheet.getRange("Operator").getCell(1,1).getValue();

  var link = '=HYPERLINK("' + url +'","View Request")';

  var ss = SpreadsheetApp.openById(projectTrackingNew);
  sheet = ss.getSheetByName("Project Tracking");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  Logger.log(lastRow);
  sheet.insertRowBefore(lastRow);
  var range = sheet.getRange(lastRow, 1, 1, lastColumn);
  var rangeOld = sheet.getRange(lastRow + 1, 1, 1, lastColumn);
  rangeOld.copyTo(range);

  range.getCell(1,1).setValue(project);
  range.getCell(1,2).setValue(order);
  range.getCell(1,3).setFormula(link);
  range.getCell(1,4).setValue(contact);
  range.getCell(1,5).setValue(operator);
  range.getCell(1,6).setValue(dateRequest);
  range.getCell(1,7).setValue(dateNeeded);

}

function sendDataToRoughmillList(sheet, url) {
  var dateRequest = sheet.getRange("DateRequested").getCell(1,1).getValue();
  var dateNeeded = sheet.getRange("DateNeeded").getCell(1,1).getValue();
  var project = sheet.getRange("Project").getCell(1,1).getValue();
  var order = sheet.getRange("Order").getCell(1,1).getValue();
  var contact = sheet.getRange("Contact").getCell(1,1).getValue();
  var operator = sheet.getRange("Operator").getCell(1,1).getValue();

  var link = '=HYPERLINK("' + url +'","View Request")';

  var ss = SpreadsheetApp.openById(projectTracking);
  sheet = ss.getSheetByName("Rough Mill");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  Logger.log(lastRow);
  sheet.insertRowBefore(lastRow);
  var range = sheet.getRange(lastRow, 1, 1, lastColumn);
  var rangeOld = sheet.getRange(lastRow + 1, 1, 1, lastColumn);
  rangeOld.copyTo(range);

  range.getCell(1,1).setValue(project);
  range.getCell(1,2).setValue(order);
  range.getCell(1,3).setFormula(link);
  range.getCell(1,4).setValue(contact);
  range.getCell(1,5).setValue(operator);
  range.getCell(1,6).setValue(dateRequest);
  range.getCell(1,7).setValue(dateNeeded);


}
function clearContents(sheet, range, lastRow){
  for(var i = 1; i <= 8; i++){
    for(var j = 10; j < lastRow - 2; j++){
      range.getCell(j, i).setValue(0);
      range.getCell(j, i).setValue("");
    }
  }

 sheet.getRange("Finish").getCell(1, 1).setValue("");
 sheet.getRange("TotalBoardFeet").getCell(1, 1).setValue(0);
 sheet.getRange("Project").getCell(1, 1).setValue("");
 sheet.getRange("Order").getCell(1, 1).setValue("");
 sheet.getRange("Operator").getCell(1, 1).setValue("");

}

function copySheetsToFolder(printSheet, projectName){

  var ss = SpreadsheetApp.openById(materialRequest);
  var sentRequests = DriveApp.getFolderById("0B4DPqJEPGYGYanA0ZV93ZW4zUHM");

  var copy = DriveApp.getFileById(ss.getId()).makeCopy(projectName, sentRequests);
  var copyID = copy.getId();
  var copySS = SpreadsheetApp.openById(copyID);
  var colorSheet = copySS.getSheetByName("Material Request");

  copySS.deleteSheet(colorSheet);
  ss.deleteSheet(printSheet);

  printSheet = copySS.getSheetByName(projectName);


  var lastRow = printSheet.getLastRow();
  var printSheetRange = printSheet.getRange(1, 1, lastRow, 9);
  printSheetRange.setFontColor("black");
  printSheetRange.setBackground("white");


}


/* Send Spreadsheet in an email as PDF, automatically */
function emailSpreadsheetAsPDF(printSheet, projectName) {
  var sheet = printSheet
  // Send the PDF of the spreadsheet to this email address
  var email = ""; //Add user email address here
  var range = sheet.getRange(1,1,sheet.getLastRow(),9);
  var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");
  var requestDate = Utilities.formatDate(range.getCell(3, 3).getValue(), "GMT-4", "MM/dd/YYYY");
  var neededDate = Utilities.formatDate(range.getCell(4, 3).getValue(), "GMT-4", "MM/dd/YYYY");
  // Get the currently active spreadsheet URL (link)
  // Or use SpreadsheetApp.openByUrl("<<SPREADSHEET URL>>");
  var ss = SpreadsheetApp.openById(materialRequest);

  // Subject of email message
  var subject = sheet.getName() + " Material Request"

  // Email Body can  be HTML too with your logo image - see ctrlq.org/html-mail
  var body = "Date requested: " + requestDate + "<br>" + "Date needed by: " + neededDate;

  // Base URL
  var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());

  /* Specify PDF export parameters
  From: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579
  */

  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&size=letter'                       // paper size legal / letter / A4
  + '&portrait=false'                    // orientation, false for landscape
  + '&fitw=true&source=labnol'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=';                             // the sheet's Id

  var token = ScriptApp.getOAuthToken();

  //make an empty array to hold your fetched blobs




    // Convert individual worksheets to PDF
    var response = UrlFetchApp.fetch(url + url_ext + sheet.getSheetId(), {
      headers: {
        'Authorization': 'Bearer ' +  token
      }});

    //convert the response to a blob and store in our array
    var blobs = response.getBlob().setName(sheet.getName() + '.pdf');



  //create new blob that is a zip file containing our blob array
//  var zipBlob = Utilities.zip(blobs).setName(ss.getName() + '.zip');

  //optional: save the file to the root folder of Google Drive
  var pdf = DriveApp.createFile(blobs);
  var destination = DriveApp.getFolderById(""); //Add folder ID here
  var copy = pdf.makeCopy(destination);
  var copyURL = copy.getUrl();
  pdf.setTrashed(true);

  // Define the scope
  Logger.log("Storage Space used: " + DriveApp.getStorageUsed());

  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0)
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[blobs]
    });
    return copyURL;

}
