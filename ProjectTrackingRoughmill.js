//Globals
var projectTracking = ""; //Put project tracking sheet ID here
var projectTrackingComplete = ""; //Put completed project tracking sheet ID here

function moveCompletedRoughmill() {
  var ss = SpreadsheetApp.openById(projectTracking);
  var ssComplete = SpreadsheetApp.openById(projectTrackingComplete);
  var sheet = ss.getSheetByName("Rough Mill");
  var cncSheet = ss.getSheetByName("CNC");
  var lastRowCNC = cncSheet.getLastRow();
  var lastColumnCNC = cncSheet.getLastColumn();
  var rangeCNC = cncSheet.getRange(3, 1, lastRowCNC, lastColumnCNC);

  var completedSheet = ssComplete.getSheetByName("Rough Mill Completed");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var lastColumnComplete = completedSheet.getLastColumn();
  var range = sheet.getRange(3, 1, lastRow-3, lastColumn);
  var rowNum = 0;
  var row = 0;
  var destination = 0;

  var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");
  rowNum = completedSheet.getLastRow();
  for(var i = 1; i <= lastRow-3; i++) {
    if (range.getCell(i, 11).getValue() == "Complete"){

      for(var k = 1; k <= lastRowCNC - 2; k++){
        if (rangeCNC.getCell(k, 1).getValue() == range.getCell(i, 1).getValue() | rangeCNC.getCell(k, 2).getValue() == range.getCell(i, 2).getValue()){
          rangeCNC.getCell(k, lastColumnCNC).setValue("Material Received");
        }
      }
      completedSheet.insertRowBefore(rowNum);
      completedSheet.getRange(rowNum+1, 1, 1, lastColumnComplete).copyTo(completedSheet.getRange(rowNum, 1, 1, lastColumnComplete));
      row = completedSheet.getRange(rowNum, 1, 1, lastColumnComplete);
      row.getCell(1, 7).setValue(today);
      row.getCell(1,1).setValue(range.getCell(i,1).getValue());
      row.getCell(1,2).setValue(range.getCell(i,2).getValue());
      row.getCell(1,3).setValue(range.getCell(i,3).getFormula());
      row.getCell(1,4).setValue(range.getCell(i,4).getValue());
      row.getCell(1,5).setValue(range.getCell(i,6).getValue());
      row.getCell(1,6).setValue(range.getCell(i,7).getValue());
      row.getCell(1,8).setValue(range.getCell(i,10).getValue());

      sendEmail(range, i);

      range.getCell(i,11).setValue("Complete & CNC Notified");

    }
  }
}

function sendEmail(range, i) {

var emailContact = range.getCell(i,4).getValue();
var emailOperator = range.getCell(i,5).getValue();



if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}
else if (emailContact == "") { //Programmer contact name here
  emailContact = "" //Programmer email address here
}

if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}
else if (emailOperator == "") { //Operator contact name here
  emailOperator = "" //Operator email address here
}


var email = emailContact;
if (emailOperator != "") {
  email = email + "," + emailOperator
}

var subject = range.getCell(i,1).getValue() + " Milling Completed";
var location = range.getCell(i,8).getValue();
var notes = range.getCell(i,9).getValue();
var link = range.getCell(i,3).getFormula();


  MailApp.sendEmail({
  to: email,
  subject: subject,
  htmlBody: "<b>" + "Location: " + "</b>" + location + "<br>" +  "<b>" + "Notes: "+ "</b>" + notes  + "<font color = white>" + link + "</font>",
  });

}

function dueTwoWeeks(){

  var dateNeeded = 7

  var ss = SpreadsheetApp.openById(projectTracking);
  var sheet = ss.getSheetByName("Rough Mill");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(3, dateNeeded, lastRow -1, 1);
  var data = range.getValues();

  var today = new Date();
  var twoWeeks = addDays(today,14);
  var row = 0;

  for(var i = 0; i <= lastRow-4; i++){
    if (data[i][0] <= twoWeeks){
      row = sheet.getRange(i+3, 1, 1, lastColumn-1);
      row.setBackground("#9fc5e8");
    }
    else {
      row = sheet.getRange(i+3, 1, 1, lastColumn-1);
      row.setBackground("#f3f3f3");
    }
  }


}

function addDays(theDate, days) {
    return new Date(theDate.getTime() + days*24*60*60*1000);
}


function searchTwoWeeks() {

  var dateNeeded = 7

  var ss = SpreadsheetApp.openById(projectTracking);
  var sheet = ss.getSheetByName("Rough Mill");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(3, dateNeeded, lastRow -1, 1);
  var data = range.getValues();

  var today = new Date();
  var twoWeeks = addDays(today,14);
  var row = 0;

  var twoWeeksFormat = Utilities.formatDate(twoWeeks, "GMT-4", "MM/dd/YYYY");
  var dateFormat = 0;
  Logger.log(twoWeeksFormat);

  for(var i = 0; i <= lastRow-4; i++){
    dateFormat = Utilities.formatDate(data[i][0], "GMT-4", "MM/dd/YYYY");
    if (dateFormat == twoWeeksFormat){
      row = sheet.getRange(i+3, 1, 1, lastColumn-1);
      emailTwoWeeks(row);
    }
  }

}

function emailTwoWeeks(row) {

var email = ""; //Add email addresses here
var subject = "DUE TO CNC IN TWO WEEKS: " + row.getCell(1,1).getValue();
var link = row.getCell(1,3).getFormula();
var dateRequested = Utilities.formatDate(row.getCell(1,6).getValue(), "GMT-4", "MM/dd/YYYY");
var dateNeeded = Utilities.formatDate(row.getCell(1,7).getValue(), "GMT-4", "MM/dd/YYYY");

var body = "<b>" + "Date Requested: " + "</b>" + dateRequested + "<br>" + "<b>" + "Due to CNC: " + "</b>" + dateNeeded + "<br>" + "<font color = white>" + link + "</font>";

MailApp.sendEmail({
to: email,
subject: subject,
htmlBody: body,
});

}
