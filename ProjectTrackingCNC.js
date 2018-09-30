var projectTracking = ""; //Put project tracking sheet ID here
var projectTrackingComplete = ""; //Put completed project tracking sheet ID here

function moveCompletedCNC(){

  var ss = SpreadsheetApp.openById(projectTracking);
  var ssComplete = SpreadsheetApp.openById(projectTrackingComplete);
  var completedSheet = ssComplete.getSheetByName("CNC Completed");
  var sheet = ss.getSheetByName("CNC");
  var lastRowCNC = sheet.getLastRow();
  var lastColumnCNC = sheet.getLastColumn();
  var rangeCNC = sheet.getRange(3, 1, lastRowCNC, lastColumnCNC);
  var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");

  for(var i = 1; i < lastRowCNC-3; i++){
    var rowNum = completedSheet.getLastRow();
    var status = rangeCNC.getCell(i, 9).getValue();

    if(status == "Complete"){
      completedSheet.insertRowBefore(rowNum);
      completedSheet.getRange(rowNum+1, 1, 1, 9).copyTo(completedSheet.getRange(rowNum, 1, 1, 9));
      var row = completedSheet.getRange(rowNum, 1, 1, 9);
      var project = rangeCNC.getCell(i, 1).getValue();
      var order = rangeCNC.getCell(i, 2).getValue();
      var qty = rangeCNC.getCell(i, 8).getValue();
      var request = rangeCNC.getCell(i, 3).getFormula();
      var dims = rangeCNC.getCell(i, 4).getValue();
      var prog = rangeCNC.getCell(i, 5).getValue();
      var oper = rangeCNC.getCell(i, 6).getValue();
      var dueAssem = rangeCNC.getCell(i, 7).getValue();
      row.getCell(1,1).setValue(project);
      row.getCell(1,2).setValue(order);
      row.getCell(1,3).setValue(qty);
      row.getCell(1,4).setFormula(request);
      row.getCell(1,5).setValue(dims);
      row.getCell(1,6).setValue(prog);
      row.getCell(1,7).setValue(oper);
      row.getCell(1,8).setValue(dueAssem);
      row.getCell(1,9).setValue(today);
      rangeCNC.getCell(i,9).setValue("Complete and Assembly Notified");
    }
  }
}
