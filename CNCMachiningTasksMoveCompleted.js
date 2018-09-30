
function moveToCompletedTasks(){

 var ss = SpreadsheetApp.openById(""); //Put sheet ID here
 var numSheets = ss.getNumSheets();
 var machines = ["<b>Accord 1</b>", "<b>Accord 2</b>", "<b>Accord 3</b>", "<b>Ergon</b>", "<b>Lasers</b>", "<b>Record 125</b>", "<b>Record 130</b>", "<b>Record 132</b>", "<b>Record 240</b>", "<b>Techni</b>"];
 var data = ["","","","","","","","","",""];
 var target = SpreadsheetApp.openById(""); //Put target sheet ID here

 for(var k = 0; k < numSheets; k++){

   var sheetSource = ss.getSheets()[k];
   var sheetTarget = target.getSheets()[k];
   var sourceLength = sheetSource.getLastRow();
   var targetLength = sheetTarget.getLastRow() + 1;
   var length = sheetSource.getLastRow();
   if(length == 3){
   }
   else if (length > 3){
   var range = sheetSource.getRange(4,1,length-3,4);
   for(var i = 1; i <= length-3;i++){
   var cell = range.getCell(i, 2);
   if (cell.getValue() == "") {
   }
   else if (cell.getValue() != ""){
   var check = "0";
   var date = cell.getValue();
   var task = range.getCell(i,3).getValue();
   var comments = range.getCell(i,4).getValue();

   var targetCheck = sheetTarget.getRange(targetLength,1);
   var targetDate = sheetTarget.getRange(targetLength,2);
   var targetTask = sheetTarget.getRange(targetLength,3);
   var targetComments = sheetTarget.getRange(targetLength,4);

   targetCheck.setValue(check);
   targetDate.setValue(date);
   targetTask.setValue(task);
   targetComments.setValue(comments);

   sheetSource.getRange(i+3,1,1,4).clearContent();
   sheetSource.getRange(i+3,1,1,4).setBackground("white");

   targetLength = targetLength + 1;

     }
   }
sheetSource.sort(3);

  }
}
}
