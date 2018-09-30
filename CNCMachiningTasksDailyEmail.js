function dailyEmailReport() {
 var ss = SpreadsheetApp.openById(""); //Put sheet ID here
 var numSheets = ss.getNumSheets();
 Logger.log(numSheets);
 var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");
 Logger.log(today);
 var machines = ["<b>" + ss.getSheets()[0].getName() + "</b>", "<b>" + ss.getSheets()[1].getName() + "</b>", "<b>" + ss.getSheets()[2].getName() + "</b>", "<b>" + ss.getSheets()[3].getName() + "</b>", "<b>" + ss.getSheets()[4].getName() + "</b>", "<b>" + ss.getSheets()[5].getName() + "</b>", "<b>" + ss.getSheets()[6].getName() + "</b>", "<b>" + ss.getSheets()[7].getName() + "</b>", "<b>" + ss.getSheets()[8].getName() + "</b>", "<b>" + ss.getSheets()[9].getName() + "</b>"];
 var data = ["","","","","","","","","",""];

 var error = 0;

 for(var k = 0; k < numSheets; k++){

   var sheet = ss.getSheets()[k];
   var length = sheet.getLastRow();
   if(length == 3){
   }
   else if (length > 3){
   var range = sheet.getRange(4,1,length-3,3);
   for(var i = 1; i <= length-3;i++){
   var cell = range.getCell(i, 2);
   var cell2 = range.getCell(i, 1);
   if (cell.getValue() == "") {
   }
   else if (cell.getValue() != ""){
     var checked = cell2.getValue();
     try{
     var date = Utilities.formatDate(cell.getValue(), "GMT-4", "MM/dd/YYYY");
     }
     catch(e){
       error = error + 1;
     }
     if (date == today && checked == "0"){
       var task = range.getCell(i,3);
       data[k] = data[k] + task.getValue() + "<br>";
     }
   }
  }
  }
  if (data[k] == ""){
  data[k] = "-----------" + "<br>";
} 
}
  MailApp.sendEmail({
  to: "", //Add email addresses here
  subject: today + " Completed Machining Tasks",
  htmlBody:
  "<br>" + machines[0] + "<br>" + data[0] +
  "<br>" + machines[1] + "<br>" + data[1] +
  "<br>" + machines[2] + "<br>" + data[2] +
  "<br>" + machines[3] + "<br>" + data[3] +
  "<br>" + machines[4] + "<br>" + data[4] +
  "<br>" + machines[5] + "<br>" + data[5] +
  "<br>" + machines[6] + "<br>" + data[6] +
  "<br>" + machines[7] + "<br>" + data[7] +
  "<br>" + machines[8] + "<br>" + data[8] +
  "<br>" + machines[9] + "<br>" + data[9] +
  "<br>" + "There were " + error + " date related errors",
  });

}
