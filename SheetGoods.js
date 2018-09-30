function copyQtyData(){
 var ss = SpreadsheetApp.openById(""); //Put sheet goods ID here
 var numSheets = ss.getNumSheets();
 for(var j = 0; j < numSheets; j++){
   var temp = ss.getSheets()[j];
   var sheetName = temp.getSheetName()

   if(sheetName == "Sheet Goods"){
     var k = j;
   }
  }
 var sheet = ss.getSheets()[k];

 var qtyData = sheet.getRange("C2:C40").copyValuesToRange(sheet, 5, 5, 2, 40);

}


function updateDate(){

 var ss = SpreadsheetApp.openById(""); //Put sheet goods ID here
 var numSheets = ss.getNumSheets();
 for(var j = 0; j < numSheets; j++){
   var temp = ss.getSheets()[j];
   var sheetName = temp.getSheetName()

   if(sheetName == "Sheet Goods"){
     var k = j;
   }
  }

 var sheet = ss.getSheets()[k];

 var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");
 var sheetNm = sheet.getSheetName();
  Logger.log(sheetNm)
 for(var i = 2; i < 40; i++){
  if(sheet.getRange("C" + i).getValue() != sheet.getRange("E" + i).getValue()){
    sheet.getRange("D" + i).setValue(today)
    }
  else if (sheet.getRange("C" + i).getValue() == sheet.getRange("E" + i).getValue()){
    }
  }

}


function sendEmailLowStock(){

var ss = SpreadsheetApp.openById(""); //Put sheet goods ID here
var numSheets = ss.getNumSheets();
  for(var j = 0; j < numSheets; j++){
    var temp = ss.getSheets()[j];
    var sheetName = temp.getSheetName();

    if(sheetName == "Sheet Goods"){
      var k = j;
    }
  }
 var sheet = ss.getSheets()[k];

 var qtyRange = sheet.getRange("C2:C40");
 var typeRange = sheet.getRange("A2:A40");
 var dimRange = sheet.getRange("B2:B40");

 var qtyTemp = qtyRange.getValues();
 var typeTemp = typeRange.getValues();
 var dimTemp = dimRange.getValues();
 var email = "<br>"

   for(var i = 0; i < 37; i++){
     if(qtyTemp[i] <= 5){
       email = email + qtyTemp[i] + " sheets remaining of " + "<b>" + typeTemp[i] + " ( " + dimTemp[i] + " ) " + "</b>" + "<br>"
     }
   }
  MailApp.sendEmail({
  to: "", //Add email addresses here
  subject: "Low Quantity Sheet Goods",
  htmlBody:
  email,
  });

}
