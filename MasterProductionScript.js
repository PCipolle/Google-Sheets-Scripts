
//Constants
var MASTERPRODUCTION = '' //Add master production sheet ID here
var MPSHEETS = ['M.CROW', 'SPECIAL PROJECTS', 'SOLID WOOD', 'LAKE CREDENZA', 'LEATHER CREDENZA', 'BRONZE CREDENZA', 'CHAIRS', 'TABLES', 'SLABS', 'SOFA', 'UPHOLSTERY']
var SALES = '' //Add sales sheet ID here
var SHOPORDERSEXPORTFOLDER = '' //Add shop orders export folder ID here

//function onOpen() {
//  var ui = SpreadsheetApp.getUi();
//  ui.createMenu('Scripts')
//          .addItem('Import New Orders', 'ImportShopOrders')
//      .addToUi();
//}

function ImportShopOrders() {
//This is the main entry point for this script

    var ss = SpreadsheetApp.openById(MASTERPRODUCTION);
    var result = checkForOrderEmail();
    if(result == false){
      sendErrorEmail('No email attachment found');
      var dailyTotal = 0;
      updateSalesReport(dailyTotal);
      return;
    }
    else{

    var csv = uploadFromEmail();
    var data = Utilities.parseCsv(csv, ',');
    var dataSheet = ss.getSheetByName('IMPORTED DATA');

    var dailyTotal = findNewOrders(data, ss);

    updateSalesReport(dailyTotal);

    findChangedOrders(data,ss);

    replaceOldData(data,dataSheet);

    archiveDailyEmail();
  }

  return;

}

function updateSalesReport(total){
//This function will update the Sales sheet
//Returns nothing

  var ss = SpreadsheetApp.openById(SALES);
  var sheet = ss.getSheetByName('Daily Sales');
  var range = sheet.getRange(sheet.getLastRow() + 1, 1,1, 2);
  var today = Utilities.formatDate(new Date(), "GMT-4", "MM/dd/YYYY");
  range.getCell(1,1).setValue(today);
  range.getCell(1,2).setValue(total);

  return;
}


function sendErrorEmail(strError){
//This functio will send an error email based upon the message sent to the function
//Returns nothing

  var email = '';  //Add email address for error reporting
  var subject = 'Master Production Script Error';
  var body = strError;
  if (MailApp.getRemainingDailyQuota() > 0)
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
    });

  return;
}

function findChangedOrders(data,ss){
//This function searches for order changes by comparing old and new data
//Modify i start to search back further
//Returns nothing
try{
  var tempSheet = ss.getSheetByName('TEMP SHEET');
  tempSheet.clearContents();
  var due = "";
  var oldDataSheet = ss.getSheetByName('IMPORTED DATA');
  var oldDataRange = oldDataSheet.getRange(2, 1, oldDataSheet.getLastRow(), oldDataSheet.getLastColumn());
  var oldData = oldDataRange.getValues();
  var newDataRange = tempSheet.getRange(1, 1, data.length, data[0].length);
  newDataRange.setValues(data);
  var newData = newDataRange.getValues();
  var upper = "";
  var changedOrdersSheet = ss.getSheetByName('CHANGED ORDERS');

  for(var i = oldData.length - 600; i < oldData.length - 1; i++){
    for(var j = 0; j < oldData[0].length; j++){

       if(newData[i][j].toString() != oldData[i][j].toString()){

          if(newData[i][0].toString() != oldData[i][0].toString()){
            var errorMsg = 'Order numbers did not match while searching for changed orders';
            sendErrorEmail(errorMsg);
            return;
          }

          else{
            var range = changedOrdersSheet.getRange(changedOrdersSheet.getLastRow()+1, 1, 1, changedOrdersSheet.getLastColumn());
            range.setBorder(true, true, true, true, true, true);
            var notes = sortNotes(newData, i);

            upper = makeUpperCase(newData[i][0]);
            range.getCell(1,1).setValue(upper);
            upper = makeUpperCase(newData[i][1]);
            range.getCell(1,2).setValue(upper);
            upper = makeUpperCase(newData[i][2]);
            range.getCell(1,3).setValue(upper);
            upper = makeUpperCase(newData[i][3]);
            range.getCell(1,4).setValue(upper);
            upper = makeUpperCase(newData[i][4]);
            range.getCell(1,5).setValue(upper);
            upper = makeUpperCase(newData[i][5]);
            range.getCell(1,6).setValue(upper);
            upper = makeUpperCase(newData[i][6]);
            range.getCell(1,7).setValue(upper);
            upper = makeUpperCase(newData[i][8]);
            range.getCell(1,8).setValue(upper);
            upper = makeUpperCase(notes);
            range.getCell(1,9).setValue(upper);
            upper = makeUpperCase(newData[i][22]);
            due = Utilities.formatDate(new Date(upper), 'GMT-04', 'MM/dd/YYYY');
            range.getCell(1,10).setValue(due);
            upper = makeUpperCase(newData[i][25]);
            range.getCell(1,11).setValue(upper);
            upper = makeUpperCase(newData[i][23]);
            range.getCell(1,12).setValue(upper);


            if(j == 0){
              range.getCell(1,1).setBackground('red');
            }
            else if(j == 1){
              range.getCell(1,2).setBackground('red');
            }
            else if(j == 2){
              range.getCell(1,3).setBackground('red');
            }
            else if(j == 3){
              range.getCell(1,4).setBackground('red');
            }
            else if(j == 4){
              range.getCell(1,5).setBackground('red');
            }
            else if(j == 5){
              range.getCell(1,6).setBackground('red');
            }
            else if(j == 6){
              range.getCell(1,7).setBackground('red');
            }
            else if(j == 8){
              range.getCell(1,8).setBackground('red');
            }
            else if(j == 7 || j == 20 || j == 16 || j == 21 || j == 10 || j == 15 || j == 13 || j == 9 || j == 11 || j == 12 || j == 14 || j == 19 || j == 18 || j ==17){
              range.getCell(1,9).setBackground('red');
            }
            else if(j == 22){
              range.getCell(1,10).setBackground('red');
            }
            else if(j == 25){
              range.getCell(1,11).setBackground('red');
            }
            else if(j == 23){
              range.getCell(1,12).setBackground('red');
            }
          }
       }
     }
  }
}

catch(e){
  var errorMsg = 'There was an error while searching for changed orders';
  sendErrorEmail(errorMsg);
}

  return;
}
function replaceOldData(data, dataSheet){
//This function replaces the old data in the IMPORTED DATA sheet with new data
//Returns nothing

try{
  var dataRange = dataSheet.getRange(2, 1,dataSheet.getLastRow(),dataSheet.getLastColumn());

  dataRange.clearContent();
  dataRange =  dataSheet.getRange(2, 1, data.length, data[0].length);
  dataRange.setValues(data);
}

catch(e){
  var errorMsg = 'There was an error while replacing IMPORTED DATA sheet';
  sendErrorEmail(errorMsg);
}

  return;
}

function uploadFromEmail(){
//This function will retrieve a daily email attachment as a csv file
//Returns the csv file as a Data String
 var threads = GmailApp.getInboxThreads(0, 50);
 for (var i = 0; i < threads.length; i++) {
   if(threads[i].getFirstMessageSubject() == "SHOP ORDERS DAILY"){
      var message = GmailApp.getMessagesForThread(threads[i]);
      var attachment = message[0].getAttachments();
      break;
   }

 }

   var blob = attachment[0].copyBlob();
   archiveIntoDrive(blob);

   return attachment[0].getDataAsString();
}

function archiveIntoDrive(blob){
//This function will archive the daily email attachment into Shop Orders Export Archive folder
//Returns nothing
   var folder = DriveApp.getFolderById(SHOPORDERSEXPORTFOLDER)
   var today = Utilities.formatDate(new Date(), 'GMT-4', 'MM/dd/YYYY');
   var file = DriveApp.createFile(blob);
   var copy = file.makeCopy(folder);
   copy.setName('Daily Shop Orders Export ' + today);
   file.setTrashed(true);

   return;

}


function checkForOrderEmail(){
//This function will check for the SHOP ORDERS DAILY email
//Returns boolean
 var result = false;
 var threads = GmailApp.getInboxThreads(0, 50);
 var user = Session.getActiveUser().getEmail();
 if(user == ""){ //Add user email address here
   for (var i = 0; i < threads.length; i++) {
     Logger.log(threads[i].getFirstMessageSubject());
     if(threads[i].getFirstMessageSubject() == "SHOP ORDERS DAILY"){
        result = true;
        break;
     }
   }
 }
   return result;

}



function archiveDailyEmail(){
//This function will archive the email "SHOP ORDERS DAILY"
//Returns nothing
 var threads = GmailApp.getInboxThreads(0, 50);
 for (var i = 0; i < threads.length; i++) {
   if(threads[i].getFirstMessageSubject() == "SHOP ORDERS DAILY"){
      threads[i].moveToArchive();
   }

 }

  return;
}


function findNewOrders(data, ss){
//Function to search through imported data and find all
//new orders from the current week

try{
  var oldDataSheet = ss.getSheetByName('IMPORTED DATA');
  var oldSize = oldDataSheet.getLastRow();
  var newSize = data.length;
  var start = newSize - (newSize-oldSize) - 1;
  var due = "";
  var newOrdersSheet = ss.getSheetByName('NEW ORDERS');
  var total = 0;
  var temp = 0;
  var upper = 0;
  if(start == newSize){
    var errorMsg = 'There were no new orders to import';
    sendErrorEmail(errorMsg);
    return total;
  }
  else{
    for(var i = start; i < data.length; i++){


          var range = newOrdersSheet.getRange(newOrdersSheet.getLastRow()+1, 1, 1, newOrdersSheet.getLastColumn());
          range.setBorder(true, true, true, true, true, true);
          var notes = sortNotes(data, i);

          upper = makeUpperCase(data[i][0]);
          range.getCell(1,1).setValue(upper);
          upper = makeUpperCase(data[i][1]);
          range.getCell(1,2).setValue(upper);
          upper = makeUpperCase(data[i][2]);
          range.getCell(1,3).setValue(upper);
          upper = makeUpperCase(data[i][3]);
          range.getCell(1,4).setValue(upper);
          upper = makeUpperCase(data[i][4]);
          range.getCell(1,5).setValue(upper);
          upper = makeUpperCase(data[i][5]);
          range.getCell(1,6).setValue(upper);
          upper = makeUpperCase(data[i][6]);
          range.getCell(1,7).setValue(upper);
          upper = makeUpperCase(data[i][8]);
          range.getCell(1,8).setValue(upper);
          upper = makeUpperCase(notes);
          range.getCell(1,9).setValue(upper);
          upper = makeUpperCase(data[i][22]);
          due = Utilities.formatDate(new Date(upper), 'GMT-04', 'MM/dd/YYYY');
          range.getCell(1,10).setValue(upper);
          upper = makeUpperCase(data[i][25]);
          range.getCell(1,11).setValue(upper);
          upper = makeUpperCase(data[i][23]);
          range.getCell(1,12).setValue(upper);

          total = parseInt(total) + parseInt(data[i][25]);
    }
  }
}
catch(e){
  var errorMsg = 'There was an error while searching for new orders';
  sendErrorEmail(errorMsg);
}

return total;
}

function sortNotes(data, i){
//Function to sort the notes portion of the shop orders data
//Returns sorted notes

     var notes = data[i][7];

        if(data[i][20] != ""){
          notes = notes + " / " + data[i][20];
        }
        if(data[i][16] == "YESPARTICULARKINDA PICKY"){
          notes = notes + " / " + "KINDA PICKY"
        }
        else if(data[i][16] == "YESPARTICULARMAKE IT PRETTY"){
          notes = notes + " / " + "MAKE IT PRETTY"
        }
        else{
          notes = notes + "";
        }
        if(data[i][21] != ""){
          notes = notes + " / " + data[i][21];
        }
        if(data[i][10] != ""){
          notes = notes + " / PATCHES: " + data[i][10];
        }
        if(data[i][15] != ""){
          notes = notes + " / CRACKS AND HOLES: "+ data[i][15];
        }
        if(data[i][13] != ""){
          notes = notes + " / EDGES: "+ data[i][13];
        }
        if(data[i][9] != ""){
          notes = notes + " / BRONZE BUTTERFLIES: "+ data[i][9];
        }
        if(data[i][11] != ""){
          notes = notes + " / BUTTERFLIES FINISH: "+ data[i][11];
        }
        if(data[i][12] != ""){
          notes = notes + " / WOOD BUTTERFLIES: "+ data[i][12];
        }
        if(data[i][14] != ""){
          notes = notes + " / BASE FINISH: "+ data[i][14];
        }
        if(data[i][19] != ""){
          notes = notes + " / " + data[i][19];
        }
        if(data[i][18] != ""){
          notes = notes + " / " + data[i][18];
        }
        if(data[i][17] != ""){
          notes = notes + " / " + data[i][17];
        }

    return notes;

}
function makeUpperCase(value){
//Function to make all characters uppercase
  value = value.toString();
  var temp = value.toUpperCase();
  return temp;

}

function subDays(theDate, days) {
//Function to subtract a certain amount of days from the current date
  return new Date(theDate.getTime() - days*24*60*60*1000);
}
