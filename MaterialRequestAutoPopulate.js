var materialRequest = ""; //Put material request sheet ID here

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Saved Requests')
      .addSubMenu(ui.createMenu('Chairs')
          .addItem('Berin', 'berin')
          .addItem('Square Guest', 'squareGuestChair')
          .addItem('Counter', 'squareGuestCounter')
          .addItem('Bar', 'squareGuestBar'))
      .addToUi();
}

function berin(){
  var berin = ""; //Put sheet ID for berin here
  copyAndInsertRequest(berin);
}

function squareGuestChair(){
  var squareGuestChair = ""; //Put sheet ID for square guest here
  copyAndInsertRequest(squareGuestChair);
}

function squareGuestCounter(){
  var squareGuestCounter = ""; //Put sheet ID for square guest counter here
  copyAndInsertRequest(squareGuestCounter);
}

function squareGuestBar(){
  var squareGuestBar = ""; //Put sheet ID for square guest bar here
  copyAndInsertRequest(squareGuestBar);
}


function copyAndInsertRequest(requestID) {
  var multiplier = Browser.inputBox("Number of pieces?");
  var species = Browser.inputBox("Species?");
  var dest = SpreadsheetApp.openById(materialRequest);
  var source = SpreadsheetApp.openById(requestID);

  var destSheet = dest.getSheetByName("Material Request");
  var sourceSheets = source.getSheets();
  var sourceSheet = sourceSheets[0];

  var lastRowSource = sourceSheet.getLastRow();
  var lastRowDest = destSheet.getLastRow();

  var sourceRange = sourceSheet.getRange(10,1,lastRowSource-3,8);
    for (var j = 1; j<=lastRowSource-9; j++){
      if (sourceRange.getCell(j,1).getValue() == "" & sourceRange.getCell(j,2).getValue() == "" & sourceRange.getCell(j,5).getValue() == "" & sourceRange.getCell(j,6).getValue() == "" & sourceRange.getCell(j,7).getValue() == "" & sourceRange.getCell(j,8).getValue() == ""){
        var lastRow = j-1;
          break;
      }
      else {
        var lastRow = lastRowSource-9;
      }
    }
  var destRange = destSheet.getRange(10, 1, lastRowDest-3, 8);
    for (var j = 1; j<=lastRowDest-9; j++){
      if (destRange.getCell(j,1).getValue() == "" & destRange.getCell(j,2).getValue() == "" & destRange.getCell(j,5).getValue() == "" & destRange.getCell(j,6).getValue() == "" & destRange.getCell(j,7).getValue() == "" & destRange.getCell(j,8).getValue() == ""){
        var lastRowNew = j-1;
          break;
      }
      else {
        var lastRowNew = lastRowDest-9;
      }
    }
  var emptyRows = lastRowDest-12-lastRowNew;

  if (lastRow <= emptyRows){
   var numRowsToAdd = 0;
  }
  else {
    var numRowsToAdd = lastRow - emptyRows;
    insertRows(numRowsToAdd);
  }

  sourceRange = sourceSheet.getRange(10, 1, lastRow, 8);

  var startRow = lastRowNew+10;
  var endRow = lastRow;

  destRange = destSheet.getRange(startRow, 1, endRow, 8);

  var sourceData = sourceRange.getValues();
  Logger.log(multiplier);

  Logger.log(sourceData[0][0]);
  Logger.log(sourceData[0][1]);
  Logger.log(sourceData[0][4]);
  Logger.log(sourceData[0][5]);
  Logger.log(sourceData[0][6]);
  Logger.log(sourceData[0][7]);
  var qty = 0;

   for (var i = 0; i <= lastRow-1; i++){

       qty = sourceData[i][0];
       qty = qty * multiplier
       destRange.getCell(i+1,1).setValue(qty);
       destRange.getCell(i+1,2).setValue(species);
       destRange.getCell(i+1,5).setValue(sourceData[i][4]);
       destRange.getCell(i+1,6).setValue(sourceData[i][5]);
       destRange.getCell(i+1,7).setValue(sourceData[i][6]);
       destRange.getCell(i+1,8).setValue(sourceData[i][7]);
   }
 totalBoardFeet();


}

function insertRows(amount) {
  var ss = SpreadsheetApp.openById(materialRequest);
  var sheet = ss.getSheetByName("Material Request");

  for(var i = 1; i <= amount; i++) {
    var range = sheet.getRange("BlankRow");
    var rowNum = range.getRow();
    sheet.insertRowBefore(rowNum);
    var newRange = sheet.getRange("BlankRow");
    newRange.copyTo(range)
  }
 }
