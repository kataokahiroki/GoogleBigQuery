var SHEET = {SQL : 0, RESULT : 1};

var LOC = {
  PROJECT_ID : {X : 2, Y : 1},
  SQL : {X : 2, Y : 2},  
  RESULT : {XHS: 1, YHS : 1, XDS : 1, YDS : 2}
};

function onOpen() {
  addMenu();
}

function addMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [];

  menu.push({name : "Do query", functionName : "doQuery"});
  menu.push(null);
  menu.push({name : "Clear result", functionName : "clearResult"});
  
  ss.addMenu("Custom Menu...", menu);
}

function clearResult() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[SHEET.RESULT]; 
  if (sheet.getMaxRows() > 1) {sheet.deleteRows(1, sheet.getMaxRows() - 1);}
  if (sheet.getMaxColumns() > 1) {sheet.deleteColumns(1, sheet.getMaxColumns() - 1);}
  sheet.clear();     
}

function doQuery () {
  try {    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var projectID = ss.getSheets()[SHEET.SQL].getRange(LOC.PROJECT_ID.Y, LOC.PROJECT_ID.X).getValue();
    var sql = ss.getSheets()[SHEET.SQL].getRange(LOC.SQL.Y, LOC.SQL.X).getValue();
    
    clearResult();
    var sheet_result = ss.getSheets()[SHEET.RESULT];  
    sheet_result.activate(); 
    BQLib.doQuery(projectID, sql, sheet_result, LOC);
  } catch(e) {
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().getSheets()[SHEET.RESULT].getRange(1, 1).setValue(e);
  }    
}

