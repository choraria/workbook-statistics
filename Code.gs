function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var addOnMenu = ui.createAddonMenu();
  addOnMenu.addItem(' 🔢 Get Stats', 'updateMenu')
  .addToUi();
  
  // This always leads to `needEdit` function with no error on browser console
  //  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
  //    addOnMenu.addItem(' 🔢 Get Stats', 'needEdit')
  //    .addToUi();
  //  } else {
  //    addOnMenu.addItem(' 🔢 Get Stats', 'updateMenu')
  //    .addToUi();
  //  }
  
}

function needEdit() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("You need 'Edit' access to trigger this add-on");
}

function wbStats() {
  var start = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var allSheets = ss.getSheets();
  var title = 'Workbook Statistics';
  
  // Current Sheet Data
  
  ss.toast(" 🔥 Fetching: Current Sheet > Sheet Name...", title, -1);
  var sheetNameCS = activeSheet.getName();
  Logger.log("Current Sheet > Sheet Name " + sheetNameCS);
  
  if (activeSheet.getType() == 'GRID') {
    var dataRange = activeSheet.getDataRange();
    var dataValues = dataRange.getValues();
    var formulaRanges = dataRange.getFormulas();  
    
    ss.toast(" 🔥 Fetching: Current Sheet > End of Sheet...", title, -1);
    var endOfSheetCS;
    try {
      endOfSheetCS = activeSheet.getRange(activeSheet.getLastRow(), activeSheet.getLastColumn()).getA1Notation();
    } catch (e) {
      endOfSheetCS = 'A1';
    }
    Logger.log("Current Sheet > End of Sheet " + endOfSheetCS);
    
    ss.toast(" 🔥 Fetching: Current Sheet > Cells with Data...", title, -1);
    var cellsWithDataCS = dataValues.map(function(sub) {
      return sub.reduce(function(prev, cur) {
        return prev + (!!cur);
      }, 0);
    }).reduce(function(a,b){
      return a + b;
    }, 0);
    Logger.log("Current Sheet > Cells with Data " + cellsWithDataCS);
    
    ss.toast(" 🔥 Fetching: Current Sheet > Named Ranges...", title, -1);
    var namedRangesCS = activeSheet.getNamedRanges().length;
    Logger.log("Current Sheet > Named Ranges " + namedRangesCS);
    
    ss.toast(" 🔥 Fetching: Current Sheet > Pivot Tables...", title, -1);
    var pivotTablesCS = activeSheet.getPivotTables().length;
    Logger.log("Current Sheet > Pivot Tables " + pivotTablesCS);
    
    ss.toast(" 🔥 Fetching: Current Sheet > Formulas...", title, -1);
    var formulasCS = formulaRanges.map(function(sub) {
      return sub.reduce(function(prev, cur) {
        return prev + (!!cur);
      }, 0);
    }).reduce(function(a,b){
      return a + b
    }, 0);
    Logger.log("Current Sheet > Formulas " + formulasCS);
  } else {
    endOfSheetCS = 'N/A';
    cellsWithDataCS = 'N/A';
    namedRangesCS = 'N/A';
    pivotTablesCS = 'N/A';
    formulasCS = 'N/A';
  }
  
  ss.toast(" 🔥 Fetching: Current Sheet > Charts...", title, -1);
  var chartsCS = activeSheet.getCharts().length;
  Logger.log("Current Sheet > Charts " + chartsCS);
  
  // Workbook data
  
  ss.toast(" 🔥 Fetching: Workbook > Sheets...", title, -1);
  var totalSheetsWB = allSheets.length;
  Logger.log("Workbook > Sheets " + totalSheetsWB);
  
  ss.toast(" 🔥 Fetching: Workbook > Cells with Data...", title, -1);
  var cellsWithDataWB = 0;
  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    if (sheet.getType() == 'GRID') {
      var currentSheetValues = ss.getSheetByName(sheet.getName()).getDataRange().getValues();
      var currentSheetData = currentSheetValues.map(function(sub) {
        return sub.reduce(function(prev, cur) {
          return prev + (!!cur);
        }, 0);
      }).reduce(function(a,b){
        return a + b;
      }, 0);
      cellsWithDataWB = cellsWithDataWB + currentSheetData;
    }
  }
  Logger.log("Workbook > Cells with Data " + cellsWithDataWB);
  
  ss.toast(" 🔥 Fetching: Workbook > Named Ranges...", title, -1);
  var namedRangesWB = allSheets.map(function(sheet) {
    return ss.getSheetByName(sheet.getName()).getNamedRanges().length;
  }).reduce(function(a,b){
    return a + b;
  }, 0);
  Logger.log("Workbook > Named Ranges " + namedRangesWB)
  
  ss.toast(" 🔥 Fetching: Workbook > Pivot Tables...", title, -1);
  var pivotTablesWB = allSheets.filter(function(sheet) {
    return ss.getSheetByName(sheet.getName()).getType() == 'GRID';
  }).map(function(sheet) {
    return ss.getSheetByName(sheet.getName()).getPivotTables().length;
  }).reduce(function(a,b){
    return a + b;
  }, 0);
  Logger.log("Workbook > Pivot Tables " + pivotTablesWB);
  
  ss.toast(" 🔥 Fetching: Workbook > Formulas...", title, -1);
  var formulasWB = 0;
  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    if (sheet.getType() == 'GRID') {
      var currentSheetValues = ss.getSheetByName(sheet.getName()).getDataRange().getFormulas();
      var currentSheetFormulas = currentSheetValues.map(function(sub) {
        return sub.reduce(function(prev, cur) {
          return prev + (!!cur);
        }, 0);
      }).reduce(function(a,b){
        return a + b;
      }, 0);
      formulasWB = formulasWB + currentSheetFormulas;
    }
  }
  Logger.log("Workbook > Formulas " + formulasWB);  
  
  ss.toast(" 🔥 Fetching: Workbook > Charts...", title, -1);
  var chartsWB = allSheets.map(function(sheet) {
    return ss.getSheetByName(sheet.getName()).getCharts().length;
  }).reduce(function(a,b){
    return a + b
  }, 0);
  Logger.log("Workbook > Charts " + chartsWB);
  
  var end = new Date().getTime();
  var diffSec = (end - start)/1000;
  var duration = diffSec > 60 ? (diffSec / 60).toFixed(2) + " mins." : diffSec.toFixed(2) + " secs.";
  
  ss.toast(" 🎉 It took " + duration + " to fetch all details. Updating Add-on Menu...", title, 10);
  return {
    "sheetNameCS": sheetNameCS,
    "endOfSheetCS": endOfSheetCS,
    "cellsWithDataCS": cellsWithDataCS,
    "namedRangesCS": namedRangesCS,
    "pivotTablesCS": pivotTablesCS,
    "formulasCS": formulasCS,
    "chartsCS": chartsCS,
    "totalSheetsWB": totalSheetsWB,
    "cellsWithDataWB": cellsWithDataWB,
    "namedRangesWB": namedRangesWB,
    "pivotTablesWB": pivotTablesWB,
    "formulasWB": formulasWB,
    "chartsWB": chartsWB
  }
}

function noAction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(" 🥑 No Action Taken. Please click 'Refresh All' to update stats.", 'Workbook Statistics', -1);
}

function updateMenu() {
  var ui = SpreadsheetApp.getUi();
  var addOnMenu = ui.createAddonMenu();  
  var data = wbStats();
  addOnMenu
  .addSubMenu(ui.createMenu(' ⚡ Combined')
              .addItem(' 📄 Current Sheet', 'noAction')
              .addItem(' ☞ Sheet Name: ' + data.sheetNameCS, 'noAction')
              .addItem(' ☞ End of Sheet: ' + data.endOfSheetCS, 'noAction')
              .addItem(' ☞ Cells with Data: ' + data.cellsWithDataCS, 'noAction')
              .addItem(' ☞ Named Ranges: ' + data.namedRangesCS, 'noAction')
              .addItem(' ☞ Pivot Tables: ' + data.pivotTablesCS, 'noAction')
              .addItem(' ☞ Formulas: ' + data.formulasCS, 'noAction')
              .addItem(' ☞ Charts: ' + data.chartsCS, 'noAction')
              .addSeparator()
              .addItem(' 📓 Workbook', 'noAction')
              .addItem(' ☞ Sheets: ' + data.totalSheetsWB, 'noAction')
              .addItem(' ☞ Cells with Data: ' + data.cellsWithDataWB, 'noAction')
              .addItem(' ☞ Named Ranges: ' + data.namedRangesWB, 'noAction')
              .addItem(' ☞ Pivot Tables: ' + data.pivotTablesWB, 'noAction')
              .addItem(' ☞ Formulas: ' + data.formulasWB, 'noAction')
              .addItem(' ☞ Charts: ' + data.chartsWB, 'noAction')
              .addSeparator()
              .addItem(' 🔄 Refresh All', 'updateMenu'))
  .addSeparator()
  .addSubMenu(ui.createMenu(' ⚡ Current Sheet')
              .addItem(' ☞ Sheet Name: ' + data.sheetNameCS, 'noAction')
              .addItem(' ☞ End of Sheet: ' + data.endOfSheetCS, 'noAction')
              .addItem(' ☞ Cells with Data: ' + data.cellsWithDataCS, 'noAction')
              .addItem(' ☞ Named Ranges: ' + data.namedRangesCS, 'noAction')
              .addItem(' ☞ Pivot Tables: ' + data.pivotTablesCS, 'noAction')
              .addItem(' ☞ Formulas: ' + data.formulasCS, 'noAction')
              .addItem(' ☞ Charts: ' + data.chartsCS, 'noAction')
              .addItem(' 🔄 Refresh Current Sheet', 'updateMenu'))
  .addSeparator()
  .addSubMenu(ui.createMenu(' ⚡ Workbook')
              .addItem(' ☞ Sheets: ' + data.totalSheetsWB, 'noAction')
              .addItem(' ☞ Cells with Data: ' + data.cellsWithDataWB, 'noAction')
              .addItem(' ☞ Named Ranges: ' + data.namedRangesWB, 'noAction')
              .addItem(' ☞ Pivot Tables: ' + data.pivotTablesWB, 'noAction')
              .addItem(' ☞ Formulas: ' + data.formulasWB, 'noAction')
              .addItem(' ☞ Charts: ' + data.chartsWB, 'noAction')
              .addItem(' 🔄 Refresh Workbook', 'updateMenu'))
  .addSeparator()
  .addItem(' 🔄 Refresh All', 'updateMenu')
  .addToUi();  
}
