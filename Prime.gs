/*
Author
--
Bryan C.
@indieprog
http://indieprogrammer.com
--
*/

/*
TODO: 
Add help/instructions interface
Error handling when submitting an incomplete form
*/

//Global Variables BEGIN --

var ss = SpreadsheetApp.getActive();

//An array that holds the value of each negative transaction, chronologically
var allDebitsArray = new Array();
//An array that holds the value of each positive transaction, chronologically
var allCreditsArray = new Array();
//An array that holds the resting total account balance as of the end of each transaction
var accBalanceArray = new Array();
//A static value representing the total account balance before the first transaction of the month
var startingBalance = 1856.13;
//A variable that mutates over time to gather data for the Balance Over Time graph
var accChangeOverTime = 0;

var debitCategories = ['Gas', 'Rent', 'Bill', 'Rec', 'Groceries', 'Credit Card', 'Misc'];

//-- Global Variables END

//Create the buttons the User will see at the top of the Spreadsheet page
function onOpen() {
  createMenuButtons();
}

//Create the Feesheets Menu
function createMenuButtons() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Feesheets')
      .addItem('Import..', 'showDialogImport')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Import')
      .addItem('Wells Fargo', 'importData')
      .addItem('Capital One', 'importDataCO'))
      .addItem('Create Pie Chart', 'createPieChart')
      .addToUi(); 
}

//Show the file upload dialog, called from Page_Import_Select.html
//@param sourceType - A string representing the name of the bank the records came from
function showDialogImport() {
  var html = HtmlService.createTemplateFromFile('Page_Import_File')
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi()
      .showModalDialog(html, "File Upload");
}

function initData(form) {
  var file = form.theFile;
  var bank = form.theBank;
  var month = form.theMonth;
  var year = form.theYear;
  
  //Check to make sure form was completed correctly.  If not, break method
  if (file == null || bank == "NULL" || month == "NULL" || year < 1900) {
    SpreadsheetApp.getUi().alert("Returned from initData, year: " + year);
    return;
  }
  
  var sheet = ss.insertSheet("Newer");
  
  var values = [];
  var rows = file.contents.split('\n');

  for (var i = 0; i < rows.length; i++) {
    Logger.log(rows[i]);
    sheet.getRange(i + 1, 1).setValue(rows[i]);
  }
  
  //Format data according to what bank it came from
  switch (bank) {
      
    case "Wells Fargo":
      initWellsFargo(form);
      break;
      
    case "Capital One":
      initCapitalOne(form);
      break;
      
    default:
      SpreadsheetApp.getUi().alert("Could not match bank in initData");
  }
}


function initCapitalOne(form) {


}

function initWellsFargo() {
  //Method Variables --
  var sheet = SpreadsheetApp.getActive().getSheetByName("Newer");
  //-- 
  
  //Display alert to prevent user interferance with initialization
  showDialog();
  
  //WF data downloads with the most recent transactions at the top, this reverses the order
  sheet.sort(1);
  
  //Loop through each row of unprocessed bank data
  for (var i = 1; i <= sheet.getLastRow(); i++) {
    
    //Get a range of 5 columns for each row (columns 2 - 5 are empty at this point)
    var rowRange = sheet.getRange(i, 1, 1, 5);
    
    
    
    //Set the text alignment of each cell in row i
    var hAlign = [
      ["center", "center", "center", "center", "left"]
    ];
    rowRange.setHorizontalAlignments(hAlign);
    
    //Separate row i into 5 distinct cells based on a , split from the pasted CSV
    var rowString = sheet.getRange(i, 1).getDisplayValue();
    Logger.log("134: rowString: " + rowString);
    var rowSplitArray = rowString.split(",");
    
    //Fill the 5 cells of rowRange with the 5 data values that were just split
    rowRange.setValues([rowSplitArray]);
    
    //Remove the quotes that surround the transaction amount numbers in column 2, so they can be formatted
    var transacAmount = sheet.getRange(i, 2).getDisplayValue().replace(/\"/g, "");
    
    //Modify accChangeOverTime for each transaction to track resting account balance over time
    accChangeOverTime += parseFloat(transacAmount);
    sheet.getRange(i, 3).setValue(startingBalance + parseFloat(accChangeOverTime));
    
    //Sort transactions into positive or negative arrays
    if (transacAmount < 0) {
      allDebitsArray.push(transacAmount);
      sheet.getRange(i, 4).setBackground("Red");
    
    } else
      allCreditsArray.push(transacAmount);
    
    //Set 'Amount' cell to show transaction amount
    sheet.getRange(i, 2).setValue(transacAmount);
  }
  
  //Set formatting for the transaction amount column to USD
  var transAmounts = sheet.getRange(1, 2, sheet.getLastRow(), 1);
  transAmounts.setNumberFormat("$###,##0.00");
  
  //Wells Fargo-specific function complete, now format form for metadata
  formatMetaData(sheet);
}

function formatMetaData(sheet) {
  
  //Move the table down and to the right to make room for metadata at top and left
  sheet.insertRows(1, 6);
  sheet.insertColumns(1, 1);
  
  sheet.getRange(1, 5, 6, 1).setBackground("White");
  
  var headerLabels = [
    ["Starting Balance", "Total Debits", "Total Credits", "Ending Balance"]
  ];
  
  var columnLabels = [
    ["Date", "Amount", "Ending Balance", "Category", "Description"]
  ];
  var totalDebitsAmount = 0;
  var totalCreditsAmount = 0;
  for (var i = 0; i < allDebitsArray.length; i++)
    totalDebitsAmount += parseFloat(allDebitsArray[i]);
  for (var i = 0; i < allCreditsArray.length; i++)
    totalCreditsAmount += parseFloat(allCreditsArray[i]);
  sheet.getRange(2, 3, 1, 4).setValues(headerLabels);
  sheet.getRange(6, 2, 1, 5).setValues(columnLabels);
  sheet.getRange(3, 3).setValue(startingBalance);
  sheet.getRange(3, 4).setValue(totalDebitsAmount);
  sheet.getRange(3, 5).setValue(totalCreditsAmount);
  sheet.getRange(3, 6).setFormula("=C3+D3+E3");
  
  //Create data validation rule to allow user to choose which category each charge falls under
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(debitCategories).build();
  sheet.getRange(7, 5, sheet.getLastRow() - 6).setDataValidation(rule).clearContent();
  
  createCharts();
}

//Pops up warning to prevent user from interfering with data initialization
function showDialog() {
  
  
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Data Initialization in Progress');
}

function createCharts() {
  //createLineChart();
  //createPieChart();
  //fncOpenMyDialog();
}

function createPieChart() {
  var chartCategories = new Array();
  for (var i = 0; i < debitCategories.length; i++) {
    chartCategories.push([debitCategories[i]]);
    Logger.log("cat: " + chartCategories[i][0]);
  }
  var categoryAmounts = new Array();
  for (var i = 0; i < chartCategories.length; i++)
    categoryAmounts.push([0]);
  
  
  
  
  for (var i = 1; i < sheet.getLastRow(); i++) {
    
    
    /*
    TODO: Replace this abomination of a switch statement with a nested for loop
    */
    
    switch (sheet.getRange(i + 6, 5).getDisplayValue()) {
        
      case chartCategories[0][0]:
        categoryAmounts[0][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[1][0]:
        categoryAmounts[1][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[2][0]:
        categoryAmounts[2][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[3][0]:
        categoryAmounts[3][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[4][0]:
        categoryAmounts[4][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[5][0]:
        categoryAmounts[5][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      case chartCategories[6][0]:
        categoryAmounts[6][0] += parseFloat(sheet.getRange(i + 6, 3).getDisplayValue().replace("$",""));
        break;
        
      default:
        Logger.log("Couldn't match category in switch statement.");
        break;
    }
  }
  for ( var i = 0; i < categoryAmounts.length; i++)
    categoryAmounts[i][0] *= -1;
  
  sheet.getRange(sheet.getLastRow() + 1, 2, chartCategories.length, 1).setValues(chartCategories);
  
  //The lastRow has changed by the amount of categories, so we subtract that length to start on the same line as the above code
  sheet.getRange(sheet.getLastRow() + 1 - chartCategories.length, 3, categoryAmounts.length, 1).setValues(categoryAmounts);
  
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.PIE)
     .addRange(sheet.getRange(sheet.getLastRow() - 5, 2, 6, 1))
     .addRange(sheet.getRange(sheet.getLastRow() - 5, 3, 6, 1))
     .setPosition(12, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}

function createLineChart() {
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.LINE)
     .addRange(sheet.getRange("B5:B12"))
     .addRange(sheet.getRange("D5:D12"))
     .setPosition(12, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}

function logi() {
  Logger.log("AHHHHHH");
}


function fncOpenMyDialog() {
  //Open a dialog
  var htmlDlg = HtmlService.createHtmlOutputFromFile('HTML_myHtml')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(150);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'A Title Goes Here');
};

//Strangely enough, this method is required to use custom CSS in HTML Services
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

//DEPRECATED --

//Make a selection screen for what type of data to import
function importSelect() {

   var html = HtmlService.createTemplateFromFile('Page_Import_Select')
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
    SpreadsheetApp.getUi()
      .showModalDialog(html, 'Where are you importing data from?');
}

//Pops up warning to prevent user from interfering with data initialization
function showDialogCOImport(a) {
  
  
  
  /*
  var html = HtmlService.createHtmlOutputFromFile('Page_CO_Import')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(200);
  */
  
  
  var html = HtmlService.createTemplateFromFile('Page_CO_Import')
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Data Initialization in Progress');
}
