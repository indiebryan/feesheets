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
Add Credit Card support (Capital One)
*/

//Global Variables BEGIN --

var ss = SpreadsheetApp.getActive();
var sheet = ss.getSheetByName("Now");

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

function onOpen() {
  //Create the buttons the User will see at the top of the Spreadsheet page
  createMenuButtons();
  
}

function createMenuButtons() {
  //Create the Initialize Data menu and Wells Fargo submenu
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('MAT')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Import')
      .addItem('Wells Fargo', 'importData'))
      .addItem('Create Pie Chart', 'createPieChart')
      .addToUi();
  
}

//Import data from CSV file - courtesy of this fellow: https://gist.github.com/dommmel/b7e0f52a046b392c9c93
function importData(e) {
  activeCell = sheet.getActiveCell();
  var app = UiApp.createApplication().setTitle("Upload CSV file");
  var formContent = app.createVerticalPanel();
  formContent.add(app.createFileUpload().setName('thefile'));
  formContent.add(app.createSubmitButton('Start Upload'));
  var form = app.createFormPanel();
  form.add(formContent);
  app.add(form);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

function doPost(e) {
  // data returned is a blob for FileUpload widget
  var fileBlob = e.parameter.thefile;

  // parse the data to fill values, a two dimensional array of rows
  // Assuming newlines separate rows and commas separate columns, then:
  var values = [];
  var rows = fileBlob.contents.split('\n');
  //for(var r=0, max_r=rows.length; r<max_r; ++r) {
    //values.push( rows[r].split(',') );  // rows must have the same number of columns
    //For Wells Fargo: Remove these 2 columns as they do not contain relevant data as of 12/11/16
    //values[r].splice(2, 2);
  //}

  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < rows.length; i++) {
    Logger.log(rows[i]);
    sheet.getRange(i + 1, 1).setValue(rows[i]);
  }
  
  //Get range of transaction description column
  //transColRange = sheet.getRange(5, 5, sheet.getLastRow(), 1);
  
  //Copy the transaction description column values
  //var copyVal = transColRange.getValues();
  
  //Create data validation rule to allow user to choose which category each charge falls under
  //var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Gas', 'Rent', 'Bill', 'Rec', 'Groceries', 'Misc', 'Credit Card']).build();
  //transColRange.setDataValidation(rule).clearContent();
  
  //Paste the transaction description column values over 1
  //sheet.getRange(5, 6, sheet.getLastRow(), 1).setValues(copyVal);
  
  initWFData();
  
}

/*
END DOMMEL
*/

function initWFData() {
  //Method Variables --
  
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



//ALL METHODS BELOW THIS LINE ARE DEPRECATED
/*






*/

function onOpen2() {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Add Statement')
      .addItem('Wells Fargo', 'addWellsFargoStatement')
      .addToUi();
}

function addWellsFargoStatement2() {
  Logger.log("starting");
  var ui = SpreadsheetApp.getUi();
  var statementData = ui.prompt('Paste Statement and hit OK').getResponseText();
  Logger.log("about to replace");
  var sd = statementData.replace(/\" /g, ",");
  Logger.log("about to split on ,");
  var dataArray = sd.split(",");
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet("MONTH_NAME");
  var sheet = spreadsheet.getActiveSheet();
  Logger.log("about to sort");
  sortDataArray(dataArray);
}



function sortDataArray2(dat) {
  //for (var i = dat.length-1; i >= 0; i--) {
  //  if (dat[i] === ",") {
  //      dat.splice(i, 1);
  //      // break;       //<-- Uncomment  if only the first term has to be removed
  //  }
  //}
  
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  for (var i = 0; i < dat.length; i ++) {
    if (i % 5 ==0 || i == 0) {
      sheet.appendRow([dat[i], dat[i+1], dat[i+2], dat[i+3], dat[i+4]]);
      //sheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    
  }
}
