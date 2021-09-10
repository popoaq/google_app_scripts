/**
 * A Google Apps Script to take IB reports and calculate time weighted returns of each stock
 *
 * To get the input report, go to IB website > Reports > Statements > Activity
 * In the dialog box, choose "Custom Date Range" for "Period"
 * choose a begin and end date
 * choose "CSV" for "Format"
 * then hit the "Run" button
 * copy the contents of the CSV into "Sheet1" of the Google spreadsheet, then run this script
 *
 * The goal is to assess stock picking results, and not provide accurate portfolio returns data
 * methodology of calculating annualized time weighted return for each stock is
 * Σ(current price(i) - buy price(i) / buy price(i) * numShares(i) / (days held) * 365
 * where i is each transaction.
 * e.g. transaction 1 = buy 50 shares FB at $130 on 1/1/2033
 * transaction 2 = buy 100 shares of FB at $200 on 2/1/2033
 *
 * all sell orders are omitted since we're only interested in the performance of our initial picks
 * not our hedging behavior
 * @popoaq
 */
function IB_weighted_returns() {
  var spreadsheet = SpreadsheetApp.getActive()
  var defaultDataEntrySheetName = "Sheet1";
  var calculationSheetName = "trade_calculation";
  var returnSummarySheetName = "return_summary";

  // create new sheets used for calculation and summary pages
  var existingSheet = spreadsheet.getSheetByName(calculationSheetName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  spreadsheet.insertSheet(1).setName(calculationSheetName);
  existingSheet = spreadsheet.getSheetByName(returnSummarySheetName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  spreadsheet.insertSheet(1).setName(returnSummarySheetName);

  // choose data entry sheet
  var sheet = spreadsheet.getSheetByName(defaultDataEntrySheetName);

  var findText = "Trades"; // the rows in the "Trades" category are what we're interested in
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  var categoryData = sheet.getRange(1,1,lastRow).getValues();
  var startRowOfData = null;
  var endRowOfData = null;

// find the start and end row number containing trades information
  for (var i = 0; i < categoryData.length; i++) {
    if (categoryData[i][0] == findText) {
      if (startRowOfData == null) {
        startRowOfData = i + 1;
      }
    } else if (startRowOfData != null && endRowOfData == null) {
      endRowOfData = i;
    }
  }
  if (startRowOfData == null || endRowOfData == null) {
    throw ("unable to find " + findText + " ending the macro");
  }

  var lastRowOfData = endRowOfData - startRowOfData;

  // copy trades into the trade calculation sheet
  sheet.getRange(startRowOfData, 1, lastRowOfData + 1, endRowOfData).copyTo(spreadsheet.getSheetByName(calculationSheetName).getRange(1,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  // change active sheet
  sheet = spreadsheet.getSheetByName(calculationSheetName);

  // add our own data to the end of the data range
  sheet.getRange(1, lastColumn + 1).setValue('current date');
  sheet.getRange(2,lastColumn + 1, lastRowOfData).setFormula('=today()');
  sheet.getRange(1, lastColumn + 2).setValue('current price');
  sheet.getRange(2,lastColumn + 2, lastRowOfData).setFormula('=GOOGLEFINANCE(F2)');

  // this forces the data to be calculated, or else it'll be null for copying purposes
  Logger.log(sheet.getRange(2, lastColumn + 2).getValue())
  // copy the data values only to preserve it in time
  sheet.getRange(1, lastColumn + 1, lastRowOfData, 2).copyTo(sheet.getRange(1, lastColumn + 3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet.getRange(1, lastColumn + 3, lastRowOfData).setNumberFormat("MM/dd/yyyy");

  sheet.getRange(1, lastColumn + 3).setValue('current date frozen');
  sheet.getRange(1, lastColumn + 4).setValue('current price frozen');

  // make calculations to find return
  var dateTimeValues = sheet.getRange(2, 7, lastRowOfData).getValues() // column G == col(7) contains data time data of transaction
  var calculationStartColumn = lastColumn + 5;

  sheet.getRange(1, calculationStartColumn).setValue('timeWeightedReturn');

  // 2D array of [[stock_ticker, return]]
  var stockReturnArray = [];
  var shareWeightedReturnTally = 0;
  var numSharesTally = 0;
  // if value is null, it's a subTotal column, if not null, then parse date
  for (var i = 0; i < dateTimeValues.length; i++) {
    var currentRowIndex = i + 2;
    // if there's no current price, it's an invalid row, skip it.
    var currentPriceVal = parseFloat(sheet.getRange(currentRowIndex, lastColumn + 4 /*current price*/).getValue());
    if (isNaN(currentPriceVal)) {
      continue;
    }

    if (dateTimeValues[i][0] !== "") {
      var currentDateTimeValue = dateTimeValues[i][0];
      // we only care about returns on stocks we have ever bought, so we omit sell order
      var numShares = parseInt(sheet.getRange(currentRowIndex, 8 /*hardcoded share quantity*/).getValue());
      if (numShares < 0) {
        continue;
      }
      // date format in data is "2020-12-17, 12:15:35"
      var dateStrArr = dateTimeValues[i][0].split(",")[0].split("-");
      var transactionDate = new Date(parseInt(dateStrArr[0]), parseInt(dateStrArr[1]) - 1, parseInt(dateStrArr[2]));
      var currentDate = sheet.getRange(currentRowIndex, lastColumn + 3 /*current date*/).getValue();
      var diffInDays = getDiffInDays(currentDate,transactionDate);
      var prevPriceVal = parseFloat(sheet.getRange(currentRowIndex, 9 /*hardcoded transaction price*/).getValue());

      // set cell to weighted return
      var timeWeightedReturn = (currentPriceVal - prevPriceVal) / prevPriceVal / getDiffInDays(currentDate,transactionDate) * 365;
      sheet.getRange(currentRowIndex, calculationStartColumn).setValue(timeWeightedReturn).setNumberFormat("#.###%")

      shareWeightedReturnTally += (timeWeightedReturn * numShares);
      numSharesTally += numShares;
    } else {
      // when the datetime cell is null, this is a subtotal
      var totalWeightedReturn = shareWeightedReturnTally / numSharesTally;
      stockReturnArray.push([sheet.getRange(currentRowIndex, 6).getValue(), totalWeightedReturn]);
      sheet.getRange(currentRowIndex, calculationStartColumn).setValue(totalWeightedReturn).setNumberFormat("#.###%");
      shareWeightedReturnTally = 0;
      numSharesTally = 0;
    }
  }

  // cleaning up for human readability
  sheet.hideColumns(1);
  sheet.hideColumns(3, 3);
  sheet.hideColumns(10, 10);

  // copy subtotals or each stock's time weighted return into a separate sheet for ease of reading
  // change active sheet
  sheet = spreadsheet.getSheetByName(returnSummarySheetName);
  sheet.getRange(1, 1, stockReturnArray.length, stockReturnArray[0].length).setValues(stockReturnArray);
  sheet.getRange(1, 2, stockReturnArray.length, 1).setNumberFormat("#.###%");
};

function getDiffInDays(laterDate, earlierDate) {
  var oneDay = 1000 * 60 * 60 * 24;
  // Calculating the time difference between two dates
  var diffInTime = laterDate.getTime() - earlierDate.getTime();

  // Calculating the no. of days between two dates
  var diffInDays = Math.round(diffInTime / oneDay);

  return diffInDays;
}

