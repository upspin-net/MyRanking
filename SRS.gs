// This is a backup from App Script Editor
function myFunction() {
  //generateTTCANRatingSheet();
  generateUSATTRatingSheet();
}

function generateTTCANRatingSheet() {
  var tsn = "ttcan-temp"; // temp sheet name
  var rsn = "TTCAN";  // result sheet name
  var spreadsheet = SpreadsheetApp.getActive();
  var ttcan_url = "http://www.ttcan.ca/ratingSystem/ctta_ratings2.asp?Category_code=39&Full_Name=&Period_Issued=381&Prov=&Reg=&Region=&Sex=&Formv_ctta_ratings_Page=%d#v_ctta_ratings"

  //spreadsheet.deleteSheet(spreadsheet.getSheetByName(tsn))
  spreadsheet.deleteSheet(spreadsheet.getSheetByName(rsn))
  spreadsheet.insertSheet(rsn);
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 50000);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('Serial');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Name');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Prov');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setValue('Gender');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setValue('Rating');
  spreadsheet.getRange('F1').activate();
  spreadsheet.getCurrentCell().setValue('Period');
  spreadsheet.getRange('G1').activate();
  spreadsheet.getCurrentCell().setValue('Last Played');
  spreadsheet.getRange('H1').activate();
  spreadsheet.getCurrentCell().setValue('Temp');

  for(var i = 1; i < 45; i++){
    var url = Utilities.formatString(ttcan_url, i);
    Logger.log("url to fetch: %s", url);
    var formula = Utilities.formatString("=importhtml(\"%s\", \"table\")", url);
    Logger.log("formula: %s", formula);

    spreadsheet.insertSheet(tsn);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().setFormula(formula);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(rsn), true);
    var range = Utilities.formatString("A%d", (i - 1) * 100 + 2);
    spreadsheet.getRange(range).activate();
    spreadsheet.getRange('\'ttcan-temp\'!A2:G101').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(tsn))
  }
}

function generateUSATTRatingSheet() {
  var tsn = "usatt-temp"; // temp sheet name
  var rsn = "USATT";  // result sheet name
  var spreadsheet = SpreadsheetApp.getActive();
  var pageSize = 1000;
  var usatt_url = "https://usatt.simplycompete.com/userAccount/s2?citizenship=&gamesEligibility=&gender=&minAge=&maxAge=&minTrnRating=&maxTrnRating=&minLeagueRating=&maxLeagueRating=&state=&region=Any+Region&favorites=&q=&displayColumns=First+Name&displayColumns=Last+Name&displayColumns=USATT%23&displayColumns=Location&displayColumns=Home+Club&displayColumns=Tournament+Rating&displayColumns=Last+Played+Tournament&displayColumns=League+Rating&displayColumns=Last+Played+League&displayColumns=Membership+Expiration&pageSize=%d&format=&offset=%d&max=%d"

  var s = spreadsheet.getSheetByName(tsn);
  if (s != null) {
    spreadsheet.deleteSheet(s);  
  }
  s = spreadsheet.getSheetByName(rsn);
  if (s != null) {
    spreadsheet.deleteSheet(s);
  }  
  
  //spreadsheet.deleteSheet(spreadsheet.getSheetByName(tsn))
  //spreadsheet.deleteSheet(spreadsheet.getSheetByName(rsn))
  spreadsheet.insertSheet(rsn);
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 100000);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('#');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('First Name');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Last Name');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setValue('USATT#');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setValue('Location');
  spreadsheet.getRange('F1').activate();
  spreadsheet.getCurrentCell().setValue('Home Club');
  spreadsheet.getRange('G1').activate();
  spreadsheet.getCurrentCell().setValue('Tournament Rating');
  spreadsheet.getRange('H1').activate();
  spreadsheet.getCurrentCell().setValue('Last Played Tournament');
  spreadsheet.getRange('I1').activate();
  spreadsheet.getCurrentCell().setValue('League Rating');
  spreadsheet.getRange('J1').activate();
  spreadsheet.getCurrentCell().setValue('Last Played League');
  spreadsheet.getRange('K1').activate();
  spreadsheet.getCurrentCell().setValue('Membership Expiration');
  spreadsheet.getRange('L1').activate();
  spreadsheet.getCurrentCell().setValue('Gender');

  for(var i = 1; i < 13; i++){
    var url = Utilities.formatString(usatt_url, pageSize, (i - 1) * pageSize, pageSize);
    Logger.log("url to fetch: %s", url);
    var formula = Utilities.formatString("=importhtml(\"%s\", \"table\")", url);
    Logger.log("formula: %s", formula);

    spreadsheet.insertSheet(tsn);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().setFormula(formula);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(rsn), true);
    var range = Utilities.formatString("A%d", (i - 1) * pageSize + 2);
    spreadsheet.getRange(range).activate();
    var srcRange = Utilities.formatString("\'%s\'!B2:L1001", tsn);
    var dstRange = Utilities.formatString("\'%s\'!%s", rsn, range);
    Logger.log("Copy %s to %s", srcRange, dstRange);
    spreadsheet.getRange('\'usatt-temp\'!B2:L1001').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(tsn))
  }
}
