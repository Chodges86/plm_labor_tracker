let ss;
let sheet;
let source0;
let source1;

function getSpreadsheets() {
  //Call Spreadsheet and get it//
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = SpreadsheetApp.getActiveSheet();
  source0 = ss.getSheets()[0];
  source1 = ss.getSheets()[1];
}

function printEventsDetails(events) {
  formatSheet();
  printEvents(events);
  removeCalculationFormulas();
  // analyzeDescriptions() Function found in SD_View.gs file.  Function not finished
} // END OF: printEventsDetails function

// SHOW THE CUSTOM DIALOGUE.  Called from the filter button on the spreadsheet
function showFilterDialogue() {
  // This passes a string value of "Location" or "Title" back to the filter() function through a script function in the FilterAlert.html file
  showDialog("SD_FilterAlert", 400, 300, "Filter");
} // END OF showCustomFilterDialogue function

function didOpen() {
  // TODO: Check if the getSpreadsheets call below takes care of not needing it in each function
  getSpreadsheets();
  source1.getRange(2, 4).setValue("Enter Year (yyyy)");
  source1.getRange(2, 3).setValue("");
}
