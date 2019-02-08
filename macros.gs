function SortingFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(2, true);
  spreadsheet.getRange('D1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(4, false);  
};