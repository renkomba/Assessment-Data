var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var outline = sheets[0];

// B1. Delete sheets
function deleteSheets() {
  clearAway(sheets);
  ss.getSheetByName("Outline").activate();
  cleanOutline();
}

// A2. group rows
function cleanOutline() {
  ungroup(outline);
  var blankRange = outline.getRange("C23:M106");  // everything below the first table
  var rows = outline.getMaxRows();
  outline.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  outline.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  outline.showRows(1, rows);  // unhide all rows
  blankRange.clear();  // clear the values and formatting of the blank range
  blankRange.removeCheckboxes();
  outline.hideRows(23, rows-22);
}