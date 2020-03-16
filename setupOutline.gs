var ss = SpreadsheetApp.getActiveSpreadsheet();

// A1. ungroup rows
function ungroup(outline) {
  var rows = [23, 44, 65, 86];
  rows.forEach(function(row) {
    var trigger = outline.getRange(row, 1).getValue();
    var title = outline.getRange(row+1, 3).getValue();
    if (trigger == "grouped" && title == "#") {
      var group = outline.getRowGroup(row+1, 1);
      group.remove();
      outline.getRange(row, 1).clear();
    }
  })
}

// A2. group rows
function clean(outline) {
  var blankRange = outline.getRange("C23:M106");  // clear everything below the first table
  var rows = outline.getMaxRows();
  outline.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  outline.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  outline.showRows(1, rows);  // unhide all rows
  blankRange.clear();  // clear the values and formatting of the blank range before writing onto them
  blankRange.removeCheckboxes();
}

// A3. spawn assessment outlines
function spawnOutlines(outline, format) {
  var info = outline.getRange("B2:D2").getValues();
  var row = 23;
  var testAmt = info[0][0]; // # of outline
  var header = info[0][2];  // get the "Formative 1" header;
  var rangeToCopy = outline.getNamedRanges()[0].getRange();  // copy C2:M22
  for (var i=0; i<testAmt; i++) {  // run the following code one fewer times than there are formatives
    if (format) {
      var bandingRange = outline.getRange((row+2), 3);
      var banding = bandingRange.getBandings()[0];
      banding.remove();
      rangeToCopy.copyFormatToRange(outline, 3, 13, row, (row+20));
    } else {
      rangeToCopy.copyTo(outline.getRange(row, 3, 20, 11));  // copy to C23:M43, or C44...
      var borderRange = outline.getRange(row, 3, 1, 11);
      var test = outline.getRange(row+1, 3, 20, 11);
      test.shiftRowGroupDepth(1);
      outline.getRange(row, 1).setValue("grouped").setFontColor("white");
      if (i == (testAmt-1)) {
        outline.getRange(row, 4).setValue("Summative");
      } else {
        outline.getRange(row, 4).setValue(header.replace("1",(i+2)));  // set header of copied cell and change the 1 to 2, 3, etc.
      }
    }
    row += 21;  // row == 44, 65, etc.
  }
}

// A4. hide blank rows
function hideInOutline(outline) {
  var br = outline.getLastRow();  // bottom row
  var rows = outline.getMaxRows(); // total rows
  outline.hideRows(br+1, rows-br);
}