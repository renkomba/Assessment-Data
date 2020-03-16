var ss = SpreadsheetApp.getActiveSpreadsheet();

// B1. Delete sheets
function clearAway(sheets) {
  if (sheets.length > 2) {  // if there are more than 2 sheets
    for (var i=2; i<sheets.length; i++) {  // starting at index 2 (3rd sheet)
      ss.deleteSheet(sheets[i]);  // delete the sheet
    }
  }
}

// B2. Create Formative 1 Page
function createF1() {
  var f1Title = getOutlineInfo()[2][18][0];
  ss.insertSheet(f1Title, 2);  // insert sheet called "F1" after "Students"
  var sheets = ss.getSheets();
  populatef1(sheets);
  formatSheet(sheets);
  return sheets;
}

// B3. Populate Formative 1 sheet
function populatef1(sheets) { 
  var students = sheets[1];
  var f1 = sheets[2];
  // 0 (per, sec num, end index), 1 (same w/ start index), 2 (irrelevant), 3-end (student info)
  var rangeToCopy = sheets[1].getRange("A1").getDataRegion().getValues();  // student info
  var classInfo = students.getRange("H1").getDataRegion().getValues();  // teachers, students per class
  var rowNum = 4;
  var studentInfo = rangeToCopy.slice(1);
  var totClassLen = sheets[1].getLastRow()-1;
  var maxRows = f1.getMaxRows();
  var maxCol = f1.getMaxColumns();
  var copyToRange = f1.getRange(4, 1, totClassLen, 5);
  ss.setActiveSheet(f1);  // set "Formative 1" as active sheet
  copyToRange.setValues(studentInfo);
  f1.setFrozenColumns(6);  // freeze student info and % column
  f1.setFrozenRows(3);  // freeze all rows with header
  f1.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  f1.getRange(1, 4, maxRows, 2).shiftColumnGroupDepth(3);
  f1.getRange(1, 3, maxRows).shiftColumnGroupDepth(2);
  f1.getRange(1, 1, maxRows, 2).shiftColumnGroupDepth(1);
  classInfo.forEach(function(row) { // for each teacher
    var i = 2;                 // index of first prep
    var preps = row[1] + i;    // classes per teacher
    for (i; i < preps; i++) {  // for each prep
      if (row[i]) {            // if not blank
        var period = row[i];   // class period
        var classLen = row[i+5];  // # of students
        var prep = f1.getRange(rowNum, 1, classLen, maxCol);
        rowNum += classLen;
        prep.shiftRowGroupDepth(i-1);
      }
    }
  })
  var emptyRow = totClassLen+4;
  f1.hideRows(emptyRow, maxRows-emptyRow+1);
}

// B4. Format sheet
function formatSheet(sheets) {
  var sheetName = sheets[0].getRange("D2").getValue();
  var f1 = ss.getSheetByName(sheetName);
  var rows = f1.getMaxRows();
  var headers = ["Teacher", "P.", "ID", "First", "Last", "Total"];
  var columns = f1.getMaxColumns();
  var colour = findColour(ss.getSheets()[0]);
  var bandingRange = f1.getRange(4, 1, f1.getLastRow(), columns);
  f1.getRange(1, 1, 3, columns)  // change first 3 rows
    .setFontColor("white")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(colour.plain);
  f1.getRange(2, 6, 1, columns)
    .setBackground(colour.deep);  // change part of 2nd row deep colour
  f1.getRange(3, 6, 1, columns)
    .setBackground(colour.dark);  // change part of 3rd row dark colour
  f1.getRange(3, 1, 1, 5).setFontSize(10).setBackground(colour.deep);  // change 3rd row
  if (colour.name == "red") {
    bandingRange.clearFormat()
    .applyRowBanding(pink.banding, false, false)
    .setFirstRowColor("#fddcdc");
  } else {
    bandingRange.applyRowBanding(colour.banding, false, false);
  }
  f1.getRange(4, 22, rows, 5).copyTo(f1.getRange(4, 27, rows, 5));
  columns = f1.getMaxColumns();
  f1.getRange(1, 1, rows, columns)
    .setHorizontalAlignment("center");  // center all text
  f1.getRange(1, 4, rows, 2)
    .setHorizontalAlignment("left");  // except in D:E
  for (var i=1; i<headers.length; i++) {
    f1.getRange(3, i).setValue(headers[i-1]);  // get column-3 and set it to headers[i-1]
  }
  f1.getRange(2, 6).setValue(headers[5]);
  f1.autoResizeColumns(1, 5);
  f1.setTabColor(colour.plain);
}

// B5. Create formative and summative calculation sheets based on info in the "Assessments" sheet
function createFormativeSheets() {
  var sheets = ss.getSheets();
  ss.setActiveSheet(sheets[2]);
  var colour = findColour(sheets[0]);
  var names = sheets[0].getRange("A8").getDataRegion(SpreadsheetApp.Dimension.ROWS)
                       .getValues().slice(6);
  names.forEach(function(name) {
    var len = sheets.length;
    SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();  // duplicate "F1"
    sheets = ss.getSheets();
    var sheet = sheets[len];  // get new duplicated sheet
    sheet.setName(name);  // name it the table title
    if (name == "Summative") {
      sheet.setTabColor(colour.dark);
    }
  })
  ss.getSheetByName("Outline").activate();
}