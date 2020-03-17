// B1. delete all sheets save for 'Outline' and 'Students'
function clearAway(sheets) {
  if (sheets.length > 2) {  // if there are more than 2 sheets
    for (let i=2; i<sheets.length; i++) {  // starting at index 2 (3rd sheet)
      ss.deleteSheet(sheets[i]);  // delete the sheet
    }
  }
}

// B2. create sheet called 'Formative 1'
function createF1() {
  let f1Title = getOutlineInfo()[2][18][0];  // 'Formative 1'
  ss.insertSheet(f1Title, 2);
  let sheets = ss.getSheets();
  populateSheet(sheets);
  formatSheet(sheets[2]);
  return sheets;
}

// B3. paste student info in sheet, make sheet active, freeze & group columns
function populateSheet(sheets) { 
  let students = sheets[1];
  let f1 = sheets[2];
  let rangeToCopy = students.getRange('A1').getDataRegion().getValues();
  let studentInfo = rangeToCopy.slice(1);  // remove header row
  let totalStudents = students.getLastRow() - 1;
  let copyToRange = f1.getRange(4, 1, totalStudents, 5);
  ss.setActiveSheet(f1);   // set 'Formative 1' as active sheet
  copyToRange.setValues(studentInfo);
  f1.setFrozenColumns(6);  // freeze student info and % column
  f1.setFrozenRows(3);     // freeze all header rows
  f1.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  let rows = f1.getMaxRows();
  f1.getRange(1, 4, rows, 2).shiftColumnGroupDepth(3);  // group first & last names
  f1.getRange(1, 3, rows).shiftColumnGroupDepth(2);     // group student ID
  f1.getRange(1, 1, rows, 2).shiftColumnGroupDepth(1);  // group teacher & period
  let firstEmptyRow = totalStudents + 4;
  let allEmptyRows = rows - firstEmptyRow + 1;
  f1.hideRows(firstEmptyRow, allEmptyRows);
}

// B4. insert formatting, resize columns, & set tab colour 
function formatSheet(sheet) {
  let headers = [['Teacher', 'P.', 'ID', 'First', 'Last', 'Total']];
  let columns = sheet.getMaxColumns();
  let colour = findColour(ss.getSheetByName('Outline'));
  let defaultBanding = SpreadsheetApp.BandingTheme.LIGHT_GREY;
  let bandingRange = sheet.getRange(4, 1, sheet.getLastRow(), columns);
  bandingRange.applyRowBanding(defaultBanding, false, false)
              .setSecondRowColor(colour.lightest);
  sheet.getRange(1, 1, 3, columns)  // all columns, only first 3 rows
       .setFontColor('white')
       .setFontSize(12)
       .setFontWeight('bold')
       .setBackground(colour.plain);
  sheet.getRange(2, 6, 1, columns)  // all columns from 'total', 2nd row
       .setBackground(colour.deep);
  sheet.getRange(3, 6, 1, columns)  // all columns from 'total', 3rd row
       .setBackground(colour.dark);
  let total = headers[0].splice(5, 1);      // remove 'total' from headers & store
  sheet.getRange(3, 1, 1, 5)        // column of headers, 3rd row
       .setFontSize(10)
       .setBackground(colour.deep)
       .setValues(headers);
  let rows = sheet.getMaxRows();
  sheet.getRange(4, 22, rows, 5).copyTo(sheet.getRange(4, 27, rows, 5));
  columns = sheet.getMaxColumns();
  sheet.getRange(1, 1, rows, columns)    // all text
       .setHorizontalAlignment('center');
  sheet.getRange(1, 4, rows, 2)          // except for names (D:E)
       .setHorizontalAlignment('left');
  sheet.getRange('F2').setValue(total);
  sheet.autoResizeColumns(1, 5);         // resize header columns (first 5)
  sheet.setTabColor(colour.deep);
}

// B5. Create formative and summative sheets based on info in 'Outline'
function createFormativeSheets(sheets) {
  ss.setActiveSheet(sheets[2]);
  let colour = findColour(sheets[0]);
  let names = sheets[0].getRange('A8').getDataRegion(SpreadsheetApp.Dimension.ROWS)
                       .getValues().slice(6);  // assessments after 'Formative 1'
  names.forEach(function(name) {
    let len = sheets.length;
    SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();  // duplicate 'F1'
    sheets = ss.getSheets();
    let sheet = sheets[len];  // get new duplicated sheet
    sheet.setName(name);  // name it the table title
    if (name == 'Summative') sheet.setTabColor(colour.dark);
  })
  ss.getSheetByName('Outline').activate();
}