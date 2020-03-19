// B1. delete all sheets save for 'Outline' and 'Students'
function clearAway(sheets) {
  if (sheets.length > 2) {  // if there are more than 2 sheets
    for (let i=2; i<sheets.length; i++) {  // starting at index 2 (3rd sheet)
      ss.deleteSheet(sheets[i]);  // delete the sheet
    }
  }
}

// B2. create sheet called 'Formative 1'
function createF1(alreadyCreated) {
  let f1Title = getOutlineInfo()[2][18][0];  // 'Formative 1'
  ss.insertSheet(f1Title, 2);
  let sheets = ss.getSheets();
  populateSheet(sheets);
  formatSheet(sheets[2], alreadyCreated);
  return sheets;
}

// B3. paste student info in sheet, make sheet active, freeze & group columns
function populateSheet(sheets) { 
  let students = sheets[1];
  let f1 = sheets[2];
  let rangeToCopy = students.getRange('A1').getDataRegion().getValues();
  let studentInfo = rangeToCopy.slice(1);  // remove header row
  let totalStudents = students.getLastRow() - 1;
  let destination = f1.getRange(4, 1, totalStudents, 5);
  ss.setActiveSheet(f1);   // set 'Formative 1' as active sheet
  destination.setValues(studentInfo);
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
function formatSheet(sheet, alreadyCreated) {
  let columns = sheet.getMaxColumns();
  let colour = findColour(ss.getSheetByName('Outline'));
  let lastRow = sheet.getLastRow();
  let bandingRange = sheet.getRange(4, 1, lastRow, columns);
  let first3Rows = sheet.getRange(1, 1, 3, columns);
  let row2FromColumn6 = sheet.getRange(2, 6, 1, columns); 
  let row3FromColumn6 = sheet.getRange(3, 6, 1, columns);
  let headersRange = sheet.getRange(3, 1, 1, 5);
  first3Rows.setBackground(colour.plain);
  row2FromColumn6.setBackground(colour.deep);
  row3FromColumn6.setBackground(colour.dark);
  headersRange.setBackground(colour.deep);
  if (alreadyCreated) {
    bandingRange.setSecondRowColor(colour.lightest);
  } else {
    let headers = [['Teacher', 'P.', 'ID', 'First', 'Last']];  // nested for setValues()
    let defaultBanding = SpreadsheetApp.BandingTheme.LIGHT_GREY;
    bandingRange.applyRowBanding(defaultBanding, false, false)
    .setSecondRowColor(colour.lightest);
    first3Rows.setFontColor('white')
    .setFontSize(12)
    .setFontWeight('bold')
    headersRange.setFontSize(10)
    .setValues(headers);
    let rows = sheet.getMaxRows();
//    sheet.getRange(4, 22, rows, 5).copyTo(sheet.getRange(4, 27, rows, 5)); NECESSARY WITH MATCHROWS() ?
//    columns = sheet.getMacColumns();
    sheet.getRange(1, 1, rows, columns)  // all cells
    .setHorizontalAlignment('center');
    sheet.getRange(1, 4, rows, 2)        // except for names (D:E)
    .setHorizontalAlignment('left');
    sheet.autoResizeColumns(1, 5);       // resize header columns (first 5)
    sheet.setColumnWidth(2, 20);         // widen peiod column to better see filter
    sheet.setTabColor(colour.deep);
  }
}

// B5. Create formative and summative sheets based on info in 'Outline'
function createFormativeSheets(sheets) {
  ss.setActiveSheet(sheets[2]);
  let colour = findColour(sheets[0]);
  let names = sheets[0].getRange('A8:A11').getValues();  // assessments after 'Formative 1'
  for (let name of names) {
    console.log('Name: ' + name + ' of ' + names);
    if (name !== '') {
      let len = sheets.length;
      SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();  // duplicate 'F1'
      sheets = ss.getSheets();
      let sheet = sheets[len];  // get new duplicated sheet
      sheet.setName(name);  // name it the table title
      if (name == 'Summative') sheet.setTabColor(colour.dark);
    }
  }
  ss.getSheetByName('Outline').activate();
}