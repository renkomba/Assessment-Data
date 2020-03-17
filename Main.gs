function onEdit(e) {
  let activeSheet = e.source.getActiveSheet();
  let tabs = ['Outline'];
  if (tabs.indexOf(activeSheet.getName()) !== -1) {  // if Outline sheet
    showCreator(activeSheet);
    let columns = [2, 3];
    let cell = e.range;
    let column = cell.getColumn();
    let row = cell.getRow();
    if (columns.indexOf(column) == 0) {  // if column 'B'
      if (row == 2) {  // if new # of formatives
        if (cell != '- - - - -') {
          setupOutline(activeSheet);
        }
      } else if (row == 3) {  // if new colour
        changeColour(activeSheet);
        if (!activeSheet.getRange('D23').isBlank()) {  // if outline already set up
          spawnOutlines(activeSheet, true);  // copy formatting down
        }
      } else if (row == 5) {  // if ready to create
        if (e.value == 'TRUE') {
          setupAssessments();
        }
      }
    } else if (columns.indexOf(column) == 1) {  // if column 'C'
      if (e.value == 'TRUE') {  // if box is checked
        let name = activeSheet.getRange(row, 4).getValue();
        fill(name);
      }
    }
  }
}  // ♚ for corrections and ♟ for failing

let ss = SpreadsheetApp.getActiveSpreadsheet();

// Place my info at top of Outline
function showCreator(outline) {
  let range = outline.getRange('C1');
  let message = 'Made by Rhode N\'komba (renkomba@fcps.edu).' +
      'Do not share or reproduce without permission.';
  range.setFontSize(18).setFontColor('white');
  if (range.getValue() != message) {
    range.setValue(message);
  }
}

// A. create & group tables for each assessment, then hide blank rows.
function setupOutline(outline) {
  ungroup(outline);
  clean(outline);
  spawnOutlines(outline, false)
  hideInOutline(outline);
}

// B. create 'Formative 1' & use as template for other assessment sheets.
function setupAssessments() {
  let sheets = ss.getSheets();
  clearAway(sheets);  // remove every sheet save for 'Outline' & 'Students'
  createFormativeSheets(createF1());
}

// C. fill sheet with formatting, formulas, & question measures
function fill(sheetName) {
  let num = getIndex(sheetName);              // test sheet 'index'
  let outline = matchColumns(sheetName, num); // outline == array of 'Outline' tables
  label(outline, sheetName, num);
  addFormulas(outline, sheetName, num);
}

// D. change colours in 'Outline' sheet
function changeColour(outline) {
  let students = ss.getSheets()[1];
  let colour = findColour(outline);
  console.log('The colour is ${colour}.');
  outline.setTabColor(colour.light);
  students.setTabColor(colour.light);
  setBackgrounds(outline, colour);
  setBorders(outline, colour);
}