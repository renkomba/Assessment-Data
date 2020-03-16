function onEdit(e) {
  var activeSheet = e.source.getActiveSheet();
  var tabs = ["Outline"];
  if (tabs.indexOf(activeSheet.getName()) !== -1) {  // if Outline sheet
    copyright();
    var columns = [2, 3];
    var cell = e.range;
    var col = cell.getColumn();
    var row = cell.getRow();
    if (columns.indexOf(col) == 0) {
      if (row == 2) {  // if new # of formatives
        if (cell != "- - - - -") {
          console.log("New number of formatives. Creating outline and sheet...")
          setupOutline();
          activeSheet.getRange("B3").setValue("- - - - -");
        }
      } else if (row == 3) {
        changeColour();
      } else if (row == 5) {
        if (e.value == "TRUE") {
          setupAssessments();
        }
      }
    } else if (columns.indexOf(col) == 1) {  // if box checked
      if (e.value == "TRUE") {
        var name = activeSheet.getRange(row, 4).getValue();
        console.log("Filling out "+name+"...");
        fill(name);
      }
    }
  }
}  //♚ for corrections and ♟ for failing

var ss = SpreadsheetApp.getActiveSpreadsheet();

// Place my info
function copyright() {
  var outline = ss.getSheets()[0];
  var range = outline.getRange("C1");
  range.setFontSize(18).setFontColor("white");
  if (range.getValue() != message) {
    var message = "Made by Rhode N'komba (renkomba@fcps.edu). Do not share or reproduce without permission.";
    range.setValue(message);
  }
}

// A. Set up the "Assessments" sheet
function setupOutline() {
  var outline = ss.getSheets()[0];
  ungroup(outline);
  clean(outline);
  spawnOutlines(outline)
  hideInOutline(outline);
}

// B. create assessment sheets
function setupAssessments() {
  var sheets = ss.getSheets();
  clearAway(sheets);  // remove existing formative/summative sheets
  createFormativeSheets(createF1());  // create F1, duplicate it, and rename it for each sheet
}

// C. fill sheet
function fill(sheetName) {
  var num = getIndex(sheetName);
  var outline  = matchColumns(sheetName, num);
  label(outline, sheetName, num);
  addFormulas(outline, sheetName, num);
}

// D. change sheet colours
function changeColour() {
  var outline = ss.getSheets()[0];
  var students = ss.getSheets()[1];
  var colour = findColour(outline);
  outline.setTabColor(colour.light);
  students.setTabColor(colour.light);
  setBackgrounds(outline, colour);
  setBorders(outline, colour);
}