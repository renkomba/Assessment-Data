var ss = SpreadsheetApp.getActiveSpreadsheet();

// C1. get an index to pull correctly from arrays for each sheet
function getIndex(sheetName) {
  var outline = ss.getSheets()[0];
  if (sheetName == "Summative") {
    return outline.getRange("B2").getValue();
  } else {
    return sheetName.slice(10)-1;  // turn formative number into int
  }
}

// C2. Get the proper amount of colums
function matchColumns(sheetName, num) {
  var sheet = ss.getSheetByName(sheetName);
  var outline = getOutlineInfo();
  var neededColumns = outline[num+2][17][6];  // the number of columns needed per test
  var currentColumns = sheet.getMaxColumns(); // find how many columns in the sheet
  if (currentColumns != neededColumns) {      // create or delete columns
    var dif = Math.abs(neededColumns - currentColumns);
    if (currentColumns < neededColumns) {     // if not enough columns
      var copyRange = sheet.getRange(1, (currentColumns - dif), sheet.getMaxRows(), dif);
      sheet.insertColumns(currentColumns, dif);  // add needed amount at end
      copyRange.copyTo(sheet.getRange(1, (currentColumns+1), sheet.getMaxRows(), dif)); // copy format onto new additions
    } else {  // if too many columns
      sheet.deleteColumns(neededColumns, dif-1);
    }
  }
  neededColumns = sheet.getMaxColumns() + 1;
  var range = sheet.getRange(3, 1, sheet.getMaxRows()-1, neededColumns);
  range.createFilter();
  return outline[num+2];
}

// C3. set all headers and make groups
function label(outline, sheetName, num) {
  var sheet = ss.getSheetByName(sheetName);
  var gradingColumns = outline[15][6];
  var checkboxColumn = outline[16][6]+1;
  var numOfSections = outline[16][0];
  var headersAndPoints = [[],[]];  // i-0 is section letter, i-1 is point value
  var questions = [[],[]];
  var rows = [[19, 20, 21, 4], [40, 41, 42, 25],  // 0-section headers, 1-section points
             [61, 62, 63, 46], [82, 83, 84, 67],  // 2-bonus per section & total points
             [103, 104, 105, 88]];  // 3-first row of measures
  var cl = ["E","F", "G", "H", "I", "J"];  // section column letters
  var colour = findColour(ss.getSheets()[0]);
  var lineIndex = 0;
  var colIndex = 1;
  var section;
  sheet.getRange(3, checkboxColumn).setValue("✓");
  sheet.getRange(4, checkboxColumn, sheet.getMaxRows()-3, 1).insertCheckboxes();
  if (numOfSections > 1) { // if more than one section, else overall is all we need
    for (var line of outline) {
      line = line.slice(0, 6);
      var column = checkboxColumn + colIndex;
      var row = rows[num][3] + lineIndex;
      var sectionLabels = [];
      if (line[0] == "" || line[0] == "SECTIONS") {
        break;
      } else {
        for (var m=0; m<6; m++) {
          if (m == 0) {
            if (section != line[m]) {  // if it's a new section
              sectionLabels.push(column);
              section = line[m];
              sheet.getRange(1, column, sheet.getMaxRows())
              .setBorder(null, true, null, null, null, null, colour.dark, SpreadsheetApp.BorderStyle.DOUBLE);
              questions[0].push(line[m]);  // add section letter to array
            } else {
              questions[0].push("");
            }
          } else if (line[m] == "") {  // or if measure is blank
            if (!sectionLabels.includes(column)) {
              sheet.getRange(1, column, sheet.getMaxRows())
              .setBorder(null, true, null, null, null, null, colour.dark, SpreadsheetApp.BorderStyle.DASHED);
              }
            lineIndex++;
            break;  // move to the next line
          } else {  // if not the first measure
            if (m == 5) {  // if last one
              lineIndex++;
              sheet.getRange(1, column, sheet.getMaxRows())
              .setBorder(null, null, null, true, null, null, colour.dark, SpreadsheetApp.BorderStyle.DOUBLE);
            } else if (m > 1) {
              questions[0].push("");
            }
            questions[1].push("=Outline!"+cl[m]+row);
            colIndex++;
          }
        }
      }
    }
    for (var i=0; i<numOfSections; i++) {  // for each section
      headersAndPoints[0].push("=Outline!"+cl[i+1]+rows[num][0]);  // add reference to section letter to array
      headersAndPoints[1].push("=Outline!"+cl[i+1]+rows[num][1]);  // same for section points
      if (i == (numOfSections - 1)) {  // if last section
        var afterGraded = sheet.getRange(1, 7+numOfSections, sheet.getMaxRows(), 1);  // insert dark magenta border
        afterGraded.setBorder(null, true, null, true, null, null, "#4c1130", SpreadsheetApp.BorderStyle.SOLID_THICK);
      }
    }
  }
  sheet.getRange(2, 7, 2, numOfSections).setValues(headersAndPoints);
  sheet.getRange(1, checkboxColumn+1, 2, gradingColumns).setValues(questions);
  sheet.getRange(3, 6).setValue("=Outline!M"+rows[num][0]);
}

// C4. 
function addFormulas(outline, sheetName, num) {
  var sheet = ss.getSheetByName(sheetName);
  var numOfSections = outline[16][0];
  var start = outline[16][6] + 2;
  var span = outline[15][6];
  var end = start + span;
  var sections = outline[15].slice(1, 5);
  var secPoints = outline[16].slice(1, 5);
  var secBonus = outline[17].slice(1, 5);
  var cl = [0,"A","B","C","D","E","F","G","H","I","J",
           "K","L","M","N","O","P","Q","R","S","T","U",
           "V","W","X","Y","Z","AA","AB","AC","AD","AE",
           "AF","AG","AH","AI","AJ","AK","AL","AM","AN",
           "AO","AP","AQ","AR","AS","AT","AU","AV","AW",
           "AX","AY","AZ"];
  var cb = cl[start-1];  // checkbox column letter
  var cbRef = cb+"4"; // checkbox cell
  var colIndex = 7;
  var formulae = [[]];
  var rules = sheet.getConditionalFormatRules();
  var range = sheet.getRange(4, start, sheet.getMaxRows()-3, span);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$"+cbRef+" = FALSE")
      .setBackground("#4c1130")
      .setFontColor("#4c1130")
      .setRanges([range])
      .build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground("#4c1130")
      .setFontColor("white")
      .setRanges([range])
      .build();
  rules.push(rule1);
  rules.push(rule2);
  range.setHorizontalAlignment("center");
  formulae[0].push("=if("+cbRef+"=true; countif("+cl[start]+"4:"+cl[end]+"4; \"\") / F$3; \"\")");
  for (var i in sections) {  // for each index in "sections" array
    if (sections[i] == "") {
      break;
    } else {
      var col = cl[colIndex];
      var secSpan = secPoints[i] + secBonus[i] - 1;  // calc span of section
      end = cl[start + secSpan] + "4";  // create final cell ref string
      var calcRef = cl[start] + "4:" + end;  // create range reference string
      var ifTrue = "countif("+calcRef+ "; \"\") / "+col+"$3";
      formulae[0].push("=if("+cbRef+"=true; "+ifTrue+"; \"\")");
      start += secSpan + 1;  // setup next section start
      colIndex++;
    }
  }
  sheet.getRange(4, 6, 1, numOfSections+1).setFormulas(formulae);
  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, sheet.getMaxColumns());
  if (numOfSections > 1) { // if more than one section, else overall is all we need
    var remainder = numOfSections % 2;
    if (remainder == 0) {  // if an even # of sections
      sheet.getRange(1, (6+(numOfSections/2))).setValue("=Outline!e19");  // write "SECTIONS" in first row
    } else {  // if 3 sections
      sheet.getRange(1, 8).setValue("=Outline!e19");
    }
  }
  var rows = sheet.getMaxRows();
  var copyFrom = sheet.getRange(4, 6, 1, numOfSections+1);
  var copyTo = sheet.getRange(5, 6, rows-4, numOfSections+1)
  sheet.setColumnWidths(6, numOfSections+1, 50);
  sheet.getRange(4, 6, rows-3, numOfSections+1).setNumberFormat("0.00%");
  copyFrom.copyTo(copyTo);
}