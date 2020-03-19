// C1. get an index to pull correctly from arrays for each sheet
function getIndex(sheetName) {
  let outline = ss.getSheets()[0];
  if (sheetName == 'Summative') {
    return outline.getRange('B2').getValue();
  } else {
    return sheetName.slice(10)-1;  // turn formative number into int
  }
}

// C2. get the proper amount of colums
function matchColumns(sheet, num) {
  let outline = getOutlineInfo();
  let neededColumns = outline[num+2][17][6];  // the number of columns needed per test
  let currentColumns = sheet.getMaxColumns(); // find how many columns in the sheet
  let rows = sheet.getMaxRows();
  if (currentColumns != neededColumns) {      // create or delete columns
    let dif = Math.abs(neededColumns - currentColumns);
    if (currentColumns < neededColumns) {     // if not enough columns
      let source = sheet.getRange(1, (currentColumns - dif), rows, dif);
      sheet.insertColumns(currentColumns, dif);
      source.copyTo(sheet.getRange(1, (currentColumns+1), rows, dif));
    } else {  // if too many columns
      sheet.deleteColumns(neededColumns, dif-1);
    }
  }
  neededColumns += 2;
  let range = sheet.getRange(3, 1, rows-1, neededColumns);
  range.createFilter();  // filter all but top 2 rows
  return outline[num+2];
}

// C
function getMeasures(outline) {
  for (let line of outline) {
    if (line[0] == '') {
      let i = outline.indexOf(line);
      return outline.slice(0, i);
    }
  }
}

// C3
function labelHeaders(outline, sheet, num) {
  let numOfSections = outline[16][0];
  let headersAndPoints = [[],[]];
  let sectionLetter = headersAndPoints[0];
  let sectionPoints = headersAndPoints[1];
  let cellInfo = [['M', 'F', 'G', 'H', 'I', 'J', 'E'], [19, 20],
                  [40, 41], [61, 62], [82, 83], [103, 104]];
  outline = '=Outline!';
  for (let i=0; i<=numOfSections; i++) {
    let section, row2;
    let letter = cellInfo[0][i];
    if (i == 0) {
      section = 'Total';
      row2 = cellInfo[1][0];
    } else {
      let row1 = cellInfo[num][0];
      section = outline + letter + row1;
      row2 = cellInfo[num][1]
    }
    let points = outline + letter + row2;
    sectionLetter.push(section);
    sectionPoints.push(points);
  }
  sectionLetter.push('');
  sectionPoints.push('✓');
  if (numOfSections > 1) {
    let remainder = numOfSections % 2;
    let e = cellInfo[0][6];
    let nineteen = cellInfo[1][0];
    let header = outline + e + nineteen;  // 'SECTIONS'
    if (remainder == 0) {  // if an even # of sections
      let column = 6 + (numOfSections / 2);
      sheet.getRange(1, column).setValue(header);
    } else {  // if 3 sections
      sheet.getRange(1, 8).setValue(header);
    }
//  GET ARRAY OF ROW 2 & USE IT TO MAKE BORDERS
//    if (i == (numOfSections - 1)) {  // if last section
//      let afterGraded = sheet.getRange(1, 7+numOfSections, maxRows, 1);
//      afterGraded.setBorder(null, true, null, true, null, null, colour.dark, solidThick);
//    }
    let columns = numOfSections + 2;
    sheet.getRange(2, 6, 2, columns).setValues(headersAndPoints);
  }
}

// A
function labelQuestions(outline, sheet, num) {
  let checkboxColumn = outline[16][6] + 1;
  let numOfSections = outline[16][0];
  let questions = [[],[]];
  let questionSection = questions[0];
  let questionMeasure = questions[1];
  let rows = [4, 25, 46, 67, 88];
  let row = rows[num];
  outline = getMeasures(outline);
  let measureColumns = ['F', 'G', 'H', 'I', 'J'];
  let outlineSheet = '=Outline!';
  for (let i in outline) {  // for each question (index)
    let rowNum = row + i;
    let question = outline[i].slice(0, 7); // section, measure, & total
    let section = question[0];
    let numOfMeasures = question[6];
    let h = i - 1;  // prior index
    if (h < 0) h = 0;
    for (let j=0; j<numOfMeasures; j++) {
      if (questionSection[h] == section) {  // if not new section
        questionSection.push('');
      } else {
        questionSection.push(section);
      }
      let column = measureColumns[j];
      questionMeasure.push(outlineSheet+column+rowNum);
    }
  }
  sheet.getRange(1, checkboxColumn+1, 2, gradingColumns).setValues(questions)
}
  
//// C3. set references to headers in 'Outline' & add borders
//function label(outline, sheet, num) {
//  let gradingColumns = outline[15][6];
//  let checkboxColumn = outline[16][6]+1;
//  let numOfSections = outline[16][0];
//  let rows = [[19, 20, 21, 4], [40, 41, 42, 25],  // 0-section headers, 1-section points
//             [61, 62, 63, 46], [82, 83, 84, 67],  // 2-bonus per section & total points
//             [103, 104, 105, 88]];  // 3-first row of measures
//  let cl = ['E','F', 'G', 'H', 'I', 'J'];  // section column letters
//  let colour = findColour(ss.getSheets()[0]);
//  let lineIndex = 0;
//  let columnIndex = 1;
//  let section;
//  let maxRows = sheet.getMaxRows();
//  let double = SpreadsheetApp.BorderStyle.DOUBLE;
//  let dashed = SpreadsheetApp.BorderStyle.DASHED;
//  let solidThick = SpreadsheetApp.BorderStyle.SOLID_THICK;
//  sheet.getRange(3, checkboxColumn).setValue('✓');
//  sheet.getRange(4, checkboxColumn, maxRows-3, 1).insertCheckboxes();
//  if (numOfSections > 1) { // if more than one section, else overall is all we need
//    for (let line of outline) {
//      line = line.slice(0, 6);
//      let column = checkboxColumn + columnIndex;
//      let row = rows[num][3] + lineIndex;
//      let sectionLabels = [];
//      if (line[0] == '' || line[0] == 'SECTIONS') {
//        break;
//      } else {
//        for (let m=0; m<6; m++) {
//          if (m == 0) {
//            if (section != line[m]) {  // if it's a new section
//              sectionLabels.push(column);
//              section = line[m];
//              sheet.getRange(1, column, maxRows)
//              .setBorder(null, true, null, null, null, null, colour.dark, double);
//              questions[0].push(line[m]);  // add section letter to array
//            } else {
//              questions[0].push('');
//            }
//          } else if (line[m] == '') {  // or if measure is blank
//            if (!sectionLabels.includes(column)) {
//              sheet.getRange(1, column, maxRows)
//              .setBorder(null, true, null, null, null, null, colour.dark, dashed);
//              }
//            lineIndex++;
//            break;  // move to the next line
//          } else {  // if not the first measure
//            if (m == 5) {  // if last one
//              lineIndex++;
//              sheet.getRange(1, column, maxRows)
//              .setBorder(null, null, null, true, null, null, colour.dark, double);
//            } else if (m > 1) {
//              questions[0].push('');
//            }
//            questions[1].push('=Outline!'+cl[m]+row);
//            columnIndex++;
//          }
//        }
//      }
//    }
//    sheet.getRange(1, checkboxColumn+1, 2, gradingColumns).setValues(questions);
//  }
//}

// C4. set up conditional formatting & add formulas
function addFormulas(outline, sheet) {
  let numOfSections = outline[16][0];
  let firstQuestion = outline[16][6] + 2;  // where the grading starts
  let allQuestions = outline[15][6];  // how many columns of grading
  let sections = outline[15].slice(1, 5);  // letters
  let points = outline[16].slice(1, 5); // points
  let bonus = outline[17].slice(1, 5);  // bonus
  let checkbox = sheet.getRange(4, (firstQuestion-1)).getA1Notation();  // checkbox cell
  console.log("Checkbox starts at " + checkbox);
//  let range = sheet.getRange(4, firstQuestion, rows, allQuestions);  // grading columns
//  range.setHorizontalAlignment('center');  // NOT NECESSARY IF C2 COPIES FORMATTING
  let fullGradingRange = sheet.getRange(4, firstQuestion, 1, allQuestions).getA1Notation();
  let condition = '=if('+checkbox+' = true; ';  // beginning of formula
  let countIf = 'countif(';
  let blank = '\"\")';
  let columnIndex = 7;  // index of first section letter
  let sectionTotals = ['G$3; ', 'H$3; ', 'I$3; ', 'J$3; '];
  let formulae = [[]];
  formulae[0].push(condition+countIf+fullGradingRange+'; '+blank+'/ F$3; '+ blank);
  for (let i in sections) {  // for each section letter
    if (sections[i] == '') { // if there is no letter
      break;
    } else {
      let sectionAverageCell = sheet.getRange(4, columnIndex).getA1Notation();
      let sectionSpan = points[i] + bonus[i];
      let sectionQuestions = sheet.getRange(4, firstQuestion, 1, sectionSpan);
      let ifTrue = countIf+sectionQuestions+'; '+blank+' / '+ sectionTotals[i];
      formulae[0].push(condition+ifTrue+blank);
      firstQuestion += sectionSpan;
      columnIndex++;
    }
  }
  sheet.getRange(4, 6, 1, numOfSections+1).setFormulas(formulae);
}

//// C5. setup or modify conditional formatting
function formatOnCondition(outline, sheet, alreadyFilled) {
  let rules = sheet.getConditionalFormatRules();
  if (alreadyFilled) {  // if sheet was already filled
    let sheetName = sheet.getSheetName();
    outline = outline[getIndex(sheetName)];
    sheet.clearConditionalFormatRules();
  }
  let firstQuestion = outline[16][6] + 2;  // where the grading starts
  let allQuestions = outline[15][6];  // how many columns of grading
  let checkbox = sheet.getRange(4, (firstQuestion-1)).getA1Notation();
  let colour = findColour(ss.getSheets()[0]);
  let rows = sheet.getMaxRows() - 3;
  let questionRange = sheet.getRange(4, firstQuestion, rows, allQuestions);
  let rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$'+checkbox+' = FALSE')  // if box is not checked
      .setBackground(colour.dark)
      .setFontColor(colour.dark)
      .setRanges([questionRange])
      .build();
  let rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()  // if text in cell (for mistake)
      .setBackground(colour.dark)
      .setFontColor('white')
      .setRanges([questionRange])
      .build();
  rules.push(rule1);
  rules.push(rule2);
  sheet.setConditionalFormatRules(rules);
}

// C5. format filled sheet
function formatFilledSheet(outline, sheet) {
  let numOfSections = outline[16][0];
  sheet.autoResizeColumns(6, sheet.getMaxColumns()-5);
  sheet.setColumnWidths(6, numOfSections+1, 60);
  let rows = sheet.getMaxRows();
  let firstStudentFormulas = sheet.getRange(4, 6, 1, numOfSections+1);
  let allStudentFormulas = sheet.getRange(5, 6, rows-4, numOfSections+1)
  sheet.getRange(4, 6, rows-3, numOfSections+1).setNumberFormat('0.00%');
  firstStudentFormulas.copyTo(allStudentFormulas);
}