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
function matchColumns(sheetName, num) {
  let sheet = ss.getSheetByName(sheetName);
  let outline = getOutlineInfo();
  let neededColumns = outline[num+2][17][6];  // the number of columns needed per test
  let currentColumns = sheet.getMaxColumns(); // find how many columns in the sheet
  let rows = sheet.getMaxRows();
  if (currentColumns != neededColumns) {      // create or delete columns
    let dif = Math.abs(neededColumns - currentColumns);
    if (currentColumns < neededColumns) {     // if not enough columns
      let copyRange = sheet.getRange(1, (currentColumns - dif), rows, dif);
      sheet.insertColumns(currentColumns, dif);
      copyRange.copyTo(sheet.getRange(1, (currentColumns+1), rows, dif));
    } else {  // if too many columns
      sheet.deleteColumns(neededColumns, dif-1);
    }
  }
  neededColumns += 2;
  let range = sheet.getRange(3, 1, rows-1, neededColumns);
  range.createFilter();  // filter all but top 2 rows
  return outline[num+2];
}

// C3. set references to headers in 'Outline' & add borders
function label(outline, sheetName, num) {
  let sheet = ss.getSheetByName(sheetName);
  let gradingColumns = outline[15][6];
  let checkboxColumn = outline[16][6]+1;
  let numOfSections = outline[16][0];
  let headersAndPoints = [[],[]];  // i-0 is section letter, i-1 is point value
  let questions = [[],[]];
  let rows = [[19, 20, 21, 4], [40, 41, 42, 25],  // 0-section headers, 1-section points
             [61, 62, 63, 46], [82, 83, 84, 67],  // 2-bonus per section & total points
             [103, 104, 105, 88]];  // 3-first row of measures
  let cl = ['E','F', 'G', 'H', 'I', 'J'];  // section column letters
  let colour = findColour(ss.getSheets()[0]);
  let lineIndex = 0;
  let columnIndex = 1;
  let section;
  let maxRows = sheet.getMaxRows();
  let double = SpreadsheetApp.BorderStyle.DOUBLE;
  let dashed = SpreadsheetApp.BorderStyle.DASHED;
  let solidThick = SpreadsheetApp.BorderStyle.SOLID_THICK;
  sheet.getRange(3, checkboxColumn).setValue('✓');
  sheet.getRange(4, checkboxColumn, maxRows-3, 1).insertCheckboxes();
  if (numOfSections > 1) { // if more than one section, else overall is all we need
    for (let line of outline) {
      line = line.slice(0, 6);
      let column = checkboxColumn + columnIndex;
      let row = rows[num][3] + lineIndex;
      let sectionLabels = [];
      if (line[0] == '' || line[0] == 'SECTIONS') {
        break;
      } else {
        for (let m=0; m<6; m++) {
          if (m == 0) {
            if (section != line[m]) {  // if it's a new section
              sectionLabels.push(column);
              section = line[m];
              sheet.getRange(1, column, maxRows)
              .setBorder(null, true, null, null, null, null, colour.dark, double);
              questions[0].push(line[m]);  // add section letter to array
            } else {
              questions[0].push('');
            }
          } else if (line[m] == '') {  // or if measure is blank
            if (!sectionLabels.includes(column)) {
              sheet.getRange(1, column, maxRows)
              .setBorder(null, true, null, null, null, null, colour.dark, dashed);
              }
            lineIndex++;
            break;  // move to the next line
          } else {  // if not the first measure
            if (m == 5) {  // if last one
              lineIndex++;
              sheet.getRange(1, column, maxRows)
              .setBorder(null, null, null, true, null, null, colour.dark, double);
            } else if (m > 1) {
              questions[0].push('');
            }
            questions[1].push('=Outline!'+cl[m]+row);
            columnIndex++;
          }
        }
      }
    }
    for (let i=0; i<numOfSections; i++) {
      headersAndPoints[0].push('=Outline!'+cl[i+1]+rows[num][0]);  // ref to section letter
      headersAndPoints[1].push('=Outline!'+cl[i+1]+rows[num][1]);  // ref to section points
      if (i == (numOfSections - 1)) {  // if last section
        let afterGraded = sheet.getRange(1, 7+numOfSections, maxRows, 1);
        afterGraded.setBorder(null, true, null, true, null, null, colour.dark, solidThick);
      }
    }
  }
  sheet.getRange(2, 7, 2, numOfSections).setValues(headersAndPoints);
  sheet.getRange(1, checkboxColumn+1, 2, gradingColumns).setValues(questions);
  sheet.getRange(3, 6).setValue('=Outline!M'+rows[num][0]);
}

// C4. set up conditional formatting & add formulas
function addFormulas(outline, sheetName, num) {
  let sheet = ss.getSheetByName(sheetName);
  let colour = findColour(ss.getSheets()[0]);
  let numOfSections = outline[16][0];
  let start = outline[16][6] + 2;
  let span = outline[15][6];
  let end = start + span;
  let sections = outline[15].slice(1, 5);
  let secPoints = outline[16].slice(1, 5);
  let secBonus = outline[17].slice(1, 5);
  let cl = [0,'A','B','C','D','E','F','G','H','I','J',
           'K','L','M','N','O','P','Q','R','S','T','U',
           'V','W','X','Y','Z','AA','AB','AC','AD','AE',
           'AF','AG','AH','AI','AJ','AK','AL','AM','AN',
           'AO','AP','AQ','AR','AS','AT','AU','AV','AW',
           'AX','AY','AZ'];
  let cb = cl[start-1];  // checkbox column letter
  let cbRef = cb+'4'; // checkbox cell
  let columnIndex = 7;
  let formulae = [[]];
  let rules = sheet.getConditionalFormatRules();
  let range = sheet.getRange(4, start, sheet.getMaxRows()-3, span);
  let rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$'+cbRef+' = FALSE')
      .setBackground(colour.dark)
      .setFontColor(colour.dark)
      .setRanges([range])
      .build();
  let rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground(colour.dark)
      .setFontColor('white')
      .setRanges([range])
      .build();
  rules.push(rule1);
  rules.push(rule2);
  range.setHorizontalAlignment('center');
  formulae[0].push('=if('+cbRef+'=true; countif('+cl[start]+'4:'+cl[end]+'4; \"\") / F$3; \"\")');
  for (let i in sections) {  // for each index in 'sections' array
    if (sections[i] == '') {
      break;
    } else {
      let col = cl[columnIndex];
      let secSpan = secPoints[i] + secBonus[i] - 1;  // calc span of section
      end = cl[start + secSpan] + '4';  // create final cell ref string
      let calcRef = cl[start] + '4:' + end;  // create range reference string
      let ifTrue = 'countif('+calcRef+ '; \"\") / '+col+'$3';
      formulae[0].push('=if('+cbRef+'=true; '+ifTrue+'; \"\")');
      start += secSpan + 1;  // setup next section start
      columnIndex++;
    }
  }
  sheet.getRange(4, 6, 1, numOfSections+1).setFormulas(formulae);
  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, sheet.getMaxColumns());
  if (numOfSections > 1) { // if more than one section, else overall is all we need
    let remainder = numOfSections % 2;
    if (remainder == 0) {  // if an even # of sections
      sheet.getRange(1, (6+(numOfSections/2))).setValue('=Outline!e19');  // write 'SECTIONS' in first row
    } else {  // if 3 sections
      sheet.getRange(1, 8).setValue('=Outline!e19');
    }
  }
  let rows = sheet.getMaxRows();
  let copyFrom = sheet.getRange(4, 6, 1, numOfSections+1);
  let copyTo = sheet.getRange(5, 6, rows-4, numOfSections+1)
  sheet.setColumnWidths(6, numOfSections+1, 60);
  sheet.getRange(4, 6, rows-3, numOfSections+1).setNumberFormat('0.00%');
  copyFrom.copyTo(copyTo);
}