// A1. ungroup rows
function ungroup(outline) {
  let rows = [23, 44, 65, 86];
  rows.forEach(function(row) {
    let triggerRange = outline.getRange(row, 1);
    let trigger = triggerRange.getValue();
    let title = outline.getRange(row+1, 3).getValue();
    if (trigger == 'grouped' && title == '#') {  // if row is grouped
      let group = outline.getRowGroup(row+1, 1);
      group.remove();
      triggerRange.clear();
    }
  });
}

// A2. show and clear all rows below first table
function clean(outline) {
  let blankRange = outline.getRange('C23:M106');  // range below the first table
  let rows = outline.getMaxRows();
  outline.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE)
         .setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER)
         .showRows(1, rows);
  blankRange.removeCheckboxes()
            .clear();
}

// A3. create outline table for each outline or change the formatting of all tables
function spawnOutlines(outline, format) {
  let row = 23;
  let colour = findColour(outline);
  let info = outline.getRange('B2:D2').getValues();
  let numOfAssessments = info[0][0];
  let header = info[0][2];  // 'Formative 1' header;
  let source = outline.getNamedRanges()[0].getRange();  // C2:M22
  for (let i=0; i<numOfAssessments; i++) {
    if (format) {
      let bandingRange = outline.getRange((row+2), 3);
      let banding = bandingRange.getBandings()[0];
      banding.setSecondRowColor(colour.lightest);
      source.copyFormatToRange(outline, 3, 13, row, (row+20));
    } else {
      source.copyTo(outline.getRange(row, 3, 20, 11));  // copy to C23:M43, C44:M64...
      let test = outline.getRange(row+1, 3, 20, 11);
      test.shiftRowGroupDepth(1);
      outline.getRange(row, 1).setValue('grouped').setFontColor('white');
      if (i == (numOfAssessments-1)) {
        outline.getRange(row, 4).setValue('Summative');
      } else {
        header = header.replace((i+1),(i+2));  // increase the last number of header
        outline.getRange(row, 4).setValue(header);
      }
    }
    row += 21;  // row == 44, 65...
  }
}

// A4. hide blank rows
function hideInOutline(outline) {
  let bottomRow = outline.getLastRow();
  let rows = outline.getMaxRows();
  outline.hideRows(bottomRow+1, rows-bottomRow);
}