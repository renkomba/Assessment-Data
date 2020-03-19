// D1. turn colour input into matching object
function findColour(outline) {
  let pink = {name: 'pink', lightest: '#fadfeb', lighter: '#d5a6bd',
              light: '#c27ba0', plain: '#a64d79', deep: '#741b47', dark: '#4c1130'};
  let purple = {name: 'purple', lightest: '#e9e1fa', lighter: '#b4a7d6',
                light: '#8e7cc3', plain: '#674ea7', deep: '#351c75', dark: '#20124d'};
  let blue = {name: 'blue', lightest: '#d5e8fa', lighter: '#9fc5e8',
              light: '#6fa8dc', plain: '#3d85c6', deep: '#0b5394', dark: '#073763'};
  let teal = {name: 'teal', lightest: '#e0f7fa', lighter: '#a2c4c9',
              light: '#76a5af', plain: '#45818e', deep: '#134f5c', dark: '#0c343d'};
  let green = {name: 'green', lightest: '#d4fac3', lighter: '#b6d7a8',
               light: '#93c47d', plain: '#6aa84f', deep: '#38761d', dark: '#274e13'};
  let yellow = {name: 'yellow', lightest: '#fff1ca', lighter: '#ffe599',
                light: '#ffd966', plain: '#f1c232', deep: '#bf9000', dark: '#7f6000'};
  let orange = {name: 'orange', lightest: '#ffe2c4', lighter: '#f9cb9c',
                light: '#f6b26b', plain: '#e69138', deep: '#b45f06', dark: '#783f04'};
  let red = {name: 'red', lightest: '#fddcdc', lighter: '#dd7e6b',
             light: '#cc4125', plain: '#a61c00', deep: '#85200c', dark: '#5b0f00'};
  let grey = {name: 'grey', lightest: '#f8f8f8', lighter: '#cccccc',
              light: '#b7b7b7', plain: '#999999', deep: '#666666', dark: '#434343'};
  let colours = [purple, pink, blue, teal, green, yellow, orange, red, grey];
  let colourNames = [purple.name, pink.name, blue.name, teal.name, green.name,
                     yellow.name, orange.name, red.name, grey.name];
  let colourString = outline.getRange('B3').getValue();
  let i = Math.abs(colourNames.indexOf(colourString));  // if not found, '-1' = '1'
  return colours[i];
}

function setBackgrounds(outline, colour) {
  let darkRange = outline.getRangeList(['A1:C1', 'A6']);
  let deepRange = outline.getRangeList(['C3:M3', 'A5', 'A13']);
  let plainRange = outline.getRangeList(['C2:M2', 'A4', 'C19:I19', 'C20:D21',
                                         'E21', 'J19:J21', 'L19:L21']);
  let lightRange = outline.getRange('A3');
  let lighterRange = outline.getRange('A2');
  let bandingRange = outline.getRange('C4:M18');
  let banding = bandingRange.getBandings()[0];
  darkRange.setBackground(colour.dark);
  deepRange.setBackground(colour.deep);
  plainRange.setBackground(colour.plain);
  lightRange.setBackground(colour.light);
  lighterRange.setBackground(colour.lighter);
  banding.setSecondRowColor(colour.lightest);
}

function setBorders(outline, colour) {
  let middleHorizontal = outline.getRangeList(['C1:M2', 'C18:M19', 'C21:M22']);
  let solidRight = outline.getRange('M1:M21');
  let middleVertical = outline.getRangeList(['B2:C21', 'J4:M18']);
  let doubleRight = outline.getRange('C4:D18');
  let solid = SpreadsheetApp.BorderStyle.SOLID_THICK;
  let double = SpreadsheetApp.BorderStyle.DOUBLE;
  middleHorizontal.setBorder(null, null, null, null, null, true, colour.dark, solid);
  middleVertical.setBorder(null, null, null, null, true, null, colour.dark, solid);
  solidRight.setBorder(null, null, null, true, null, null, colour.dark, solid);
  doubleRight.setBorder(null, null, null, null, true, null, colour.dark, double);
}