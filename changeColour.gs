var ss = SpreadsheetApp.getActiveSpreadsheet();

// D1. 
function findColour(outline) {
  var pink = {name: "pink", lightest: "#fadfeb", lighter: "#d5a6bd",
              light: "#c27ba0", plain: "#a64d79", deep: "#741b47", dark: "#4c1130"};
  var purple = {name: "purple", lightest: "#e9e1fa", lighter: "#b4a7d6",
                light: "#8e7cc3", plain: "#674ea7", deep: "#351c75", dark: "#20124d"};
  var blue = {name: "blue", lightest: "#d5e8fa", lighter: "#9fc5e8",
              light: "#6fa8dc", plain: "#3d85c6", deep: "#0b5394", dark: "#073763"};
  var teal = {name: "teal", lightest: "#e0f7fa", lighter: "#a2c4c9",
              light: "#76a5af", plain: "#45818e", deep: "#134f5c", dark: "#0c343d"};
  var green = {name: "green", lightest: "#d4fac3", lighter: "#b6d7a8",
               light: "#93c47d", plain: "#6aa84f", deep: "#38761d", dark: "#274e13"};
  var yellow = {name: "yellow", lightest: "#fff1ca", lighter: "#ffe599",
                light: "#ffd966", plain: "#f1c232", deep: "#bf9000", dark: "#7f6000"};
  var orange = {name: "orange", lightest: "#ffe2c4", lighter: "#f9cb9c",
                light: "#f6b26b", plain: "#e69138", deep: "#b45f06", dark: "#783f04"};
  var red = {name: "red", lightest: "#fddcdc", lighter: "#dd7e6b",
             light: "#cc4125", plain: "#a61c00", deep: "#85200c", dark: "#5b0f00"};
  var grey = {name: "grey", lightest: "#f8f8f8", lighter: "#cccccc",
              light: "#b7b7b7", plain: "#999999", deep: "#666666", dark: "#434343"};
  var colours = [pink, purple, blue, teal, green, yellow, orange, red, grey];
  var colourString = outline.getRange("B3").getValue();
  if (colourString == "- - - - -") {
    return pink;
  } else {
    for (var colour of colours) {
      if (colour.name == colourString) {
        return colour;
      }
    }
  }
}

// D2. 
function setBackgrounds(outline, colour) {
  var darkRange = outline.getRangeList(["A1:C1", "A6"]);
  var deepRange = outline.getRangeList(["C3:M3", "A5", "A13"]);
  var plainRange = outline.getRangeList(["C2:M2", "A4", "C19:I19", "C20:D21",
                                         "E21", "J19:J21", "L19:L21"]);
  var lightRange = outline.getRange("A3");
  var lighterRange = outline.getRange("A2");
  var bandingRange = outline.getRange("C4:M18");
  var banding = bandingRange.getBandings()[0];
  darkRange.setBackground(colour.dark);
  deepRange.setBackground(colour.deep);
  plainRange.setBackground(colour.plain);
  lightRange.setBackground(colour.light);
  lighterRange.setBackground(colour.lighter);
  banding.setSecondRowColor(colour.lightest);
}

function setBorders(outline, colour) {
  var middleHorizontal = outline.getRangeList(["C1:M2", "C18:M19", "C21:M22"]);
  var solidRight = outline.getRange("M1:M21");
  var middleVertical = outline.getRangeList(["B2:C21", "J4:M18"]);
  var doubleRight = outline.getRange("C4:D18");
  var solid = SpreadsheetApp.BorderStyle.SOLID_THICK;
  var double = SpreadsheetApp.BorderStyle.DOUBLE;
  middleHorizontal.setBorder(null, null, null, null, null, true, colour.dark, solid);
  middleVertical.setBorder(null, null, null, null, true, null, colour.dark, solid);
  solidRight.setBorder(null, null, null, true, null, null, colour.dark, solid);
  doubleRight.setBorder(null, null, null, null, true, null, colour.dark, double);
}