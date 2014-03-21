var excel = require("../elxml.js");

// testing

// create a workbook
var wb = excel.createWorkbook();

// create a default style
var defStyle = wb.createStyle("Standard");

var red = wb.color(255,0,0,0);
var black = wb.color(0,0,0,255);

// create a date format
var dateFrmt = wb.addNumberFormat("dd/mm/yy;@");

// create a bold font, the font is derived from the default font
var boldFont = wb.addFont({bold: true});

// create a fill pattern, foreground color is red with a solid fill
var redFill = wb.addPatternFill({fgColor:red, type:excel.constants.PATTERN_TYPE_SOLID});

// create a thick border presentation with a black color
var thickBorderPr = wb.createBorderPr(excel.constants.BORDER_STYLE_THICK, black);

// create a border type, bottom line is set to thinBorder
var border = wb.addBorder({bottom:thickBorderPr});

// create the style
var dateStyle = wb.addStyle(defStyle, {numFormat: dateFrmt, font: boldFont, fill: redFill, border: border});
dateStyle.setAlignment(excel.constants.CELL_ALIGNMENT_H_LEFT,null);

// create a sheet
var sheet = wb.addSheet("mySheet");

// set the width of the first column to 30
sheet.setColumn(1,1,30);

// add a row
var row = sheet.addRow(1);

// add a single cell "A1"
var cell = row.addCell("A","d");

// the the value (ISO date string)
cell.setValue("2014-02-02");

// set the style
cell.setStyle(dateStyle);

// create the file
wb.save("test.01.xlsx");

