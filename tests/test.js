var excel = require("../elxml-archiver.js");

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
var redFill = wb.addPatternFill({fgColor:red, type:excel.PATTERN_TYPE_SOLID});

// create a thick border presentation with a black color
var thickBorderPr = wb.createBorderPr(excel.BORDER_STYLE_THICK, black);

// create a border type, bottom line is set to thinBorder
var border = wb.addBorder({bottom:thickBorderPr});

// create the style
var dateStyle = wb.addStyle(defStyle, {numFormat: dateFrmt, font: boldFont, fill: redFill, border: border});
dateStyle.setAlignment(excel.CELL_ALIGNMENT_H_LEFT,null);

// create a sheet
var sheet = wb.addSheet("mySheet");

// set the width of the first column to 30
sheet.setColumn(1,1,30);

// add a row
var row = sheet.addRow(1);

// add a single cell "A1"
var cellA = row.addCell("A",excel.CELL_TYPE_DATE);
// the the value (ISO date string)
cellA.setValue("2014-02-02");
// set the style
cellA.setStyle(dateStyle);

// add a row
var cellB = row.addCell("B",excel.CELL_TYPE_NUMBER);
cellB.setValue(1);

// add a row
var cellC = row.addCell("C",excel.CELL_TYPE_NUMBER);
cellC.setValue(1);

// add a row
var cellD = row.addCell("D",excel.CELL_TYPE_NUMBER);
cellD.setValue(1);

// add a row with a simple formula
var cellE = row.addCell("E",excel.CELL_TYPE_FORMULA);
cellE.setFormula("SUM(B1:D1)");
cellE.setValue(3); // optional - should be the result of the formula

// create the file
wb.save( "test.01.xlsx" );

