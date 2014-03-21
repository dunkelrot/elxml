elxml
=====

A minimalistic Excel OOXML writer.
The main purpose is to create simple Excel files via JavaScript. The current implementation supports

1. Multiple sheets
2. Creation of rows and cells
3. Column width definition
4. PatternFills for cells
5. Borders for cells
6. Number formats for cells
7. Fonts for cells

Most of this functionality is very basic.

### Usage

Create a workbook:

```javascript
var excel = require("elxml.js");
var wb = excel.createWorkbook();
```

Create a sheet within the workbook:

```javascript
var sheet = wb.addSheet("mySheet");
```

Add a row with a single cell and set a value for the cell

```javascript
var row = sheet.addRow(1);
var cell = row.addCell("A","d");  // the value is a date in ISO standard notation
cell.setValue("2014-02-02");
```

Last step is to save the workbook:

```javascript
wb.save("test.01.xlsx");
```

### Advanced usage


#### Number formats

The above example creates an Excel file with a date cell but without a proper number format.

To apply the number format to the cell we have to create a cell style. Cell styles are based on 
a default cell style which defines the number format, fill, borders and the font for those cells
which don't use any cell style. 
In elxml the default is: Calibri with a size of 11, no borders, no fill, no number format.

Lets set a date number format to see a better formatted date in Excel:

```javascript
// create a date format
var dateFrmt = wb.addNumberFormat("dd/mm/yy;@");

// create a default style
var defStyle = wb.createStyle("Standard");
// derive a new style from the default style
var dateStyle = wb.addStyle(defStyle, {numFrmt: dateFrmt});

// apply the style
cell.setStyle(dateStyle);
```

To see which number formats are available take a look at the OOXML spec.

#### Fills

You can define pattern fills. Possible options for a pattern fill are: `fgColor`, `bgColor` and `type`.
Lets create a red solid fill:

```javascript
// create a color (RGBA)
var red = wb.color(255,0,0,0);
// create a pattern fill
var redFill = wb.addPatternFill({fgColor:red, type:excel.constants.PATTERN_TYPE_SOLID});
// create a red fill style
var redFillStyle = wb.addStyle(defStyle, {fill: redFill});
// apply the style
cell.setStyle(dateStyle);
```

#### Borders
You can define the borders for a cell. First you have to create a border representation object
which is used to define the border style.

```javascript
// create a color (RGBA)
var black = wb.color(0,0,0,0);
// create a thick border presentation with a black color
var thickBorderPr = wb.createBorderPr(excel.constants.BORDER_STYLE_THICK, black);
// create a border type, bottom line is set to thickBorderPr
var border = wb.addBorder({bottom:thickBorderPr});
// create a border style
var borderStyle = wb.addStyle(defStyle, {border: border});
// apply the style
cell.setStyle(borderStyle);
```

#### Fonts
```javascript
// create a bold font, the font is derived from the default font
var boldFont = wb.addFont({bold: true});
// create a border style
var boldFontStyle = wb.addStyle(defStyle, {font: boldFont});
// apply the style
cell.setStyle(boldFontStyle);
```

#### Column width
tbd

