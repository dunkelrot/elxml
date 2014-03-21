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
7. Multiple fonts for cells

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

