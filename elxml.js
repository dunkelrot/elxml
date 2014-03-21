var builder = require('xmlbuilder');
var fs = require('fs');
var _ = require('underscore');

require('node-zip');

exports.constants = {};

exports.constants.CELL_ALIGNMENT_H_CENTER      = "general";
exports.constants.CELL_ALIGNMENT_H_LEFT        = "left";
exports.constants.CELL_ALIGNMENT_H_CENTER      = "center";
exports.constants.CELL_ALIGNMENT_H_RIGHT       = "right";
exports.constants.CELL_ALIGNMENT_H_FILL        = "fill";
exports.constants.CELL_ALIGNMENT_H_JUSTIFY     = "justify";
exports.constants.CELL_ALIGNMENT_H_CENTER_CONT = "centerContinuous";
exports.constants.CELL_ALIGNMENT_H_DISTRIBUTED = "distributed";
exports.constants.CELL_ALIGNMENT_V_TOP         = "top";
exports.constants.CELL_ALIGNMENT_V_BOTTOM      = "left";
exports.constants.CELL_ALIGNMENT_V_CENTER      = "center";
exports.constants.CELL_ALIGNMENT_V_JUSTIFY     = exports.constants.CELL_ALIGNMENT_H_JUSTIFY;
exports.constants.CELL_ALIGNMENT_V_DISTRIBUTED = exports.constants.CELL_ALIGNMENT_H_DISTRIBUTED;

exports.constants.PATTERN_TYPE_NONE             = "none";
exports.constants.PATTERN_TYPE_SOLID            = "solid";
exports.constants.PATTERN_TYPE_MEDIUM_GRAY      = "mediumGray";
exports.constants.PATTERN_TYPE_DARK_GRAY        = "darkGray";
exports.constants.PATTERN_TYPE_LIGHT_GREY       = "lightGray";
exports.constants.PATTERN_TYPE_DARK_HORIZONTAL  = "darkHorizontal";
exports.constants.PATTERN_TYPE_DARK_VERTICAL    = "darkVertical";
exports.constants.PATTERN_TYPE_DARK_DOWN        = "darkDown";
exports.constants.PATTERN_TYPE_DARK_UP          = "darkUp";
exports.constants.PATTERN_TYPE_DARK_GRID        = "darkGrid";
exports.constants.PATTERN_TYPE_DARK_TRELLIS     = "darkTrellis";
exports.constants.PATTERN_TYPE_LIGHT_HORIZONTAL = "lightHorizontal";
exports.constants.PATTERN_TYPE_LIGHT_VERTICAL   = "lightVertical";
exports.constants.PATTERN_TYPE_LIGHT_DOWN       = "lightDown";
exports.constants.PATTERN_TYPE_LIGHT_UP         = "lightUp";
exports.constants.PATTERN_TYPE_LIGHT_GRID       = "lightGrid";
exports.constants.PATTERN_TYPE_LIGHT_TRELLIS    = "lightTrellis";
exports.constants.PATTERN_TYPE_GRAY125          = "gray125";
exports.constants.PATTERN_TYPE_GRAY0625         = "gray0625";

exports.constants.BORDER_STYLE_NONE             = "none";
exports.constants.BORDER_STYLE_THIN             = "thin";
exports.constants.BORDER_STYLE_MEDIUM           = "medium";
exports.constants.BORDER_STYLE_DASHED           = "dashed";
exports.constants.BORDER_STYLE_DOTTED           = "dotted";
exports.constants.BORDER_STYLE_THICK            = "thick";
exports.constants.BORDER_STYLE_DOUBLE           = "double";
exports.constants.BORDER_STYLE_HAIR             = "hair";
exports.constants.BORDER_STYLE_MEDIUM_DASHED    = "mediumDashed";
exports.constants.BORDER_STYLE_DASH_DOT         = "dashDot";
exports.constants.BORDER_STYLE_MEDIUM_DASH_DOT  = "mediumDashDot";
exports.constants.BORDER_STYLE_DASH_DOT_DOT     = "dashDotDot";
exports.constants.BORDER_STYLE_MEDIUM_DASH_DOT_DOT = "mediumDashDotDot";
exports.constants.BORDER_STYLE_SLANT_DASH_DOT   = "slantDashDot";

exports.createWorkbook = function() {
    return new Workbook();
}

// all internal stuff below this line

// Note that all objects should have a save method which creates the
// complete XML-element structure.
// This save method has one parameter: the parent XML-element

// schemas and content-types
var EXCEL_SCHEMA_MAIN          = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var EXCEL_SCHEMA_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
var EXCEL_SCHEMA_FILE_REL      = "http://schemas.openxmlformats.org/package/2006/relationships"
var EXCEL_SCHEMA_DOC_REL       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
var EXCEL_SCHEMA_REL_TYPE_WB   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
var EXCEL_SCHEMA_REL_STYLES    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
var EXCEL_SCHEMA_REL_SHEET     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
var EXCEL_SCHEMA_STYLES        = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

var EXCEL_TYPE_REL             = "application/vnd.openxmlformats-package.relationships+xml";
var EXCEL_TYPE_WORKBOOK        = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
var EXCEL_TYPE_SHEET           = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
var EXCEL_TYPE_STYLES          = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

// the A-Z array, gets filled as soon as it is needed
var COLUMN_IDS = null;

// creates the xf element
function _writeStyle(ele, style, id) {
    var xf = ele.ele("xf")

    xf.att("numFmtId", style.numFormat != null ? style.numFormat.id : 0);
    xf.att("fontId", style.font != null ? style.font.id : 0); 
    xf.att("fillId", style.fill != null ? style.fill.id : 0);
    xf.att("borderId", style.border != null ? style.border.id : 0);
    
    if (id != null) {
        xf.att("xfId",id);
    }
    if (style.applyNumFormat != 0) {
        xf.att("applyNumberFormat",1);
    }
    if (style.applyFont != 0) {
        xf.att("applyFont",1);
    }
    if (style.applyAlignment != 0) {
        xf.att("applyAlignment",1);
    }
    if (style.applyBorder != 0) {
        xf.att("applyBorder",1);
    }
    if (style.applyFill != 0) {
        xf.att("applyFill",1);
    }    
    if (style.alignment != null) {
        style.alignment.save(xf);
    }
}

// adds the styles from the given list as xf elements to ele
// if all is true every style is added otherwise only those which
// don't have a parent style
function _writeXF(ele, styles, all) {
    
    for (var ii in styles) {
        var style = styles[ii];
        if (style.parentStyle == null || all == true) {
            var id = style.id;
            // use parent ID (which is a zero-based index) for derived styles
            if (style.parentStyle != null) {
                id = style.parentStyle.id;
            }
            // no ID (which is a zero-based index) for cellStyleXfs (all = false)
            if (all == false) {
                id = null;
            }
            _writeStyle(ele, style, id);
        }
    }
}

// adds cellStyle elements to cellStyles for all styles without a parent style
function _writeCellStyles(cellStyles, styles) {
    for (var ii in styles) {
        var style = styles[ii];
        if (style.parentStyle == null) {
            cellStyles.ele("cellStyle").att("name", style.name).att("xfId", style.id).att("builtinId",0);
        }
    }
}

// a rgb(a) color, a (alpha) and name are optional, a defaults to 255 and name to "COLOR"
function Color(r, g, b, a, name) {
    if (a == undefined) {
        this.a = 255;
    }
    if (name == undefined) {
        this.name = "COLOR";
    }
    this.a = a;
    this.r = r;
    this.g = g;
    this.b = b;
    this._auto = false;
    this.index = -1;
}
Color.prototype = {
    constructor : Color,
    auto : function(auto) {
        this._auto = auto;
        return this;
    },
    // returns the internal color value as a ARGB hex string
    toHexARGB : function() {
        var value = this.a;
        value = value << 8;
        value = value | this.r;
        value = value << 8;
        value = value | this.g;
        value = value << 8;
        value = value | this.b;
        return (value+0x100000000).toString(16).substr(-8).toUpperCase();
    },
    // slight deviation from the pattern to make use of Color simpler
    // see PatternFill.save
    save : function(parent, name) {
        var ele = null;
        if (name != undefined) {
            ele = parent.ele(name);
        } else {
            ele = parent.ele(this.name);
        }
        console.log(this.index);
        if (this.index != -1) {
            ele.att("indexed", this.index);
        } else if (this._auto == true) {
            ele.att("auto", this._auto);
        } else {
            ele.att("rgb",this.toHexARGB());
        }
    }
}

// colum configuration
// min is the minimal column index, max is the maximal column index
// min=1 and max=1 -> first column
// min=1 and max=5 -> column 1 to 5
// width is the width (in whatever unit)
// bestFit (false by default) is optional
function Column(min, max, width, bestFit) {
    this._min = min;
    this._max = max;
    this._width = width;
    if (bestFit != undefined) {
        this._bestFit = bestFit;    
    }
}
Column.prototype = {
    constructor : Column,
    min : function(_min) {
        this._min = _min;
        return this;
    },
    max : function(_max) {
        this._max = _max;
        return this;
    },
    width : function(_width) {
        this._width = _width;
        return this;
    },
    bestFit : function(_bestFit) {
        this._bestFit = _bestFit;
        return this;
    },
    save : function(parent) {
        var col = parent.ele("col");
        col.att("min", this._min);
        col.att("max", this._max);
        if (this._width > 0) {
            col.att("width", this._width);
        }
        if (this._bestFit != undefined) {
            col.att("bestFit", this._bestFit);
        }
    }
}

function Font(opts, id) {
    // if something is changed here -> update clone as well!
    this.name = opts.name;
    this.size = opts.size;
    this.bold = opts.bold;
    this.id = id;
}
Font.prototype = {
    constructor : Font,
    save : function(parent) {
        var font = parent.ele("font");
        font.ele("sz").att("val",this.size);
        font.ele("name").att("val",this.name);
        if (this.bold) {
            font.ele("b");
        }
    },
    clone : function() {
        font = new Font({name:this.name, size:this.size, bold:this.bold}, this.id);
        return font;
    }
}
function Fonts() {
    this.fontId = 0;
    this.fonts = [];
    // create a default font
    this.defaultOpts = {bold: false, size: 11, name:"Calibri"};
    this.addFont(this.defaultOpts);
}
Fonts.prototype = {
    constructor : Fonts,
    addFont : function(opts) {
        var font = new Font(opts, this.fontId++);
        this.fonts.push(font);
        return font;
    },
    deriveFromDefault : function(opts) {
        var defFont = this.fonts[0];
        var newFont = defFont.clone();
        newFont.id = this.fontId++;
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts, this.defaultOpts);
        newFont.bold = opts.bold;
        newFont.size = opts.size;
        newFont.name = opts.name;
        this.fonts.push(newFont);
        return newFont;
    },
    save : function(stylesheet) {
        var ele = stylesheet.ele("fonts");
        ele.att("count", this.fonts.length);
        for (var ii in this.fonts) {
            var font = this.fonts[ii];
            font.save(ele);
        }
    }
}

function CellAlignment(h, v) {
    this.h = h;
    this.v = v;
}
CellAlignment.prototype = {
    constructor : CellAlignment,
    save : function(parent) {
        var el = parent.ele("alignment");
        if (this.h != null) {
            el.att("horizontal", this.h);
        }
        if (this.v != null) {
            el.att("vertical", this.v);
        }
    }
}

function BorderPr(style, color) {
    if (color == undefined) {
        this.color = new Color(0,0,0,0);
        this.color.auto = true;
    } else {
        this.color = color;
    }
    this.style = style;
}
BorderPr.prototype = {
    constructor : BorderPr,
    save : function(parent, name) {
        var borderPr = parent.ele(name);
        if (this.color != undefined) {
            this.color.save(borderPr,"color");
        }
        borderPr.att("style", this.style);
    }
}

function Border(opts, id) {
    this.id = id;
    this.top = opts.top;
    this.right = opts.right;
    this.bottom = opts.bottom;
    this.left = opts.left;
}
Border.prototype = {
    constructor : Border,
    save : function(parent) {
        var ele = parent.ele("border");
        this.left   != null ? this.left.save(ele, "left") : ele.ele("left");
        this.right  != null ? this.right.save(ele, "right") : ele.ele("right");
        this.top  != null ? this.top.save(ele, "top") : ele.ele("top");
        this.bottom != null ? this.bottom.save(ele, "bottom") : ele.ele("bottom");
        this.diagonal != null ? this.left.save(ele, "diagonal") : ele.ele("diagonal");
    }
}

function Borders() {
    this.borderId = 0;
    this.borders = [];
    // create a default border
    this.add();
}
Borders.prototype = {
    constructor : Borders,
    add : function(opts) {
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts,{top:null, bottom:null, left:null, right:null});
        var border = new Border(opts, this.borderId++);
        this.borders.push(border);
        return border;
    },
    createBorderPr : function(style, color) {
        return new BorderPr(style, color);
    },
    save : function(stylesheet) {
        var ele = stylesheet.ele("borders");
        ele.att("count", this.borders.length);
        for (var ii in this.borders) {
            var border = this.borders[ii];
            border.save(ele);
        }
    }
}

function PatternFill(opts, id) {
    this.fgColor = opts.fgColor;
    this.bgColor = opts.bgColor;
    this.type = opts.type;
    this.id = id;
}
PatternFill.prototype = {
    constructor : PatternFill,
    save : function(fills) {
        var fill = fills.ele("fill");
        var pf = fill.ele("patternFill");
        if (this.fgColor != null) {
            this.fgColor.save(pf, "fgColor");
        }
        if (this.bgColor != null) {
            this.bgColor.save(pf, "bgColor");
        }
        pf.att("patternType", this.type);
    }
}

function Fills() {
    this.fillId = 0;
    this.fills = [];
    // create two default fills, looks like that this is required
    this.addPatternFill({type: exports.constants.PATTERN_TYPE_NONE});
    this.addPatternFill({type: exports.constants.PATTERN_TYPE_GRAY125});
}
Fills.prototype = {
    constructor : Fills,
    addPatternFill : function(opts) {
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts, {fgColor: null, bgColor: null, type: exports.constants.PATTERN_TYPE_NONE});
        var fill = new PatternFill(opts, this.fillId++);
        this.fills.push(fill);
        return fill;
    },
    save : function(stylesheet) {
        var ele = stylesheet.ele("fills");
        ele.att("count", this.fills.length);
        for (var ii in this.fills) {
            var fill = this.fills[ii];
            fill.save(ele);
        }
    }
}


function NumberFormat(id, formatCode) {
    this.id = id;
    this.formatCode = formatCode;
}
NumberFormat.prototype = {
    constructor : NumberFormat
}

function NumberFormats() {
    this.formats = [];
    this.numFmtId = 1;
}
NumberFormats.prototype = {
    constructor : NumberFormats,
    add : function(formatCode) {
        var format = new NumberFormat(this.numFmtId++, formatCode);
        this.formats.push(format);
        return format;
    }
}
function CellStyle(name, id) {
    this.id = id;
    this.parentStyle = null;
    this.name = name;
    this.numFormat = null;
    this.font = null;
    this.fill = null;
    this.border = null;
    this.applyNumFormat = 0;
    this.applyFont = 0;
    this.applyAlignment = 0;
    this.applyBorder = 0;
    this.applyFill = 0;
    this.alignment = null;
}
CellStyle.prototype = {
    constructor : CellStyle,
    setAlignment : function(h, v) {
        this.alignment = new CellAlignment(h,v);
        this.applyAlignment = 1;
        return this;
    },
    setFont : function(font) {
        this.font = font;
        this.applyFont = 1;
        return this;
    },
    apply : function(style) {
        this.parentStyle = style;
        this.font = style.font;
        this.numFormat = style.numFormat;
        this.fill = style.fill;
        this.border = style.border;
        this.alignment = style.alignment;
    }
}

function CellStyles() {
    this.nextStyleId = 0;
    this.styles = [];
}
CellStyles.prototype = {
    constructor : CellStyles,
    create : function(name) {
        var style = new CellStyle(name, this.nextStyleId++);
        this.styles.push(style);
        return style;
    },
    derive : function(cellStyle, opts) {
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts, {numFrmt: null, fill: null, font: null, border: null});
        var style = new CellStyle(null, this.nextStyleId++);
        style.apply(cellStyle);
        
        if (opts.numFormat != null) {
            style.numFormat = opts.numFormat;
            style.applyNumFormat = 1;
        }
        if (opts.font != null) {
            style.font = opts.font;
            style.applyFont = 1;
        }
        if (opts.fill != null) {
            style.fill = opts.fill;
            style.applyFill = 1;
        }
        if (opts.border != null) {
            style.border = opts.border;
            style.applyBorder = 1;
        }        
        this.styles.push(style);
        return style;
    },
    countDirectStyles : function() {
        var count = 0;
        for (var ii in this.styles) {
            if (this.styles[ii].parentStyle == null) {
                count++;
            }
        }
        return count;
    },
    count : function() {
        return this.styles.length;
    },
    getStyles : function() {
        return this.styles;
    }
}

function Cell(row, index, type) {
    this.index = index;
    this.row = row;
    this.type = type;
    this.style = null;
    this.value = null;
}
Cell.prototype = {
    constructor : Cell,
    setValue : function(value) {
        this.value = value;
        return this;
    },
    setStyle : function(style) {
        this.style = style;
        return this;
    },
    save : function(row) {
        var ele = row.ele("c").att("r", this.index + this.row.index);
        ele.att("t", this.type);
        if (this.value != null) {
            if (this.type == undefined || this.type == "inlineStr") {
                ele.ele("is").ele("t").t(this.value);
            } else {
                ele.ele("v").t(this.value);
            }
        }
        if (this.style != null) {
            ele.att("s", this.style.id);
        }
    }
}

function Row(index) {
    this.index = index;
    this.cells = [];
}
Row.prototype = {
    constructor : Row,
    addCell : function(index, type) {
        var cell = new Cell(this, index, type);
        this.cells.push(cell);
        return cell;
    },
    save : function(sheetData) {
        var ele = sheetData.ele("row");
        ele.att("r", this.index);
        for (var ii in this.cells) {
            this.cells[ii].save(ele);
        }
    }
}

function Sheet(id, name) {
    this.id = id;
    this.name = name;
    this.rows = [];
    this.cols = [];
}
Sheet.prototype = {
    constructor : Sheet,
    addRow : function(index) {
        var row = new Row(index);
        this.rows.push(row);
        return row;
    },
    setColumn : function(min, max, width, bestFit) {
        var col = new Column(min, max, width, bestFit);
        this.cols.push(col);
        return col;
    },
    save : function(root) {
        if (this.cols.length > 0) {
            var colsEle = root.ele("cols");
            for (var ii in this.cols) {
                this.cols[ii].save(colsEle);
            }
        }
        if (this.rows.length > 0) {
            var sheetData = root.ele("sheetData")
            for (var ii in this.rows) {
                this.rows[ii].save(sheetData);
            }
        }
    }
}

function Workbook () {
    this.sheets = [];
    this.relID = 1;
    this.styles = new CellStyles();
    this.numberFormats = new NumberFormats();
    this.fills = new Fills();
    this.fonts = new Fonts();
    this.borders = new Borders();
}
Workbook.prototype = {
    constructor : Workbook,
    createStyle : function(name) {
        return this.styles.create(name);
    },
    addSheet : function(name) {
        var sheet = new Sheet(this.relID++, name);
        this.sheets.push(sheet);
        return sheet;
    },
    createContents : function(zipFolder) {
        var contents = builder.create("Types",{version: '1.0', encoding: 'utf-8'});
        contents.att("xmlns",EXCEL_SCHEMA_CONTENT_TYPES);
        
        contents.ele("Default").att("Extension","xml").att("ContentType",EXCEL_TYPE_WORKBOOK);
        contents.ele("Default").att("Extension","rels").att("ContentType",EXCEL_TYPE_REL);
        contents.ele("Override").att("PartName","/xl/styles.xml").att("ContentType",EXCEL_TYPE_STYLES);

        for (var ii in this.sheets) {
            var sheetName = "/xl/worksheets/sheet" + this.sheets[ii].id + ".xml";
            contents.ele("Override").att("PartName",sheetName).att("ContentType",EXCEL_TYPE_SHEET);
        }

        var xmlString = contents.end({ pretty: true, indent: '  ', newline: '\n' });
        zipFolder.file("[Content_Types].xml", xmlString);
    },
    createMainRelations : function(zipFolder) {
        var mainRelations = builder.create("Relationships",{version: '1.0', encoding: 'utf-8'});
        mainRelations.att("xmlns",EXCEL_SCHEMA_FILE_REL);
        mainRelations.ele("Relationship").att("Id","rId" + (this.relID++)).att("Type",EXCEL_SCHEMA_REL_TYPE_WB).att("Target","/xl/workbook.xml");
        zipFolder.file(".rels", mainRelations.end({ pretty: true, indent: '  ', newline: '\n' }));
    },
    createWorkbook : function(zipFolder) {
        var workbookRelations = builder.create("workbook",{version: '1.0', encoding: 'utf-8'});
        workbookRelations.att("xmlns",EXCEL_SCHEMA_MAIN).att("xmlns:r",EXCEL_SCHEMA_DOC_REL);

        var sheetsEle = workbookRelations.ele("sheets");
        for (var ii in this.sheets) {
            var sheet = this.sheets[ii];
            sheetsEle.ele("sheet").att("name",sheet.name).att("sheetId",sheet.id).att("r:id","rId" + sheet.id);
        }

        var xmlString = workbookRelations.end({ pretty: true, indent: '  ', newline: '\n' });
        zipFolder.file("workbook.xml", xmlString);
    },
    createWorkbookRelations : function(zipFolder) {
        
        var relations = builder.create("Relationships",{version: '1.0', encoding: 'utf-8'});
        relations.att("xmlns",EXCEL_SCHEMA_FILE_REL);

        var sheetRel = relations.ele("Relationship");
        sheetRel.att("Type",EXCEL_SCHEMA_REL_STYLES);
        sheetRel.att("Target","/xl/styles.xml");
        sheetRel.att("Id","rId" + (this.relID++));
        
        for (var ii in this.sheets) {
            var sheet = this.sheets[ii];
            var sheetRel = relations.ele("Relationship");
            sheetRel.att("Type",EXCEL_SCHEMA_REL_SHEET);
            sheetRel.att("Target","/xl/worksheets/sheet" + sheet.id + ".xml");
            sheetRel.att("Id","rId" + sheet.id);
        }

        var xmlString = relations.end({ pretty: true, indent: '  ', newline: '\n' });
        zipFolder.file("workbook.xml.rels", xmlString);
    },
    createStyles : function(zipFolder) {
        
        var stylesheet = builder.create("styleSheet",{version: '1.0', encoding: 'utf-8'});
        stylesheet.att("xmlns",EXCEL_SCHEMA_STYLES);
        
        // number formats
        var numFmts = stylesheet.ele("numFmts");
        numFmts.att("count", this.numberFormats.formats.length);
        for (var ii in this.numberFormats.formats) {
            var fmt = this.numberFormats.formats[ii];
            numFmts.ele("numFmt").att("numFmtId", fmt.id).att("formatCode", fmt.formatCode);
        }

        this.fonts.save(stylesheet);
        this.fills.save(stylesheet);
        this.borders.save(stylesheet);

        // cellStyleXfs
        var cellStyleXfs = stylesheet.ele("cellStyleXfs");
        cellStyleXfs.att("count",this.styles.countDirectStyles());
        _writeXF(cellStyleXfs, this.styles.getStyles(), false);
        
        var cellXfs = stylesheet.ele("cellXfs");
        cellXfs.att("count",this.styles.count());
        _writeXF(cellXfs, this.styles.getStyles(), true);
        
        var cellStyles = stylesheet.ele("cellStyles");
        cellStyles.att("count",this.styles.countDirectStyles());
        _writeCellStyles(cellStyles, this.styles.getStyles());
        
        var xmlString = stylesheet.end({ pretty: true, indent: '  ', newline: '\n' });
        zipFolder.file("styles.xml", xmlString);
    },
    save : function(fileName) {
        
        var zip = new JSZip();
        var relsFolder   = zip.folder("_rels");
        var xlFolder     = zip.folder("xl");
        var xlRelsFolder = xlFolder.folder("_rels");
        var sheetsFolder = xlFolder.folder("worksheets");

        // create content-types
        this.createContents(zip);
        
        // create main relationships
        this.createMainRelations(relsFolder);
        
        // create sheets relationships
        this.createWorkbookRelations(xlRelsFolder);
        
        // create styles and the workbook
        this.createStyles(xlFolder);
        this.createWorkbook(xlFolder);
        
        // create sheet files
        for (var ii in this.sheets) {
            var sheet = this.sheets[ii];
            
            var root = builder.create("worksheet", {version: '1.0', encoding: 'utf-8'});
            root.att("xmlns", EXCEL_SCHEMA_MAIN);
            sheet.save(root);
            
            var xmlString = root.end({ pretty: true, indent: '  ', newline: '\n' });
            sheetsFolder.file("sheet" + sheet.id + ".xml", xmlString);
        }
        
        // create zip
        var data = zip.generate({base64:false,compression:'DEFLATE'});
        fs.writeFileSync(fileName, data, 'binary');
        return fileName;
    },
    coords : function(column) {
        var col = column - 1;
        if (COLUMN_IDS == null) {
            COLUMN_IDS = [];
            for (var ii = 0; ii < 26; ii++) {
                COLUMN_IDS.push(String.fromCharCode(65 + ii));
            }
        }
        var factor = ~~(col / 26);
        if (factor == 0) {
            return COLUMN_IDS[col];
        } else {
            var i2 = col - factor * 26;
            return COLUMN_IDS[factor] + COLUMN_IDS[i2];
        }
    },
    color : function(r,g,b) {
        return new Color(r,g,b,0);
    },
    createBorderPr : function(style, color) {
        return this.borders.createBorderPr(style, color);
    },
    addNumberFormat : function(value) {
        return this.numberFormats.add(value);
    },
    addFont : function(opts) {
        return this.fonts.deriveFromDefault(opts);
    },
    addPatternFill : function(opts) {
        return this.fills.addPatternFill(opts);
    },
    addBorder : function(opts) {
        return this.borders.add(opts);
    },
    addStyle : function(style, opts) {
        return this.styles.derive(style, opts)
    }
}

