var builder = require('xmlbuilder');
var fs = require('fs');
var _ = require('underscore');

require('node-zip');

/** @constant CELL_ALIGNMENT_H_CENTER */
exports.CELL_ALIGNMENT_H_CENTER      = "general";
/** @constant CELL_ALIGNMENT_H_LEFT */
exports.CELL_ALIGNMENT_H_LEFT        = "left";
/** @constant CELL_ALIGNMENT_H_CENTER */
exports.CELL_ALIGNMENT_H_CENTER      = "center";
/** @constant CELL_ALIGNMENT_H_RIGHT */
exports.CELL_ALIGNMENT_H_RIGHT       = "right";
/** @constant CELL_ALIGNMENT_H_FILL */
exports.CELL_ALIGNMENT_H_FILL        = "fill";
/** @constant CELL_ALIGNMENT_H_JUSTIFY */
exports.CELL_ALIGNMENT_H_JUSTIFY     = "justify";
/** @constant CELL_ALIGNMENT_H_CENTER_CONT */
exports.CELL_ALIGNMENT_H_CENTER_CONT = "centerContinuous";
/** @constant CELL_ALIGNMENT_H_DISTRIBUTED */
exports.CELL_ALIGNMENT_H_DISTRIBUTED = "distributed";
/** @constant CELL_ALIGNMENT_V_TOP */
exports.CELL_ALIGNMENT_V_TOP         = "top";
/** @constant CELL_ALIGNMENT_V_BOTTOM */
exports.CELL_ALIGNMENT_V_BOTTOM      = "left";
/** @constant CELL_ALIGNMENT_V_CENTER */
exports.CELL_ALIGNMENT_V_CENTER      = "center";
/** @constant CELL_ALIGNMENT_V_JUSTIFY */
exports.CELL_ALIGNMENT_V_JUSTIFY     = exports.CELL_ALIGNMENT_H_JUSTIFY;
/** @constant CELL_ALIGNMENT_V_DISTRIBUTED */
exports.CELL_ALIGNMENT_V_DISTRIBUTED = exports.CELL_ALIGNMENT_H_DISTRIBUTED;

/** @constant PATTERN_TYPE_NONE */
exports.PATTERN_TYPE_NONE             = "none";
/** @constant PATTERN_TYPE_SOLID */
exports.PATTERN_TYPE_SOLID            = "solid";
/** @constant PATTERN_TYPE_MEDIUM_GRAY */
exports.PATTERN_TYPE_MEDIUM_GRAY      = "mediumGray";
/** @constant PATTERN_TYPE_DARK_GRAY */
exports.PATTERN_TYPE_DARK_GRAY        = "darkGray";
/** @constant PATTERN_TYPE_LIGHT_GREY */
exports.PATTERN_TYPE_LIGHT_GREY       = "lightGray";
/** @constant PATTERN_TYPE_DARK_HORIZONTAL */
exports.PATTERN_TYPE_DARK_HORIZONTAL  = "darkHorizontal";
/** @constant PATTERN_TYPE_DARK_VERTICAL */
exports.PATTERN_TYPE_DARK_VERTICAL    = "darkVertical";
/** @constant PATTERN_TYPE_DARK_DOWN */
exports.PATTERN_TYPE_DARK_DOWN        = "darkDown";
/** @constant PATTERN_TYPE_DARK_DOWN */
exports.PATTERN_TYPE_DARK_DOWN          = "darkUp";
/** @constant PATTERN_TYPE_DARK_GRID */
exports.PATTERN_TYPE_DARK_GRID        = "darkGrid";
/** @constant PATTERN_TYPE_DARK_TRELLIS */
exports.PATTERN_TYPE_DARK_TRELLIS     = "darkTrellis";
/** @constant PATTERN_TYPE_LIGHT_HORIZONTAL */
exports.PATTERN_TYPE_LIGHT_HORIZONTAL = "lightHorizontal";
/** @constant PATTERN_TYPE_LIGHT_VERTICAL */
exports.PATTERN_TYPE_LIGHT_VERTICAL   = "lightVertical";
/** @constant PATTERN_TYPE_LIGHT_DOWN */
exports.PATTERN_TYPE_LIGHT_DOWN       = "lightDown";
/** @constant PATTERN_TYPE_LIGHT_UP */
exports.PATTERN_TYPE_LIGHT_UP         = "lightUp";
/** @constant PATTERN_TYPE_LIGHT_GRID */
exports.PATTERN_TYPE_LIGHT_GRID       = "lightGrid";
/** @constant PATTERN_TYPE_LIGHT_TRELLIS */
exports.PATTERN_TYPE_LIGHT_TRELLIS    = "lightTrellis";
/** @constant PATTERN_TYPE_GRAY125 */
exports.PATTERN_TYPE_GRAY125          = "gray125";
/** @constant PATTERN_TYPE_GRAY0625 */
exports.PATTERN_TYPE_GRAY0625         = "gray0625";

/** @constant BORDER_STYLE_NONE */
exports.BORDER_STYLE_NONE             = "none";
/** @constant BORDER_STYLE_THIN */
exports.BORDER_STYLE_THIN             = "thin";
/** @constant BORDER_STYLE_MEDIUM */
exports.BORDER_STYLE_MEDIUM           = "medium";
/** @constant BORDER_STYLE_DASHED */
exports.BORDER_STYLE_DASHED           = "dashed";
/** @constant BORDER_STYLE_DOTTED */
exports.BORDER_STYLE_DOTTED           = "dotted";
/** @constant BORDER_STYLE_THICK */
exports.BORDER_STYLE_THICK            = "thick";
/** @constant BORDER_STYLE_DOUBLE */
exports.BORDER_STYLE_DOUBLE           = "double";
/** @constant BORDER_STYLE_HAIR */
exports.BORDER_STYLE_HAIR             = "hair";
/** @constant BORDER_STYLE_MEDIUM_DASHED */
exports.BORDER_STYLE_MEDIUM_DASHED    = "mediumDashed";
/** @constant BORDER_STYLE_DASH_DOT */
exports.BORDER_STYLE_DASH_DOT         = "dashDot";
/** @constant BORDER_STYLE_MEDIUM_DASH_DOT */
exports.BORDER_STYLE_MEDIUM_DASH_DOT  = "mediumDashDot";
/** @constant BORDER_STYLE_DASH_DOT_DOT */
exports.BORDER_STYLE_DASH_DOT_DOT     = "dashDotDot";
/** @constant BORDER_STYLE_MEDIUM_DASH_DOT_DOT */
exports.BORDER_STYLE_MEDIUM_DASH_DOT_DOT = "mediumDashDotDot";
/** @constant BORDER_STYLE_SLANT_DASH_DOT */
exports.BORDER_STYLE_SLANT_DASH_DOT   = "slantDashDot";

/**
 * @constant CELL_TYPE_DATE
 * @desc date type for cells
 */
exports.CELL_TYPE_DATE   = "d";
/**
 * @constant CELL_TYPE_NUMBER
 * @desc number type for cells
 */
exports.CELL_TYPE_NUMBER   = "n";
/**
 * @constant CELL_TYPE_BOOLEAN
 * @desc boolean type for cells
 */
exports.CELL_TYPE_BOOLEAN   = "n";
/**
 * @constant CELL_TYPE_ERROR
 * @desc error type for cells
 */
exports.CELL_TYPE_ERROR   = "e";
/**
 * @constant CELL_TYPE_STRING
 * @desc inline string (not a shared string as shared string are currently not supported) type for cells
 */
exports.CELL_TYPE_STRING   = "inlineStr";
/**
 * @constant CELL_TYPE_FORMULA
 * @desc formula string type for cells
 */
exports.CELL_TYPE_FORMULA  = "str";

/**
 * @constant CELL_FORMULA_NORMAL
 * @desc type of formula 
 */
exports.CELL_FORMULA_NORMAL = "normal";

/**
 * Creates a {@linkcode Workbook} instance
 */
exports.createWorkbook = function() {
    return new Workbook();
}

// all internal stuff below this line

// Note that all objects should have a 'save' method which creates the
// complete XML-element structure.
// This save method has one mandatory parameter: the parent XML-element


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

/**
 * @constructor
 * @param r - red (0-255) {number}
 * @param g - green (0-255) {number}
 * @param b - blue (0-255) {number}
 * @param a - alpha (0-255) {number}
 * @param name - the name of the color {string}
 */
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
/**
 * @class
 */
Color.prototype = {
    constructor : Color,
    auto : function(auto) {
        this._auto = auto;
        return this;
    },
    /**
     * @returns the internal color value as a ARGB hex string
     */
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

/**
 * @class
 * @param style {string} one of the predefined border styles
 * @param color {Color} the border color
 * @desc use {@linkcode Workbook#createBorderPr} to create new instances
 */
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
        if (this.color != undefined && !this.color.auto) {
            this.color.save(borderPr,"color");
        }
        borderPr.att("style", this.style);
    }
}

/**
 * @class
 */
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

// Internal list of Borders
// The constructor adds a default border with no styles.
// Borders keeps track of the right internal ID for new Border objects.
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

/**
 * @class
 */
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
    this.addPatternFill({type: exports.PATTERN_TYPE_NONE});
    this.addPatternFill({type: exports.PATTERN_TYPE_GRAY125});
}
Fills.prototype = {
    constructor : Fills,
    addPatternFill : function(opts) {
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts, {fgColor: null, bgColor: null, type: exports.PATTERN_TYPE_NONE});
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

/**
 * @class
 */
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

/**
 * @class
 */
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
    setBorder : function (border) {
        this.border = border;
        this.applyBorder = 1;
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
        _.defaults(opts, {numFormat: null, fill: null, font: null, border: null});
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

/**
 * @class
 */
function Cell(row, index, type) {
    this.index = index;
    this.row = row;
    this.type = type;
    this.style = null;
    this.value = null;
    this.formula = null;
}
Cell.prototype = {
    constructor : Cell,
    /**
     * @param value {string} the cell value
     * @return this
     */
    setValue : function(value) {
        this.value = value;
        return this;
    },
    /**
     * @param formula {string} the cell formula
     * In case you set a formula, make sure to set the type to CELL_TYPE_FORMULA
     * @return this
     */
    setFormula : function(formula) {
        this.formula = formula;
        return this;
    },
    /**
     * @param style {CellStyle} the cell style
     * @return this
     */
    setStyle : function(style) {
        this.style = style;
        return this;
    },
    save : function(row) {
        var ele = row.ele("c").att("r", this.index + this.row.index);
        ele.att("t", this.type);
        if (this.formula != null) {
            var f = ele.ele("f");
            f.att("t", exports.CELL_FORMULA_NORMAL);
            f.t(this.formula);
        }
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

/**
 * @class
 */
function Row(index, opts) {
    opts = (opts == undefined ? {} : opts);
    _.defaults(opts, {height:-1});
    this.height = opts.height;
    this.index = index;
    this.cells = [];
}
Row.prototype = {
    constructor : Row,
    /**
     * @param index {number} the column index (A-...)
     * @param type {string} one of the predefined cell types
     * @desc Adds a {@linkcode Cell} to this row.
     */
    addCell : function(index, type) {
        var cell = new Cell(this, index, type);
        this.cells.push(cell);
        return cell;
    },
    save : function(sheetData) {
        var ele = sheetData.ele("row");
        ele.att("r", this.index);
        if (this.height != -1) {
            ele.att("ht", this.height);
            ele.att("customHeight", 1);
        }
        for (var ii in this.cells) {
            this.cells[ii].save(ele);
        }
    }
}

/**
 * @class
 * @param range {string} range identifier (eg. 'A1:B2')
 * @desc
 */
function MergeCell(range) {
    this.range = range;
}
MergeCell.prototype = {
    constructor : MergeCell,

    save : function(mergeCells) {
        var ele = mergeCells.ele("mergeCell");
        ele.att("ref", this.range);
    }
}

/**
 * @class
 * @param id {number} the sheet id
 * @param name {string} the sheet name
 * @desc Don't use this constructor, use {@linkcode Workbook#addSheet} instead.
 */
function Sheet(id, name) {
    this.id = id;
    this.name = name;
    this.rows = [];
    this.cols = [];
    this.merges = [];
}
Sheet.prototype = {
    constructor : Sheet,

    /**
     * @typedef RowOpts
     * @type {object}
     * @property {number} height - the row height
     * @desc all parameters are optional
     */

    /**
     * @param index {number} the row index (1-...)
     * @param opts {RowOpts} additional options
     * @desc Adds a {@linkcode Row} to this sheet.
     */
    addRow : function(index, opts) {
        var row = new Row(index, opts);
        this.rows.push(row);
        return row;
    },
    /**
     * @param min {number} the lower index
     * @param max {number} the upper index
     * @param width {number} the column width
     * @param bestFit {boolean} true if the column should fit to the contents (optional)
     * @desc Defines the width for the specified columns.
     */
    setColumn : function(min, max, width, bestFit) {
        var col = new Column(min, max, width, bestFit);
        this.cols.push(col);
        return col;
    },
    /**
     * @param range {string} range identifier (eg. 'A1:B2')
     * @desc merge range of cells
     */
    mergeCell : function(range) {
        var merge = new MergeCell(range);
        this.merges.push(merge);
        return merge;
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
        if (this.merges.length > 0) {
            var mergeCells = root.ele("mergeCells");
            mergeCells.att('count', this.merges.length);
            for (var ii in this.merges) {
              this.merges[ii].save(mergeCells);
            }
        }
    }
}

/**
 * @constructor
 * @class
 */
function Workbook () {
    this.sheets = [];
    this.relID = 1;
    this.styles = new CellStyles();
    this.numberFormats = new NumberFormats();
    this.fills = new Fills();
    this.fonts = new Fonts();
    this.borders = new Borders();
}
/**
 * @class
 * @desc This is the main interface for this module.
 */
Workbook.prototype = {
    constructor : Workbook,
    /**
     * @param name - the style name {string}
     * @returns a new {@linkcode CellStyle} with default values
     */
    createStyle : function(name) {
        return this.styles.create(name);
    },
    /**
     * @param name - the sheet name {string}
     * @returns a new {@linkcode Sheet}
     */
    addSheet : function(name) {
        var sheet = new Sheet(this.relID++, name);
        this.sheets.push(sheet);
        return sheet;
    },
    /**
     * @param column - numerical column index (1-676) {number}
     * @returns the Excel column index (A - ZZ)
     */
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
    /**
     * @param r - red value (0-255) {number}
     * @param g - green value (0-255) {number}
     * @param b - blue value (0-255) {number}
     * @returns a {@linkcode Color}
     */
    color : function(r,g,b) {
        return new Color(r,g,b,0);
    },
    /**
     * @param style - one of the predefined border styles {string}
     * @param color - the border color {Color}
     * @returns a BorderPr
     */
    createBorderPr : function(style, color) {
        return this.borders.createBorderPr(style, color);
    },
    /**
     * @param opts the number format options @type {NumberFormatOpts}
     * @returns a {NumberFormat}
     */
    addNumberFormat : function(value) {
        return this.numberFormats.add(value);
    },
    /**
     * @param opts the fonts options @type {FontOpts}
     * @returns a {Font}
     */
    addFont : function(opts) {
        return this.fonts.deriveFromDefault(opts);
    },

    /**
     * @typedef PatternFillOpts
     * @type {object}
     * @property {Color} fgColor - the foreground color
     * @property {Color} bgColor - the background color
     * @property {string} type - pattern type
     * @desc You should specifiy the type, fgColor and bgColor are optional.
     */

    /**
     * @param opts the pattern options @type {PatternFillOpts}
     * @returns a {PatternFill}
     */
    addPatternFill : function(opts) {
        return this.fills.addPatternFill(opts);
    },

    /**
     * @typedef BorderOpts
     * @type {object}
     * @property {BorderPr} top - the top border style
     * @property {BorderPr} bottom - the bottom border style
     * @property {BorderPr} left - the left border style
     * @property {BorderPr} right - the right border style
     * @desc All properties are optional. Create a border style {@linkcode BorderPr} with {@linkcode Workbook#createBorderPr}
     */

    /**
     * @param opts the border options @type {BorderOpts}
     * @returns a Border
     */
    addBorder : function(opts) {
        return this.borders.add(opts);
    },

    /**
     * @typedef StyleOpts
     * @type {object}
     * @property {NumberFormat} numFormat - the number format
     * @property {PatternFill} fill - the fill type
     * @property {Font} font - the font
     * @property {Border} border - the border type
     * @desc All properties are optional.
     *       Create a {@linkcode Border} with {@linkcode Workbook#addBorder}.
     *       Create a {@linkcode Font} with {@linkcode Workbook#addFont}.
     *       Create a {@linkcode PatternFill} with {@linkcode Workbook#addPatternFill}.
     *       Create a {@linkcode NumberFormat} with {@linkcode Workbook#addNumberFormat}.
     */

    /**
     * @param style the parent cell style {CellStyle}
     * @param opts the style options {StyleOpts}
     * @returns a Border
     */
    addStyle : function(style, opts) {
        return this.styles.derive(style, opts)
    },
    /**
     * @param name  the file name {string}
     * @desc saves the workbook as a Excel 2010 file.
     */
    save : function(fileName, cb) {

        var zip = new JSZip();
        var relsFolder   = zip.folder("_rels");
        var xlFolder     = zip.folder("xl");
        var xlRelsFolder = xlFolder.folder("_rels");
        var sheetsFolder = xlFolder.folder("worksheets");

        // create content-types
        this._saveContents(zip);

        // create main relationships
        this._saveMainRelations(relsFolder);

        // create sheets relationships
        this._saveWorkbookRelations(xlRelsFolder);

        // create styles and the workbook
        this._saveStyles(xlFolder);
        this._saveWorkbook(xlFolder);

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
        fs.writeFile(fileName, data, 'binary', cb);
        return fileName;
    },
    // internal stuff below this line
    _saveContents : function(zipFolder) {
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
    _saveMainRelations : function(zipFolder) {
        var mainRelations = builder.create("Relationships",{version: '1.0', encoding: 'utf-8'});
        mainRelations.att("xmlns",EXCEL_SCHEMA_FILE_REL);
        mainRelations.ele("Relationship").att("Id","rId" + (this.relID++)).att("Type",EXCEL_SCHEMA_REL_TYPE_WB).att("Target","/xl/workbook.xml");
        zipFolder.file(".rels", mainRelations.end({ pretty: true, indent: '  ', newline: '\n' }));
    },
    _saveWorkbook : function(zipFolder) {
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
    _saveWorkbookRelations : function(zipFolder) {

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
    _saveStyles : function(zipFolder) {

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
    }
}

