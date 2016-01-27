var builder = require('xmlbuilder');
var fs = require('fs');
var _ = require('underscore');
var archiver = require('archiver');

var stringtable = require('./modules/stringtable');

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
exports.PATTERN_TYPE_DARK_UP          = "darkUp";
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
 * @constant CELL_TYPE_MSDATE
 * @desc number type for cells
 */
exports.CELL_TYPE_MSDATE   = "n";
/**
 * @constant CELL_TYPE_BOOLEAN
 * @desc boolean type for cells
 */
exports.CELL_TYPE_BOOLEAN   = "b";
/**
 * @constant CELL_TYPE_ERROR
 * @desc error type for cells
 */
exports.CELL_TYPE_ERROR   = "e";
/**
 * @constant CELL_TYPE_STRING
 * @desc inline string (not a shared string ) type for cells
 */
exports.CELL_TYPE_STRING   = "inlineStr";
/**
 * @constant CELL_TYPE_STRING_TAB
 * @desc string which will be saved in a string table
 */
exports.CELL_TYPE_STRING_TAB = "s";
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
 * @constant AUTO_FILTER_COLOR_FONT
 * @desc to filter by cell font color
 */
exports.AUTO_FILTER_COLOR_FONT = 0;

/**
 * @constant AUTO_FILTER_COLOR_FILL
 * @desc to filter by cell fill color
 */
exports.AUTO_FILTER_COLOR_FILL = 1;

/**
 * Creates a {@linkcode Workbook} instance
 */
exports.createWorkbook = function() {
    return new Workbook();
};

// all internal stuff below this line

// Note that all objects should have a 'save' method which creates the
// complete XML-element structure.
// This save method has one mandatory parameter: the parent XML-element


// schemas and content-types
var EXCEL_SCHEMA_MAIN          = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var EXCEL_SCHEMA_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
var EXCEL_SCHEMA_FILE_REL      = "http://schemas.openxmlformats.org/package/2006/relationships";
var EXCEL_SCHEMA_DOC_REL       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
var EXCEL_SCHEMA_REL_TYPE_WB   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
var EXCEL_SCHEMA_REL_STYLES    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
var EXCEL_SCHEMA_REL_SHEET     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
var EXCEL_SCHEMA_REL_STRTAB    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
var EXCEL_SCHEMA_STYLES        = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

var EXCEL_TYPE_REL             = "application/vnd.openxmlformats-package.relationships+xml";
var EXCEL_TYPE_WORKBOOK        = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
var EXCEL_TYPE_SHEET           = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
var EXCEL_TYPE_STYLES          = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
var EXCEL_TYPE_STRINGTABLE     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

// the A-Z array, gets filled as soon as it is needed
var COLUMN_IDS = null;

// creates the xf element
function _writeStyle(ele, style, id) {
    var xf = ele.ele("xf");

    xf.att("numFmtId", style.numFormat != null ? style.numFormat.id : 0);
    xf.att("fontId", style.font != null ? style.font.id : 0);
    xf.att("fillId", style.fill != null ? style.fill.id : 0);
    xf.att("borderId", style.border != null ? style.border.id : 0);

    if (id != null) {
        xf.att("xfId",id);
    }
    if (style.numFormat != 0) {
        xf.att("applyNumberFormat",1);
    }
    if (style.font != 0) {
        xf.att("applyFont",1);
    }
    if (style.border != 0) {
        xf.att("applyBorder",1);
    }
    if (style.fill != 0) {
        xf.att("applyFill",1);
    }
    if (style.alignment != null) {
        xf.att("applyAlignment",1);
        style.alignment.save(xf);
    }
}

// adds the styles from the given list as xf elements to ele
// if all is true every style is added otherwise only those which
// don't have a parent style
function _writeXF(ele, styles, all) {

    styles.forEach(function(style) {
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
    });
}

// adds cellStyle elements to cellStyles for all styles without a parent style
function _writeCellStyles(cellStyles, styles) {
    styles.forEach(function(style) {
        if (style.parentStyle == null) {
            cellStyles.ele("cellStyle").att("name", style.name).att("xfId", style.id).att("builtinId",0);
        }
    });
}

/**
 * The AutoFilter definition for a single column.
 * 
 * @param colId
 * @constructor
 */
function FilterColumn(colId) {
    this.colId = colId;
    this.filters = [];
    this.hiddenButton = false;
    this.showButton = true;
    this.colorFilter = null;
}

FilterColumn.prototype = {
    constructor : FilterColumn,
    /**
     * Adds a simple value filter.
     * @param value
     * @returns {FilterColumn}
     */
    addFilter : function(value) {
        this.filters.push(value);
        return this;
    },
    /**
     * Sets the color to filter for.
     * @param fill  use a PatternFill
     * @param type  the type, either AUTO_FILTER_COLOR_FILL or AUTO_FILTER_COLOR_FONT
     * @returns {FilterColumn}
     */
    setColorFilter : function(fill, type) {
        this.colorFilter = {fill:fill, type:type};
        return this;
    },
    setShowButton : function(flag) {
        this.showButton = flag;
        return this;
    },
    setHiddenButton : function(flag) {
        this.hiddenButton = flag;
        return this;
    },
    save : function(parent) {
        var xFilterColumn = parent.ele("filterColumn");
        if (this.filters.length != 0) {
            var xFilters = xFilterColumn.ele("filters");
            this.filters.forEach(function (filter) {
                xFilters.ele("filter").att("val", filter);
            });
        }
        xFilterColumn.att("colId", this.colId);
        xFilterColumn.att("hiddenButton", this.hiddenButton);
        xFilterColumn.att("showButton", this.showButton);
        
        if (this.colorFilter != null) {
            xFilterColumn.ele("colorFilter")
                .att("dxfId", this.colorFilter.fill.getDxfsId())
                .att("cellColor", this.colorFilter.type);
        }
    }
};

function AutoFilter(ref) {
    this.ref = ref;
    this.filters = [];
}

AutoFilter.prototype = {
    constructor : AutoFilter,
    /**
     * Add a filter definition for the given column within the AutoFilter range
     * NOTE: As far as I know there is no way to let Excel apply the filter upon loading the file.
     * 
     * @param colId zero based ID
     * @returns {FilterColumn}
     */
    addFilter : function(colId) {
        var fCol = new FilterColumn(colId);
        this.filters.push(fCol);
        return fCol;
    },
    save : function(parent) {
        var xAutoFilter = parent.ele("autoFilter");
        xAutoFilter.att("ref", this.ref);
        this.filters.forEach(function(filter) {
            filter.save(xAutoFilter);
        });
    }
};


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
};

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
            col.att("customWidth", "1");
        }
        if (this._bestFit != undefined) {
            col.att("bestFit", this._bestFit);
        }
    }
};

function Font(opts, id) {
    // if something is changed here -> update clone and Fonts.deriveFromDefault as well
    this.name = opts.name;
    this.size = opts.size;
    this.bold = opts.bold;
    this.italic = opts.italic;
    this.color = opts.color;
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
        if (this.italic) {
            font.ele("i");
        }
        if (this.color != undefined) {
            this.color.save(font,"color");
        }
    },
    clone : function() {
        var font = new Font({name:this.name, size:this.size, bold:this.bold, italic:this.italic, color:this.color}, this.id);
        return font;
    }
};
function Fonts() {
    this.fontId = 0;
    this.fonts = [];
    // create a default font
    this.defaultOpts = {bold: false, italic: false, size: 11, name:"Calibri"};
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
        newFont.italic = opts.italic;
        newFont.size = opts.size;
        newFont.name = opts.name;
        newFont.color = opts.color;
        this.fonts.push(newFont);
        return newFont;
    },
    save : function(stylesheet) {
        var ele = stylesheet.ele("fonts");
        ele.att("count", this.fonts.length);
        this.fonts.forEach(function(font) {
            font.save(ele);
        });
    }
};

function CellAlignment(h, v, wrap) {
    this.h = h;
    this.v = v;
    this.wrap = !!wrap;
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
        if (this.rotation != null) {
            el.att("textRotation", this.rotation);
        }
        if (this.wrap) {
            el.att("wrapText", "1");
        }
    }
};

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
};

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
};

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
        this.borders.forEach(function(border) {
            border.save(ele);
        })
    }
};

/**
 * @class
 */
function PatternFill(opts, id) {
    this.fgColor = opts.fgColor;
    this.bgColor = opts.bgColor;
    this.type = opts.type;
    this.id = id;
    this.dxfsId = 0;
}
PatternFill.prototype = {
    constructor : PatternFill,
    getId : function() {
        return this.id;
    },
    getDxfsId : function() {
        return this.dxfsId;
    },
    setDxfsId : function(id) {
        this.dxfsId = id;
    },
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
};

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
        this.fills.forEach(function(fill) {
            fill.save(ele);
        });
    },
    saveDxfs : function(stylesheet) {
        var dxfs = stylesheet.ele("dxfs");
        dxfs.att("count", this.fills.length);
        this.fills.forEach(function(fill, index) {
            fill.save(dxfs.ele("dxf"));
            fill.setDxfsId(index);
        });
    }
};

/**
 * @class
 */
function NumberFormat(id, formatCode) {
    this.id = id;
    this.formatCode = formatCode;
}
NumberFormat.prototype = {
    constructor : NumberFormat
};

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
};

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
    this.alignment = new CellAlignment(null, null, false);
}
CellStyle.prototype = {
    constructor : CellStyle,
    setAlignment : function(h, v, w) {
        this.alignment = new CellAlignment(h, v, w);
        return this;
    },
    setWrapText : function(w) {
        this.alignment = new CellAlignment(this.alignment.h, this.alignment.v, w);
        return this;
    },
    setFont : function(font) {
        this.font = font;
        return this;
    },
    setBorder : function (border) {
        this.border = border;
        return this;
    },
    apply : function(style) {
        this.parentStyle = style;
        this.fill = style.fill;
        this.font = style.font;
        this.numFormat = style.numFormat;
        this.border = style.border;
        this.alignment = style.alignment;
    }
};

function CellStyles() {
    this.nextStyleId = 0;
    this.styles = [];
}
CellStyles.prototype = {
    constructor : CellStyles,
    /**
     * Creates a new CellStyle with the given name IF such a CellStyle is not already defined.
     * The new or existing CellStyle is returned.
     * Note that a new CellStyle is always returned IF name is null.
     * @param name
     * @returns {*|null}
     */
    create : function(name) {
        var style = this.getStyle(name);
        if (style == null) {
            style = new CellStyle(name, this.nextStyleId++);
            this.styles.push(style);
        }
        return style;
    },
    /**
     * Creates a new style based on an existing style.
     * The new style gets a name like "elxmlStyle_id" where id is the unique style id.
     * 
     * @param cellStyle     the style to derive from
     * @param opts          options for the new style
     * @returns {CellStyle} the new style
     */
    derive : function(cellStyle, opts) {
        opts = (opts == undefined ? {} : opts);
        _.defaults(opts, {numFormat: null, fill: null, font: null, border: null});
        
        var style = new CellStyle("elxmlStyle_" + this.nextStyleId, this.nextStyleId);
        this.nextStyleId++;
        
        style.apply(cellStyle);

        if (opts.numFormat != null) {
            style.numFormat = opts.numFormat;
        }
        if (opts.font != null) {
            style.font = opts.font;
        }
        if (opts.fill != null) {
            style.fill = opts.fill;
        }
        if (opts.border != null) {
            style.border = opts.border;
        }
        this.styles.push(style);
        return style;
    },
    countDirectStyles : function() {
        var count = 0;
        this.styles.forEach(function(style) {
            if (style.parentStyle == null) {
                count++;
            }
        });
        return count;
    },
    /**
     * Returns the current number of CellStyle instances managed by this.
     * @returns {Number}
     */
    count : function() {
        return this.styles.length;
    },
    /**
     * Returns the internal array of CellStyle instances.
     * @returns {Array}
     */
    getStyles : function() {
        return this.styles;
    },
    /**
     * Returns a style by name.
     * @param name  the name of the style to return, if null no style is returned
     * @returns null if no style with the given name is found or name is null, otherwise an instance of CellStyle with the given name
     */
    getStyle : function(name) {
        var style = null;
        if (name != null) {
            this.styles.forEach(function (st) {
                if (st.name == name) {
                    style = st;
                }
            });
        }
        return style;
    }
};

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
        if (this.type !== undefined) {
            ele.att("t", this.type);
        } else {
            ele.att("t", "inlineStr");
        }
        if (this.formula != null) {
            var f = ele.ele("f");
            f.att("t", exports.CELL_FORMULA_NORMAL);
            f.t(this.formula);
        }
        if (this.value != null) {
            if (this.type == undefined || this.type == "inlineStr") {
                ele.ele("is").ele("t").t(this.value);
            } else if (this.type == "s") {
                ele.ele("v").t(this.row.strTable.addString(this.value));
            } else {
                ele.ele("v").t(this.value);
            }
        }
        if (this.style != null) {
            ele.att("s", this.style.id);
        }
    }
};

/**
 * @class
 */
function Row(index, opts, strTable) {
    opts = (opts == undefined ? {} : opts);
    _.defaults(opts, {height:-1});
    this.height = opts.height;
    this.index = index;
    this.strTable = strTable;
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
    /**
     * @param style {CellStyle} to apply for the cells of the row
     * @desc Assigns the given style to all cells the row currently(!) holds.
     */
    setStyleForAllCells : function(style) {
        this.cells.forEach(function(cell) {
            cell.setStyle(style);
        })  
    },
    save : function(sheetData) {
        var ele = sheetData.ele("row");
        ele.att("r", this.index);
        if (this.height != -1) {
            ele.att("ht", this.height);
            ele.att("customHeight", 1);
        }
        this.cells.forEach(function(cell) {
            cell.save(ele);
        });
    }
};

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
};

/**
 * @class
 * @param id {number} the sheet id
 * @param name {string} the sheet name
 * @param strTable the string table, used to reuse strings
 * @desc Don't use this constructor, use {@linkcode Workbook#addSheet} instead.
 */
function Sheet(id, name, strTable) {
    this.id = id;
    this.name = name;
    this.rows = [];
    this.cols = [];
    this.merges = [];
    this.strTable = strTable;
    this.autoFilter = null;
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
        var row = new Row(index, opts, this.strTable);
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
    /**
     * 
     * @param range {string} range identifier (eg. 'A1:B2'), if null auto filter is disabled
     * @desc enable auto filter for the given range 
     */
    setAutoFilter : function(range) {
        if (range == null) {
            this.autoFilter = null;
        } else {
            this.autoFilter = new AutoFilter(range);
        }
        return this.autoFilter;
    },
    getAutoFilter : function() {
        return this.autoFilter;
    },
    save : function(root) {
        if (this.cols.length > 0) {
            var colsEle = root.ele("cols");
            this.cols.forEach(function(col) {
                col.save(colsEle);
            });
        }
        if (this.rows.length > 0) {
            var sheetData = root.ele("sheetData");
            this.rows.forEach(function(row) {
                row.save(sheetData);
            });
        }
        if (this.autoFilter != null) {
            this.autoFilter.save(root);
        }
        if (this.merges.length > 0) {
            var mergeCells = root.ele("mergeCells");
            mergeCells.att('count', this.merges.length);
            this.merges.forEach(function(merge) {
                merge.save(mergeCells);
            });
        }
    }
};

/**
 * @constructor
 * @class
 * @param opts  options - standard:true = create a "Standard" style
 * 
 * Each Workbook has a default CellStyle named "Standard", use this to derive
 * new styles.
 */
function Workbook (opts) {
    opts = (opts == undefined ? {} : opts);
    _.defaults(opts, {standard:true});

    this.sheets = [];
    this.relID = 1;
    
    this.styles = new CellStyles();
    if (opts.standard == true) {
        this.styles.create("Standard");
    }
    
    this.numberFormats = new NumberFormats();
    this.fills = new Fills();
    this.fonts = new Fonts();
    this.borders = new Borders();
    this.strTable = stringtable.createStringTable();
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
     * Returns a style by name.
     * @param name  
     * @returns {*|null}
     */
    getStyle : function(name) {
        return this.styles.getStyle(name);
    },
    /**
     * @param name - the sheet name {string}
     * @returns {@linkcode Sheet}
     */
    addSheet : function(name) {
        var sheet = new Sheet(this.relID++, name, this.strTable);
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
     * @returns {@linkcode Color}
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
    addNumberFormat : function(opts) {
        return this.numberFormats.add(opts);
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
     * @desc You should specify the type, fgColor and bgColor are optional.
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
     * @param fileName  the file name {string}
     * @param callback  gets an argument (err) when an error occurs.
     * @desc saves the workbook as a Excel 2010 file.
     */
    save : function( fileName, callback ) {
        var output = fs.createWriteStream( fileName );
        this.saveToStream(output, callback);
        return fileName;
    },
    /**
     * @param output  the file name {string}
     * @param callback  gets an argument (err) when an error occurs.
     * @desc saves the workbook as a Excel 2010 file.
     */
    saveToStream : function( output, callback ) {

        var archive = archiver( 'zip' );

        // for debugging
        /*output.on('close', function() {
         console.log(archive.pointer() + ' total bytes');
         console.log('archiver has been finalized and the output file descriptor has closed.');
         });*/

        archive.on('error', function(err) {
            callback(err);
        });

        output.on('finish', function() {
            callback(null);
        });

        archive.pipe( output );

        var relsFolder   = "_rels";
        var xlFolder     = "xl";
        var xlRelsFolder = "xl/_rels";
        var sheetsFolder = "xl/worksheets";

        archive.append(null, { name: relsFolder + '/' });
        archive.append(null, { name: xlFolder + '/' });
        archive.append(null, { name: xlRelsFolder + '/' });
        archive.append(null, { name: sheetsFolder + '/' });

        // create content-types
        this._saveContents( archive );

        // create main relationships
        this._saveMainRelations( archive, relsFolder );

        // create sheets relationships
        this._saveWorkbookRelations( archive, xlRelsFolder);

        // create styles and the workbook
        this._saveStyles( archive, xlFolder );
        this._saveWorkbook( archive, xlFolder );

        // create sheet files
        this.sheets.forEach(function(sheet) {
            var root = builder.create("worksheet", {version: '1.0', encoding: 'utf-8'});
            root.att("xmlns", EXCEL_SCHEMA_MAIN);
            sheet.save(root);

            var xmlString = root.end({ pretty: false });
            archive.append( xmlString, { name: sheetsFolder + "/" + "sheet" + sheet.id + ".xml" });
        });

        // save the StringTable
        this.strTable.save( archive, xlFolder );

        archive.finalize();
    },
    
    // internal stuff below this line
    _saveContents : function( archive ) {
        var contents = builder.create("Types",{version: '1.0', encoding: 'utf-8'});
        contents.att("xmlns",EXCEL_SCHEMA_CONTENT_TYPES);

        contents.ele("Default").att("Extension","xml").att("ContentType",EXCEL_TYPE_WORKBOOK);
        contents.ele("Default").att("Extension","rels").att("ContentType",EXCEL_TYPE_REL);
        contents.ele("Override").att("PartName","/xl/styles.xml").att("ContentType",EXCEL_TYPE_STYLES);
        contents.ele("Override").att("PartName","/xl/sharedStrings.xml").att("ContentType",EXCEL_TYPE_STRINGTABLE);

        this.sheets.forEach(function(sheet) {
            var sheetName = "/xl/worksheets/sheet" + sheet.id + ".xml";
            contents.ele("Override").att("PartName",sheetName).att("ContentType",EXCEL_TYPE_SHEET);
        });

        var xmlString = contents.end({ pretty: false });
        archive.append( xmlString, { name: "[Content_Types].xml" });
    },
    _saveMainRelations : function( archive, zipFolder ) {
        var mainRelations = builder.create("Relationships",{version: '1.0', encoding: 'utf-8'});
        mainRelations.att("xmlns",EXCEL_SCHEMA_FILE_REL);
        mainRelations.ele("Relationship").att("Id","rId" + (this.relID++)).att("Type",EXCEL_SCHEMA_REL_TYPE_WB).att("Target","/xl/workbook.xml");
        archive.append( mainRelations.end( { pretty: false } ), { name: zipFolder + "/" + ".rels" });
    },
    _saveWorkbook : function( archive, zipFolder ) {
        var workbookRelations = builder.create("workbook",{version: '1.0', encoding: 'utf-8'});
        workbookRelations.att("xmlns",EXCEL_SCHEMA_MAIN).att("xmlns:r",EXCEL_SCHEMA_DOC_REL);

        var sheetsEle = workbookRelations.ele("sheets");
        this.sheets.forEach(function(sheet) {
            sheetsEle.ele("sheet").att("name",sheet.name).att("sheetId",sheet.id).att("r:id","rId" + sheet.id);
        });

        workbookRelations.ele("calcPr").att("fullCalcOnLoad",true);
        
        var xmlString = workbookRelations.end({ pretty: false });
        archive.append( xmlString, { name: zipFolder + "/" + "workbook.xml" });
    },
    _saveWorkbookRelations : function( archive, zipFolder ) {

        var relations = builder.create("Relationships",{version: '1.0', encoding: 'utf-8'});
        relations.att("xmlns",EXCEL_SCHEMA_FILE_REL);

        var sheetRel = relations.ele("Relationship");
        sheetRel.att("Type",EXCEL_SCHEMA_REL_STYLES);
        sheetRel.att("Target","/xl/styles.xml");
        sheetRel.att("Id","rId" + (this.relID++));

        sheetRel = relations.ele("Relationship");
        sheetRel.att("Type",EXCEL_SCHEMA_REL_STRTAB);
        sheetRel.att("Target","/xl/sharedStrings.xml");
        sheetRel.att("Id","rId" + (this.relID++));

        this.sheets.forEach(function(sheet) {
            var sheetRel = relations.ele("Relationship");
            sheetRel.att("Type",EXCEL_SCHEMA_REL_SHEET);
            sheetRel.att("Target","/xl/worksheets/sheet" + sheet.id + ".xml");
            sheetRel.att("Id","rId" + sheet.id);
        });

        var xmlString = relations.end({ pretty: false });
        archive.append( xmlString, { name: zipFolder + "/" + "workbook.xml.rels" });
    },
    _saveStyles : function( archive, zipFolder ) {

        var stylesheet = builder.create("styleSheet",{version: '1.0', encoding: 'utf-8'});
        stylesheet.att("xmlns",EXCEL_SCHEMA_STYLES);

        // number formats
        var numFmts = stylesheet.ele("numFmts");
        numFmts.att("count", this.numberFormats.formats.length);
        this.numberFormats.formats.forEach(function(fmt) {
            numFmts.ele("numFmt").att("numFmtId", fmt.id).att("formatCode", fmt.formatCode);
        });

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

        this.fills.saveDxfs(stylesheet);
        
        var xmlString = stylesheet.end({ pretty: false });
        archive.append( xmlString, { name: zipFolder + "/" + "styles.xml" });
    }
};


