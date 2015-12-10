var builder = require('xmlbuilder');
var _ = require('underscore');

/**
 * Created by Arndt Teinert on 13.08.14.
 */

/* Example from the spec:
 <sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="8" uniqueCount="4">
    <si>
        <t>United States</t>
    </si>
    <si>
        <t>Seattle</t>
    </si>
    <si>
        <t>Denver</t>
    </si>
    <si>
        <t>New York</t>
    </si>
</sst>
*/

exports.createStringTable = function() {
    var strTable = new StringTable();
    return strTable;
};


// internal stuff below this line

var EXCEL_SCHEMA_STRTAB        = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/**
 * @constructor
 */
function StringTable() {
    this.strTable = {};
    this.strings = [];
}

/**
 * @class
 */
StringTable.prototype = {
    constructor : StringTable,
    addString : function(value) {
        if (value in this.strTable) {
            return this.strTable[value];
        } else {
            var id = _.size(this.strTable);
            this.strTable[value] = id;
            this.strings.push(value);
            return id;
        }
    },
    save: function( archive, zipFolder ) {
        var strtab = builder.create("sst", {version: '1.0', encoding: 'utf-8'});
        strtab.att("xmlns", EXCEL_SCHEMA_STRTAB);
        for (var value in this.strings) {
            strtab.ele("si").ele("t").t(this.strings[value]);
        }
        var xmlString = strtab.end({ pretty: false });
        archive.append(xmlString, { name: zipFolder + "/" + "sharedStrings.xml" });
    }
};
