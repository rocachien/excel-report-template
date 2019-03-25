/**
 * @copyright 2019
 * @author rocachien
 * @create 2019/08/06 19:14
 * @update 2019/08/06 19:14
 */
'use strict';

const etree = require('elementtree');
const XlsxTemplate = require('xlsx-template');

/*eslint-disable */
class ExcelReportTemplate extends XlsxTemplate {
    constructor(data) {
        super(data);
    }

    extractPlaceholders (string) {
        // Yes, that's right. It's a bunch of brackets and question marks and stuff.
        var re = /\${(table:)?(.+?)(?:\.(.+?))?}/gm;
        var match = null, matches = [];

        while((match = re.exec(string)) !== null) {
            var type = match[1] ? match[1].slice(0, -1) : 'normal';
            matches.push({
                placeholder: match[0],
                type: type,
                name: match[2],
                key: match[3],
                full: match[0].length === string.length
            });
        }

        return matches;
    }

    getSubstitutionWithNormal (substitutions, placeholder) {
        if (placeholder && placeholder.type === 'normal') {
            return substitutions[0][placeholder.name];
        }

        return '';
    }

    substitute (sheetName, substitutions) {
        var self = this;
        var sheet = self.loadSheet(sheetName);
        var dimension = sheet.root.find("dimension"),
            sheetData = sheet.root.find("sheetData"),
            currentRow = null,
            totalRowsInserted = 0,
            totalColumnsInserted = 0,
            namedTables = self.loadTables(sheet.root, sheet.filename),
            rows = [];

        sheetData.findall("row").forEach(function(row) {
            row.attrib.r = currentRow = self.getCurrentRow(row, totalRowsInserted);
            rows.push(row);

            var cells = [],
                cellsInserted = 0,
                newTableRows = [];

            row.findall("c").forEach(function(cell) {
                var appendCell = true;
                cell.attrib.r = self.getCurrentCell(cell, currentRow, cellsInserted);

                // If c[@t="s"] (string column), look up /c/v@text as integer in
                // `this.sharedStrings`
                if(cell.attrib.t === "s") {

                    // Look for a shared string that may contain placeholders
                    var cellValue   = cell.find("v"),
                        stringIndex = parseInt(cellValue.text, 10),
                        string      = self.sharedStrings[stringIndex];

                    if(string === undefined) {
                        return;
                    }

                    // Loop over placeholders
                    self.extractPlaceholders(string).forEach(function(placeholder) {
                        // Only substitute things for which we have a substitution
                        var newCellsInserted = 0;

                        if(placeholder.full && placeholder.type === "table") {
                            newCellsInserted = self.substituteTable(
                                row, newTableRows,
                                cells, cell,
                                namedTables, substitutions, placeholder.name
                            );

                            // don't double-insert cells
                            // this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
                            if (newCellsInserted !== 0 || substitutions.length) {
                                if (substitutions.length === 1) {
                                    appendCell = true;
                                }
                                if (substitutions[0][placeholder.key] instanceof Array) {
                                    appendCell = false;
                                }
                            }

                            // Did we insert new columns (array values)?
                            if(newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        } else {
                            var substitution = self.getSubstitutionWithNormal(substitutions, placeholder);
                            string = self.substituteScalar(cell, string, placeholder, substitution);
                        }
                    });
                }

                // if we are inserting columns, we may not want to keep the original cell anymore
                if(appendCell) {
                    cells.push(cell);
                }

            }); // cells loop

            // We may have inserted columns, so re-build the children of the row
            self.replaceChildren(row, cells);

            // Update row spans attribute
            if(cellsInserted !== 0) {
                self.updateRowSpan(row, cellsInserted);

                if(cellsInserted > totalColumnsInserted) {
                    totalColumnsInserted = cellsInserted;
                }

            }

            // Add newly inserted rows
            if(newTableRows.length > 0) {
                newTableRows.forEach(function(row) {
                    rows.push(row);
                    ++totalRowsInserted;
                });
                self.pushDown(self.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
            }

        }); // rows loop

        // We may have inserted rows, so re-build the children of the sheetData
        self.replaceChildren(sheetData, rows);

        // Update placeholders in table column headers
        self.substituteTableColumnHeaders(namedTables, substitutions);

        // Update placeholders in hyperlinks
        self.substituteHyperlinks(sheet.filename, substitutions);

        // Update <dimension /> if we added rows or columns
        if(dimension) {
            if(totalRowsInserted > 0 || totalColumnsInserted > 0) {
                var dimensionRange = self.splitRange(dimension.attrib.ref),
                    dimensionEndRef = self.splitRef(dimensionRange.end);

                dimensionEndRef.row += totalRowsInserted;
                dimensionEndRef.col = self.numToChar(self.charToNum(dimensionEndRef.col) + totalColumnsInserted);
                dimensionRange.end = self.joinRef(dimensionEndRef);

                dimension.attrib.ref = self.joinRange(dimensionRange);
            }
        }

        //Here we are forcing the values in formulas to be recalculated
        // existing as well as just substituted
        sheetData.findall("row").forEach(function(row) {
            row.findall("c").forEach(function(cell) {
                var formulas = cell.findall('f');
                if (formulas && formulas.length > 0) {
                    cell.findall('v').forEach(function(v){
                        cell.remove(v);
                    });
                }
            })
        });

        // Write back the modified XML trees
        self.archive.file(sheet.filename, etree.tostring(sheet.root));
        self.archive.file(self.workbookPath, etree.tostring(self.workbook));

        // Remove calc chain - Excel will re-build, and we may have moved some formulae
        if(self.calcChainPath && self.archive.file(self.calcChainPath)) {
            self.archive.remove(self.calcChainPath);
        }

        self.writeSharedStrings();
        self.writeTables(namedTables);
    }
}

/*eslint-enable */
module.exports = ExcelReportTemplate;
