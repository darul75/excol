// Imports
import  { Cell, NewCell } from '../Cell'
import { Range, CellValue } from '../range/Range'
import { toCoordinates, validateA1 } from '../converter/ToCoordinates'
import {Errors} from '../Error'
import {Spreadsheet} from '../spreadsheet/Spreadsheet';

// Interfaces
export interface Dimension {
  numRows?: number,
  numColumns?: number
}

export interface RangeType extends RowColumnRangeType {
  A1?: string;
}

export interface RowColumnRangeType extends Dimension {
  row?: number;
  column?: number;
}

export interface SheetConfig extends Dimension {
  id: number
  name: string
  cellValue?: CellValue
  numRows?: number
  numColumns?: number
}

// Constants
const DIMENSION = 1000;

// Class & methods
/**
 * Sheet class.
 */
export class Sheet {

  private _parent: Spreadsheet;
  private cfg: SheetConfig;
  private _range: Range;
  private _active: Range;
  private _mergedRanges: Range[] = [];

  private _id: number;
  public _name: string;
  public _numRows: number;
  public _numColumns: number;
  private _lastRow: number = 1;
  private _lastColumn: number = 1;

  /**
   * Constructor
   *
   * @param {object?} cfg Sheet configuration
   */
  constructor(cfg: SheetConfig, parent?: Spreadsheet) {
    this.cfg = cfg;

    if (parent) {
      this._parent = parent;
    }

    this.init(this.cfg);
  }

  /**
   * Initialize Sheet.
   *
   * @param {number?} numRows Row dimension size
   * @param {number?} numColumns Column dimension size
   * @param {object?} cell Cell value object
   */
  init({id, name, numRows = DIMENSION, numColumns = DIMENSION, cellValue} :
      { id: number, name: string, numRows?: number, numColumns?: number, cellValue?: CellValue }): void {

    this._id = id;
    this._name = name;
    this._numRows = numRows;
    this._numColumns = numColumns;

    this.initRange(new Range(1, numRows, 1, numColumns, null), cellValue);
  }

  /**
   * Appends a row to the spreadsheet.
   * This operation is atomic; it prevents issues where a user asks for the last row,
   * and then writes to that row,
   * and an intervening mutation occurs between getting the last row and writing to it.
   *
   * @param rowContents
   * @returns {Sheet}
   */
  public appendRow(rowContents: Object[]) : Sheet {
    const range = this.getRange({row: this._lastRow, column: 1, numRows: 1, numColumns: rowContents.length});

    range.setValues([rowContents]);

    return this;
  }

  /**
   * Sets the width of the given column to fit its contents.
   *
   * @param columnPosition
   */
  public autoResizeColumn(columnPosition: number) : Sheet {
    return this;
  }

  /**
   * Clears the sheet of content and formatting information.
   * or
   * Clears the sheet of contents and/or format, as specified with the given advanced options.
   *
   * @param options
   */
  public clear(options?: Object) : Sheet {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);

    //return this;
  }

  /**
   * Clears the sheet of contents, while preserving formatting information.
   */
  public clearContents() : Sheet {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);

    //return this;
  }

  /**
   * Clears the sheet of formatting, while preserving contents.
   */
  public clearFormats() : Sheet {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);

    //return this;
  }

  /**
   * Clears the sheet of all notes.
   */
  public clearNotes() : Sheet {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);

    //return this;
  }

  /**
   * Deletes the column at the given column position.
   *
   * @param columnPosition
   * @returns {any}
   */
  public deleteColumn(columnPosition: number, howMany?: number) {
    const newValues: any[][] = createArray(this._numRows, this._numColumns);
    const values: any[][] =this._range.values;
    const toDelNum: number = howMany ? howMany : 1;

    for (var r = 0; r < values.length; r++) {
      values[r].splice(columnPosition - 1, toDelNum);
      for(var i = 0; i < toDelNum; i++) values[r].push(null);
      newValues[r] = values[r];
    }

    const allRange: Range = this.getRange({row: 1, column: 1, numRows: this._numRows, numColumns: this._numColumns});

    allRange.setValues(newValues);

    return this;
  }

  /**
   * Deletes a number of columns starting at the given column position.
   * or
   * Deletes a number of columns starting at the given column position.
   *
   * @param columnPosition
   * @param howMany
   * @returns {Sheet}
   */
  public deleteColumns(columnPosition: number, howMany: number) {
    this.deleteColumn(columnPosition, howMany);

    return this;
  }

  /**
   * Deletes the row at the given column position.
   *
   * @param rowPosition
   * @returns {any}
   */
  public deleteRow(rowPosition: number, howMany?: number) {

    const values: any[][] =this._range.values;
    const toDelNum: number = howMany ? howMany : 1;

    values.splice(rowPosition - 1, toDelNum);

    for(var i = 0; i < toDelNum; i++) values.push(createArray(this._numColumns));

    const allRange: Range = this.getRange({row: 1, column: 1, numRows: this._numRows, numColumns: this._numColumns});

    allRange.setValues(values);

    return this;
  }

  /**
   * Deletes a number of rows starting at the given column position.
   *
   * @param rowPosition
   * @param howMany
   * @returns {Sheet}
   */
  public deleteRows(rowPosition: number, howMany: number) {
    this.deleteRow(rowPosition, howMany);
    return this;
  }

  /**
   * Returns the active range for the active sheet.
   */
  public getActiveRange() : Range {
    return this._active;
  }

  /**
   * Returns the position of the last column that has content.
   */
  public getLastColumn() : number {
    return this._lastColumn;
  }

  /**
   * Returns the position of the last row that has content.
   */
  public getLastRow() : number {
    return this._lastRow;
  }

  public getName() : string {
    return this._name;
  }

  /**
   * Returns the Spreadsheet that contains this sheet.
   *
   * @returns {Spreadsheet}
   */
  public getParent() : Spreadsheet {
    return this._parent;
  }

  /**
   * Get specific range by A1 or matrix.
   *
   * @param {string?} A1 A1 notation
   * @param {number?} row Row index
   * @param {number?} column Column index
   * @param {number?} numRows Number of rows
   * @param {number?} numColumns Number of columns
   *
   * @returns {Range} Range object
   */
  public getRange({A1, row, column, numRows, numColumns}:
    RangeType): Range {

    let range;

    if (A1) {

      if (!validateA1({A1})) {
        throw new Error(Errors.INCORRECT_A1_NOTATION);
      }

      const coordinates = toCoordinates({A1});

      if (coordinates.length > 1) {
        throw new Error(Errors.INCORRECT_GET_RANGE_USAGE);
      }

      const coordinate = coordinates[0];
      const firstElt = coordinate[0];

      // One Row or Column
      if (typeof firstElt === 'number') {
        range = GetSingleCellOrWholeRowColumnRange(this, coordinate);
      }
      else { // Over multiple row/columns
        range = GetAcrossMultipleRowColumnRange(this, coordinate);
      }

    }
    else {
      if (!row && !column && !numRows && !numColumns) {
        throw new Error(Errors.INCORRECT_RANGE_SIGNATURE);
      }

      numRows = numRows || 1;
      numColumns = numColumns || 1;

      range = new Range(row, numRows, column, numColumns, this);
    }

    this._active = range;

    return range;
  }

  /**
   * Get specific range by A1 or matrix.
   *
   * @param {string?} A1 A1 notation
   * @param {number?} row Row index
   * @param {number?} column Column index
   * @param {number?} numRows Number of rows
   * @param {number?} numColumns Number of columns
   *
   * @returns {Range} Range object
   */
  public getRanges({A1, row, column, numRows, numColumns}: RangeType): Range[] {

    const results = new Array<Range>();

    if (A1) {

      if (!validateA1({A1})) {
        throw new Error(Errors.INCORRECT_A1_NOTATION);
      }

      const coordinates = toCoordinates({A1});

      coordinates.forEach(elt => {

        const firstElt = elt[0];

        // One Row or Column
        if (typeof firstElt === 'number') {
          results.push(GetSingleCellOrWholeRowColumnRange(this, elt));
        }
        else { // Over multiple row/columns

          results.push(GetAcrossMultipleRowColumnRange(this, elt));
        }
      });
    }
    else {
      if (!row && !column && !numRows && !numColumns) {
        throw new Error(Errors.INCORRECT_RANGE_SIGNATURE);
      }

      numRows = numRows || 1;
      numColumns = numColumns || 1;

      results.push(new Range(row, numRows, column, numColumns, this));
    }

    if (results.length) {
      this._active = results[0];
    }

    return results;
  }

  /**
   * Returns the rectangular grid of values for this range starting at the given coordinates.
   *
   * @param startRow
   * @param startColumn
   * @param numRows
   * @param numColumns
   * @returns {any[][]}
   */
  public getSheetValues(startRow, startColumn, numRows, numColumns) : any[][] {
    const range: Range = this.getRange({row: startRow, column: startColumn, numRows: numRows, numColumns: numColumns});

    return range.values;
  }

  /**
   * Inserts a column after the given column position.
   *
   * @param afterPosition
   */
  public insertColumnAfter(afterPosition: number, howMany?: number) : Sheet {
    this._range.insertColumnAfter(afterPosition, howMany);

    return this;
  }

  /**
   * Inserts a column before the given column position.
   *
   * @param beforePosition
   */
  public insertColumnBefore(beforePosition: number, howMany?: number) : Sheet {
    this._range.insertColumnBefore(beforePosition, howMany);

    return this;
  }

  /**
   * Inserts a number of columns after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnsAfter(afterPosition: number, howMany: number) : Sheet {
    return this.insertColumnAfter(afterPosition, howMany);
  }

  /**
   * Inserts a number of columns before the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnsBefore(beforePosition: number, howMany: number) : Sheet {
    return this.insertColumnBefore(beforePosition, howMany);
  }

  /**
   * Inserts a row after the given column position.
   *
   * @param afterPosition
   */
  public insertRowAfter(afterPosition: number, howMany?: number) : Sheet {
    this._range.insertRowAfter(afterPosition, howMany);

    return this;
  }

  /**
   * Inserts a row before the given column position.
   *
   * @param beforePosition
   */
  public insertRowBefore(beforePosition: number, howMany?: number) : Sheet {
    this._range.insertRowBefore(beforePosition, howMany);

    return this;
  }

  /**
   * Inserts a number of rows after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowsAfter(afterPosition: number, howMany: number) : Sheet {
    return this.insertRowAfter(afterPosition, howMany);
  }

  /**
   * Inserts a number of rows before the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowsBefore(beforePosition: number, howMany: number) : Sheet {
    return this.insertRowBefore(beforePosition, howMany);
  }

  /**
   * Sets the active range for the active sheet.
   *
   * @param range
   */
  public setActiveRange(range) : Range {
    this._active = range;
    return this._active;
  }

  /**
   * Sets last column with content.
   *
   * @param column
   */
  public setLastColumn(column: number) : void {
    this._lastColumn = Math.max(column, this._lastColumn);
  }

  /**
   * Sets last row with content.
   *
   * @param row
   */
  public setLastRow(row: number) : void {
    this._lastRow = Math.max(row, this._lastRow);
  }

  initRange(range: Range, cellValue?: CellValue) {

    this._mergedRanges.length = 0;

    const rows = Array.apply(null, new Array(range.rowHeight));
    const isZeroNumber = cellValue === 0;

    const values = rows.map((row: Array<Cell>, rx: number) => {
      const cols = Array.apply(null, new Array<Cell>(range.columnWidth));
      return cols.map((col: CellValue, cx: number) => {
        // default init value
        const initialVal = isZeroNumber ? 0 : cellValue || {value: rx + '-' + cx};
        // init cell
        return NewCell(rx, cx, initialVal);
      });
    });

    const newRange: Range = new Range(1, range.rowHeight, 1, range.columnWidth, this, values);
    this._range = newRange;
    this._active = newRange;

  }

  public get range(): Range {
    return this._range;
  }

  public get mergedRanges(): Range[] {
    return this._mergedRanges;
  }

  public get id() : number {
    return this._id;
  }
}

const GetSingleCellOrWholeRowColumnRange = (sheet: Sheet, coordinates: Array<number>) : Range => {

  let row = coordinates[0];
  let column = coordinates[1];
  const allRows: boolean = row < 0;
  const allColumns: boolean = column < 0;

  row = allRows ? 1 : row;
  column = allColumns ? 1 : column;

  const rowHeight = allRows ? sheet._numRows : 1;
  const columnWidth = allColumns ? sheet._numColumns : 1;

  return new Range(row, rowHeight, column, columnWidth, sheet);
};

const GetAcrossMultipleRowColumnRange = (sheet: Sheet, coordinates: Array<Array<number>>) : Range => {
  // start cell
  let startCell = coordinates[0];
  let endCell = coordinates[1];

  // row bigger => inverse
  if (endCell[0] >= 0 && startCell[0] > endCell[0]) {
    const tmp = endCell;
    endCell = startCell;
    startCell = tmp;
  }

  // column bigger => inverse
  if (endCell[1] >= 0 && startCell[1] > endCell[1] ) {
    const tmp = endCell;
    endCell = startCell;
    startCell = tmp;
  }

  // Edge cell coordinates
  let startCellRow = startCell[0];
  let endCellRow = endCell[0];

  let startCellColumn = startCell[1];
  let endCellColumn = endCell[1];

  // Size
  let rowHeight = endCell[0] - startCell[0];
  let columnWidth = endCell[1] - startCell[1];

  // Handle full rows, columns
  const allRows: boolean = startCellRow < 0 || endCellRow < 0;
  const allColumns: boolean = startCellColumn < 0 || endCellColumn < 0;

  if (allRows || allColumns) {
    const isEndCellAllRows = endCellRow < 0 && startCellRow > 0;
    const isStartCellAllRows = endCellRow > 0 && startCellRow < 0;

    if (isEndCellAllRows) {
      rowHeight = sheet._numRows - startCellRow + 1;
    }
    else if (isStartCellAllRows) {
      rowHeight = sheet._numRows - endCellRow + 1;
    }
    else if (endCellRow < 0 && startCellRow < 0) {
      startCellRow = 1;
      rowHeight = sheet._numRows;
    }
    else {
      rowHeight = endCellRow - startCellRow + 1;
    }

    const isEndCellAllColumns = endCellColumn < 0 && startCellColumn > 0;
    const isStartCellAllColumns = endCellColumn > 0 && startCellColumn < 0;

    if (isEndCellAllColumns) {
      columnWidth = sheet._numColumns - startCellColumn + 1;
    }
    else if (isStartCellAllColumns) {
      columnWidth = sheet._numColumns - endCellColumn + 1;
    }
    else if (endCellColumn < 0 && startCellColumn < 0) {
      startCellColumn = 1;
      columnWidth = sheet._numColumns;
    }
    else {
      columnWidth = endCellColumn - startCellColumn + 1;
    }
  }
  else {
    rowHeight = endCellRow - startCellRow + 1;
    columnWidth = endCellColumn - startCellColumn + 1;
  }

  return new Range(startCellRow, rowHeight, startCellColumn, columnWidth, sheet);
};

function createArray(length: number, length2?: number) {
  var arr = new Array(length || 0),
    i = length;

  if (arguments.length > 1) {
    var args = Array.prototype.slice.call(arguments, 1);
    while(i--) arr[length-1 - i] = createArray.apply(this, args);
  }

  return arr;
}
