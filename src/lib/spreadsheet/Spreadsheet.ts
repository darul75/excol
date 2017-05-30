// Imports
import { Range, CellValue } from '../range/Range'
import { Errors } from '../utils/Error'
import { Sheet } from '../sheet/Sheet';
import { SheetConfig } from '../sheet/Sheet';
import { NamedRange } from '../range/NamedRange';
import { User } from '../user/User';

export interface SpreadsheetConfig {
    name: string,
    numRows?: number,
    numColumns?: number
};

// Constants
const DIMENSION = 1000;

const sheetConfig: SheetConfig = {
  id: 0,
  name: 'Sheet',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

// Class & methods
/**
 * Spreadsheet class.
 */
export class Spreadsheet {

  private cfg: SpreadsheetConfig;
  private _sheets: Sheet[] = [];
  private _namedRanges: NamedRange[] = [];
  private _activeSheet: Sheet;
  private _editors: User[] = [];
  private _viewers: User[] = [];

  private _id: number;
  private _name: string;
  public _numRows: number;
  public _numColumns: number;

  /**
   * Constructor
   *
   * @param {object?} cfg Spreadsheet configuration
   */
  constructor(cfg: SpreadsheetConfig) {
    this.cfg = cfg;

    this.init(this.cfg);
  }

  /**
   * Initialize Sheet.
   *
   * @param {number?} numRows Row dimension size
   * @param {number?} numColumns Column dimension size
   * @param {object?} cell Cell value object
   */
  init({name, numRows = DIMENSION, numColumns = DIMENSION, cellValue} :
    { name: string, numRows?: number, numColumns?: number, cellValue?: CellValue }): void {
    this._id = new Date().getTime();
    this._name = name;
    this._numRows = numRows;
    this._numColumns = numColumns;

    this.insertSheet();
  }

  /**
   * * Adds the given user to the list of editors for the Spreadsheet.
   * If the user was already on the list of viewers, this method promotes the user out of the list of viewers.
   *
   * Adds the given user to the list of editors for the Spreadsheet.
   *
   * @param user
   * @returns {Spreadsheet}
   */
  public addEditor(user: User | string) : Spreadsheet {
    if (user instanceof User) {
      this._editors.push(user);
    }
    if (typeof user === 'string') {
      this._editors.push(new User(user));
    }

    this._editors.push();
    return this;
  }

  /**
   * Adds the given array of users to the list of editors for the Spreadsheet.
   *
   * @param emailAddresses
   * @returns {Spreadsheet}
   */
  public addEditors(emailAddresses: string[]) : Spreadsheet {
    for(var email of emailAddresses) {
      this._editors.push(new User(email));
    }
    return this;
  }

  /**
   * Creates a new menu in the Spreadsheet UI.
   *
   * @param name
   * @param subMenus
   */
  public addMenu(name: string, subMenus: Object[]) : void {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Adds the given user to the list of viewers for the Spreadsheet.
   *
   * @param user
   * @returns {Spreadsheet}
   */
  public addViewer(user: User | string) : Spreadsheet {
    if (user instanceof User) {
      this._viewers.push(user);
    }
    if (typeof user === 'string') {
      this._viewers.push(new User(user));
    }

    return this;
  }

  /**
   * Adds the given array of users to the list of viewers for the Spreadsheet.
   *
   * @param emailAddresses
   * @returns {Spreadsheet}
   */
  public addViewers(emailAddresses: string[]) : Spreadsheet {
    for(var email of emailAddresses) {
      this._viewers.push(new User(email));
    }
    return this;
  }

  /**
   * Add named ranges in this spreadsheet.
   */
  public addNamedRanges(name: string, range: Range) : NamedRange {
    const namedRange = new NamedRange(this, name, range);

    this._namedRanges.push(namedRange);

    return namedRange;
  }

  /**
   * Appends a row to the spreadsheet.
   *
   * @param columns
   * @returns {Sheet}
   */
  public appendRow(columns: Array<any>) : Sheet {
    return this._activeSheet.appendRow(columns);
  }

  /**
   * Deletes the column at the given column position.
   *
   * @param columnPosition
   * @returns {any}
   */
  public deleteColumn(columnPosition: number) {
    return this._activeSheet.deleteColumn(columnPosition);
  }

  /**
   * Deletes a number of columns starting at the given column position.
   *
   * @param columnPosition
   * @param howMany
   * @returns {Sheet}
   */
  public deleteColumns(columnPosition: number, howMany: number) {
    return this._activeSheet.deleteColumns(columnPosition, howMany);
  }

  /**
   * Deletes the row at the given column position.
   *
   * @param rowPosition
   * @returns {any}
   */
  public deleteRow(rowPosition: number) {
    return this._activeSheet.deleteRow(rowPosition);
  }

  /**
   * Deletes a number of rows starting at the given column position.
   *
   * @param rowPosition
   * @param howMany
   * @returns {Sheet}
   */
  public deleteRows(rowPosition: number, howMany: number) {
    return this._activeSheet.deleteRows(rowPosition, howMany);
  }

  /**
   * Deletes the specified sheet.
   *
   * @param sheet
   */
  public deleteSheet(sheet: Sheet) : void {
    const f = (s) => s.id !== sheet.id;
    this._sheets = this._sheets.filter(f);
  }

  /**
   * Duplicates the active sheet and makes it the active sheet.
   */
  public duplicateActiveSheet() : void {
    // TODO
  }

  /**
   * Returns the active cell in this sheet.
   *
   * @returns {null}
   */
  public getActiveCell() : Range {
    throw new Error(Errors.NOT_IMPLEMENTED_YET);
  }

  /**
   * Returns the active range for the active sheet.
   *
   * @returns {Range}
   */
  public getActiveRange() : Range {
    return this._activeSheet.getActiveRange();
  }

  /**
   * Gets the active sheet in a spreadsheet.
   */
  public getActiveSheet() : Sheet {
    return this._activeSheet;
  }

  /**
   * Returns a Range corresponding to the dimensions in which data is present.
   * This is functionally equivalent to creating a Range bounded by A1 and (Range.getLastColumn(),
   * Range.getLastRow()).
   *
   */
  public getDataRange() : Range {
    return this._activeSheet.getRange(1, 1, this.getLastRow() - 1, this.getLastColumn() - 1);
  }

  /**
   * Gets the list of editors for this Spreadsheet.
   *
   * @returns {User[]}
   */
  public getEditors(): User[] {
    return this._editors;
  }

  /**
   * Gets a unique identifier for this spreadsheet.
   *
   * @returns {number}
   */
  public getId() : number {
    return this._id;
  }

  /**
   * Returns the position of the last column that has content.
   */
  public getLastColumn() : number {
    return this._activeSheet.getLastColumn();
  }

  /**
   * Returns the position of the last row that has content.
   */
  public getLastRow() : number {
    return this._activeSheet.getLastRow();
  }

  /**
   * Gets the name of the document.
   *
   * @returns {string}
   */
  public getName() : string {
    return this._name;
  }

  /**
   * Gets all the named ranges in this spreadsheet.
   *
   * @returns {NamedRange[]}
   */
  public getNamedRanges() : NamedRange[] {
    return this._namedRanges;
  }

  /**
   * Returns the number of sheets in this spreadsheet.
   *
   * @returns {number}
   */
  public getNumSheets() : number {
    return this._sheets.length;
  }

  /**
   * Returns the range as specified in A1 notation or R1C1 notation.
   *
   * @param a1Notation
   * @returns {Range}
   */
  public getRange(a1Notation: string) : Range {
    return this._activeSheet.getRange(a1Notation);
  }

  /**
   * Returns a sheet with the given name.
   *
   * @param name
   */
  public getSheetByName(name: string) : Sheet | null {
    const sameName = (sheet: Sheet) => sheet.getName() === name;
    const results = this._sheets.filter(sameName);

    return results.length ? results[0] : null;
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
    return this._activeSheet.getSheetValues(startRow, startColumn, numRows, numColumns);
  }

  /**
   * Gets all the sheets in this spreadsheet.
   */
  public getSheets() : Sheet[] {
    return this._sheets;
  }

  /**
   * Gets the list of viewers and commenters for this Spreadsheet.
   *
   * @returns {User[]}
   */
  public getViewers(): User[] {
    return this._viewers;
  }

  /**
   * Inserts a column after the given column position.
   *
   * @param afterPosition
   */
  public insertColumnAfter(afterPosition: number) : Sheet {
    return this._activeSheet.insertColumnAfter(afterPosition);
  }

  /**
   * Inserts a column before the given column position.
   * @param beforePosition
   */
  public insertColumnBefore(beforePosition: number) : Sheet {
    return this._activeSheet.insertColumnBefore(beforePosition);
  }

  /**
   * Inserts a number of columns after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnsAfter(afterPosition: number, howMany: number) : Sheet {
    return this._activeSheet.insertColumnsAfter(afterPosition, howMany);
  }

  /**
   * Inserts a number of columns before the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnsBefore(beforePosition: number, howMany: number) : Sheet {
    return this._activeSheet.insertColumnBefore(beforePosition, howMany);
  }

  /**
   * Inserts a row after the given row position.
   *
   * @param afterPosition
   */
  public insertRowAfter(afterPosition: number) : Sheet {
    return this._activeSheet.insertRowAfter(afterPosition);
  }

  /**
   * Inserts a row before the given row position.
   *
   * @param beforePosition
   */
  public insertRowBefore(beforePosition: number) : Sheet {
    return this._activeSheet.insertRowBefore(beforePosition);
  }

  /**
   * Inserts a number of rows after the given row position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowsAfter(afterPosition: number, howMany: number) : Sheet {
    return this._activeSheet.insertRowsAfter(afterPosition, howMany);
  }

  /**
   * Inserts a number of rows before the given row position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowsBefore(beforePosition: number, howMany: number) : Sheet {
    return this._activeSheet.insertRowsBefore(beforePosition, howMany);
  }

  /**
   * Inserts a new sheet in the spreadsheet, with a default name. As a side effect, it makes it the active sheet.
   *
   * @returns {Sheet}
   */
  public insertSheet(sheetIndex?: number, options?: any) : Sheet {
    // sheet props
    const sheetIndexSuffix = this._sheets.length + 1;
    const sheetName = sheetConfig.name + sheetIndexSuffix;

    var cfg = Object.assign({}, sheetConfig, {
      id: new Date().getTime(),
      name: sheetName,
      numRows: this._numRows,
      numColumns: this._numColumns
    });

    // creation
    const sheet = new Sheet(cfg, this);
    this._sheets.push(sheet);
    this._activeSheet = sheet;
    return sheet;
  }

  /**
   * Remove named ranges in this spreadsheet.
   */
  public removeNamedRanges(name: NamedRange) : void {
    const isNotGoodId = (n) => n.id !== name.id;
    this._namedRanges = this._namedRanges.filter(isNotGoodId);
  }

  /**
   * Sets the active range for the active sheet.
   *
   * @param range
   */
  public setActiveRange(range) : Range {
    return this._activeSheet.setActiveRange(range);
  }

  /**
   * Sets the active sheet in a spreadsheet.
   *
   * @param sheet
   */
  public setActiveSheet(sheet) : Sheet {
    this._activeSheet = sheet;
    return this._activeSheet;
  }

}
