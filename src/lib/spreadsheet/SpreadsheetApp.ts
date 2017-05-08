// Imports
import { Range, CellValue } from '../range/Range'
import { Sheet } from "../sheet/Sheet"
import { Spreadsheet } from './Spreadsheet'
import { Ui } from '../ui/UI'
import { DataValidationBuilder } from '../validation/DataValidationBuilder'
import {DataValidationCriteria} from '../validation/DataValidationCriteria'


// Constants
const DIMENSION = 1000;

export interface SpreadsheetAppConfig {

}

const spreadsheetConfig: SpreadsheetAppConfig = {
  numRows: DIMENSION,
  numColumns: DIMENSION
};

// Class & methods
/**
 * SpreadsheetApp class.
 */
export class SpreadsheetApp {

  private cfg: SpreadsheetAppConfig;
  private _active: Spreadsheet;

  public _numRows: number;
  public _numColumns: number;

  static DataValidationCriteria = DataValidationCriteria;

  /**
   * Constructor
   *
   * @param {object?} cfg Spreadsheet configuration
   */
  constructor(cfg: SpreadsheetAppConfig) {
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
  init({numRows = DIMENSION, numColumns = DIMENSION, cellValue} : { numRows?: number, numColumns?: number, cellValue?: CellValue }): void {
    this._numRows = numRows;
    this._numColumns = numColumns;
    this._active = this.create('Untitled spreadsheet', numRows, numColumns);
  }

  /**
   * Creates a new spreadsheet with the given name.
   *
   * @param name
   * @returns {Spreadsheet}
   */
  public create(name: string, rows: number = DIMENSION, columns: number = DIMENSION) : Spreadsheet {
    const cfg = Object.assign({}, spreadsheetConfig, {
      name: name,
      numRows: rows,
      numColumns: columns
    });
    const sSheet = new Spreadsheet(cfg);

    this._active = sSheet;

    return sSheet;
  }

  /**
   * Applies all pending Spreadsheet changes.
   */
  public flush() : void {
    // not applicable
  }

  /**
   * Returns the currently active spreadsheet
   */
  public getActive() : Spreadsheet {
    return this._active;
  }

  /**
   * Returns the range of cells that is currently considered active.
   */
  public getActiveRange() : Range {
    return this._active.getActiveRange();
  }

  /**
   * Gets the active sheet in a spreadsheet.
   */
  public getActiveSheet() : Sheet {
    return this._active.getActiveSheet();
  }

  /**
   * Returns the currently active spreadsheet, or null if there is none.
   *
   * @returns {Spreadsheet}
   */
  public getActiveSpreadsheet() : Spreadsheet {
    return this._active;
  }

  /**
   * Returns an instance of the spreadsheet's user-interface environment that allows the script to add features like menus, dialogs, and sidebars.
   *
   * @returns {Ui}
   */
  public getUi() : Ui {
    return new Ui();
  }

  /**
   * Creates a builder for a data-validation rule.
   */
  public newDataValidation() : DataValidationBuilder {
    return new DataValidationBuilder();
  }

  /**
   * Sets the active range for the active sheet.
   *
   * @param range
   */
  public setActiveRange(range) : Range {
    return this._active.setActiveRange(range);
  }

  /**
   * Sets the active sheet in a spreadsheet.
   *
   * @param sheet
   */
  public setActiveSheet(sheet) : Sheet {
    return this._active.setActiveSheet(sheet);
  }

  /**
   * Sets the active spreadsheet.
   *
   * @param newActiveSpreadsheet
   */
  public setActiveSpreadsheet(newActiveSpreadsheet) : void {
    this._active = newActiveSpreadsheet;
  }
}
