// Import
import  { Cell, DefaultEmptyString, NewCell, clearFormat, clearContent } from '../cell/Cell'
import {Errors} from '../utils/Error'
import Fields from '../cell/Fields'
import { decode } from '../decoder/index'
import {
  isTwoDimArray, isTwoDimArrayCorrectDimensions, isTwoDimArrayOfString, isTwoDimArrayOfNumber,isTwoDimArrayOfDataValidation,
  isNullUndefined
} from '../utils/Mix'
import { Sheet } from '../sheet/Sheet'
import { DataValidation } from '../validation/DataValidation'
import { DataValidationCriteria, validators } from '../validation/DataValidationCriteria'

// Interfaces

// Range cell value type
export interface CellValue { }

// Range clear option
export interface ClearOption {
  commentsOnly?: boolean
  contentsOnly?: boolean
  formatOnly?: boolean
  validationsOnly?: boolean
}

// Class & methods

/**
 * Range class.
 */
export class Range {

  // Coordinates
  private _row: number;
  private _rowHeight: number;
  private _column: number;
  private _columnWidth: number;

  // Properties
  private _parent: Sheet;
  private _gridRange: Range;
  private _mergedRanges: Range[];
  private _cells: Cell[][];
  private _cellValue: CellValue | undefined;

  /**
   * Constructor
   *
   * @param {number} row Row coordinate
   * @param {number} rowHeight Number of rows
   * @param {number} row Column coordinate
   * @param {number} columnWidth Number of columns
   * @param {Array} cells Cell values
   */
  constructor(row: number = 0, rowHeight: number, column: number = 0, columnWidth: number,
              parent: Sheet | null,
              cells?: Array<Array<Cell>>) {
    this._row = row;
    this._rowHeight = rowHeight;
    this._column = column;
    this._columnWidth = columnWidth;

    if (cells) {
      this._cells = cells;
    }

    if (parent) {
      this._parent = parent;
      this._gridRange = parent.range;
      this._mergedRanges = parent.mergedRanges;
    }
  }

  /**
   * Break any multi-column cells in the range into individual cells again.
   */
  public breakApart() : Range {

    const mergeIndexesToBeRemoved = new Array<number>();

    for (var i=0; i<this._mergedRanges.length; i++) {
      const range: Range = this._mergedRanges[i];

      const isOverlapping = this.overlap(range);

      if (isOverlapping) {

        let error: boolean = false;

        // check contained
        const isContained = this.contains(range);

        if (!isContained) throw new Error(Errors.INCORRECT_MERGE);

        // remove contained range
        if (!error && isContained) {
          mergeIndexesToBeRemoved.push(i);
        }

      }
    }

    // Removed englobed merge
    for (let i=0; i<mergeIndexesToBeRemoved.length; i++) {
      this._mergedRanges.splice(mergeIndexesToBeRemoved[i]-1, 1);
    }

    this._gridRange._cells[this._row - 1][this._column -1].merged = false;

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        if (!(r == 0 && c == 0)) {
          const cell = this._gridRange._cells[this._row + r - 1][this._column + c - 1];
          cell.merged = false;
          cell.value = null;
        }
      }

    return this;
  }

  /**
   * Clears the range of contents, formats, and data-validation rules.
   *
   * Clears the range of contents, format, data-validation rules, and/or comments, as specified with the given advanced options.
   * By default all data will be cleared.
   *
   * @param options
   * @returns {Range}
   */
  public clear(options?: ClearOption) : Range {
    if (!options) {
      this.breakApart();
    }

    this.clearAll(options);

    return this;
  }

  /**
   * Clear all cells attributes.
   */
  private clearAll(options?: ClearOption) : void {
    const hasCellValue = !isNullUndefined(this._cellValue);

    for (var rx = this._row - 1; rx < this._rowHeight; rx++)
      for (var cx = this._column - 1; cx < this._columnWidth; cx++) {
        // default init value
        const initialVal = hasCellValue ? this._cellValue : {value: rx + '-' + cx};
        if (isNullUndefined(options)) {
          this._gridRange._cells[rx][cx] = NewCell(rx, cx, initialVal);
        } else {
          if (options && options.formatOnly) {
            clearFormat(this._gridRange._cells[rx][cx]);
          }
          if (options && options.contentsOnly) {
            clearContent(this._gridRange._cells[rx][cx], this._cellValue);
          }
        }
      }
  }

  /**
   * Returns a string description of the range, in A1 notation.
   *
   * @returns {string}
   */
  public getA1Notation() : string {
    // one cell range
    const firstCell = decode(this._column) + this._row;
    if (this._rowHeight == 1 && this._columnWidth == 1) {
      return firstCell;
    }
    const secondCell = decode(this._column + this._columnWidth - 1) + (this._row + this._rowHeight - 1);
    // multiple rows or column cells range
    return firstCell + ':' + secondCell;
  }

  /**
   * Returns the background color of the top-left cell in the range.
   */
  public getBackground() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].background;
  }

  /**
   * Returns the background colors of the cells in the range.
   */
  public getBackgrounds() : string[][] {
    const bgs = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const bg = this._gridRange._cells[this._row + r - 1][this._column + c - 1].background;
        bgs[r][c] = bg;
      }

    return bgs;
  }

  /**
   * Returns a given cell within a range.
   */
  public getCell(row: number, column: number) : Range {
    const lastRow: number = this._row + this._rowHeight - 1;
    const lastColumn: number = this._column + this._columnWidth - 1;

    const rowTmp = this._row + row - 1;
    const colTmp = this._column + column - 1;

    if (rowTmp > lastRow || colTmp > lastColumn) throw new Error(Errors.INCORRECT_RANGE_CELL);

    return this._parent.getRange({row: rowTmp, column: colTmp, numRows: 1, numColumns: 1});
  }

  /**
   * Returns the starting column position for this range.
   */
  public getColumn() : number {
    return this._column;
  }

  /**
   * Returns the data-validation rule for the top-left cell in the range.
   * If data validation has not been set on the cell, this method returns null.
   */
  public getDataValidation() : DataValidation {
    return this._gridRange._cells[this._row - 1][this._column - 1].dataValidation;
  }

  /**
   * Returns the data-validation rule for the top-left cell in the range.
   * If data validation has not been set on the cell, this method returns null.
   * TODO: return null...
   */
  public getDataValidations() : DataValidation[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const validation = this._gridRange._cells[this._row + r - 1][this._column + c - 1].dataValidation;
        values[r][c] = validation;
      }

    return values;
  }

  /**
   * Returns the displayed value of the top-left cell in the range. The value will be of type String.
   *
   * @returns display value
   */
  public getDisplayValue() : string {
    const cell: Cell = this._gridRange._cells[this._row - 1][this._column - 1];
    return cell.value == null || "object" != typeof cell.value ? ''+ cell.value : cell.value.toString();
  }

  /**
   * Returns the rectangular grid of values for this range.
   *
   * @returns a two-dimensional array of values
   */
  public getDisplayValues() : string[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].value;
        const clone = value == null || "object" != typeof value ? ''+ value : value.toString();
        values[r][c] = clone;
      }

    return values;
  }

  /**
   * Returns the font color of the cell in the top-left corner of the range.
   */
  public getFontColor() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontColor;
  }

  /**
   * Returns the font colors of the cells in the range
   *
   * @returns font colors
   */
  public getFontColors() : string[][] {
    const colors = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontColor;
        colors[r][c] = value;
      }

    return colors;
  }

  /**
   * Returns the font families of the cells in the range.
   *
   * @returns font families
   */
  public getFontFamilies() : string[][] {
    const families = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontFamily;
        families[r][c] = value;
      }

    return families;
  }

  /**
   * Returns the font family of the cell in the top-left corner of the range.
   */
  public getFontFamily() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontFamily;
  }

  /**
   * Gets the line style of the cell in the top-left corner of the range.
   */
  public getFontLine() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontLine;
  }

  /**
   * Gets the line style of the cells in the range.
   *
   * @returns font lines
   */
  public getFontLines() : string[][] {
    const lines = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontLine;
        lines[r][c] = value;
      }

    return lines;
  }

  /**
   * Returns the font size in point size of the cell in the top-left corner of the range.
   */
  public getFontSize() : number {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontSize;
  }

  /**
   * Returns the font sizes of the cells in the range.
   *
   * @returns font sizes
   */
  public getFontSizes() : number[][] {
    const sizes = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontSize;
        sizes[r][c] = value;
      }

    return sizes;
  }

  /**
   * Returns the font style ('italic' or 'normal') of the cell in the top-left corner of the range.
   */
  public getFontStyle() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontStyle;
  }

  /**
   * Returns the font styles of the cells in the range.
   *
   * @returns font styles
   */
  public getFontStyles() : string[][] {
    const sizes = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontStyle;
        sizes[r][c] = value;
      }

    return sizes;
  }

  /**
   * Returns the font weight (normal/bold) of the cell in the top-left corner of the range.
   */
  public getFontWeight() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontWeight;
  }

  /**
   * Returns the font weights of the cells in the range.
   *
   * @returns font weights
   */
  public getFontWeights() : string[][] {
    const weights = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontWeight;
        weights[r][c] = value;
      }

    return weights;
  }

  /**
   * Returns the formula (A1 notation) for the top-left cell of the range, or an empty string if the cell is empty or doesn't contain a formula.
   */
  public getFormula() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].formula;
  }

  /**
   * Returns the formula (R1C1 notation) for a given cell, or null if none.
   */
  public getFormulaR1C1() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].formula;
  }

  /**
   * Returns the formulas (A1 notation) for the cells in the range. Entries in the 2D array will be an empty string for cells with no formula.
   *
   * @returns formulas
   */
  public getFormulas() : string[][] {
    const formulas = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].formula;
        formulas[r][c] = value;
      }

    return formulas;
  }

  /**
   * Returns the formulas (R1C1 notation) for the cells in the range.
   *
   * @returns formulas
   */
  public getFormulasR1C1() : string[][] {
    const formulas = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].formula;
        formulas[r][c] = value;
      }

    return formulas;
  }

  /**
   * Returns the height of the range.
   */
  public getHeight() : number {
    return this._rowHeight;
  }

  /**
   * Returns the horizontal alignment of the text (left/center/right) of the cell in the top-left corner of the range.
   */
  public getHorizontalAlignment() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].horizontalAlignment;
  }

  /**
   * Returns the horizontal alignments of the cells in the range.
   *
   * @returns font weights
   */
  public getHorizontalAlignments() : string[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].horizontalAlignment;
        values[r][c] = value;
      }

    return values;
  }

  /**
   * Returns the end column position.
   */
  public getLastColumn() : number {
    return this._column + this._columnWidth;
  }

  /**
   * Returns the end row position.
   */
  public getLastRow() : number {
    return this._row + this._rowHeight;
  }

  /**
   * Returns an array of Range objects representing merged cells that either are fully within the current range,
   * or contain at least one cell in the current range."
   *
   * @returns {Range[]}
   */
  public getMergedRanges() : Range[] {
    const results: Range[] = [];

    for (var i=0; i<this._mergedRanges.length; i++) {
      const range: Range = this._mergedRanges[i];

      if (this.overlap(range)) {
        results.push(range);
      }
    }
    return results;
  }

  /**
   * Returns the note associated with the given range.
   */
  public getNote() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].note;
  }

  /**
   * Returns the notes associated with the cells in the range.
   *
   * @returns notes
   */
  public getNotes() : string[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].note;
        values[r][c] = value;
      }

    return values;
  }

  /**
   * Returns the number of columns in this range.
   *
   * @returns {number}
   */
  public getNumColumns() : number {
    return this.columnWidth;
  }

  /**
   * Returns the number of rows in this range.
   *
   * @returns {number}
   */
  public getNumRows() : number {
    return this.rowHeight;
  }

  /**
   * Get the number or date formatting of the top-left cell of the given range.
   */
  public getNumberFormat() : string {
    const format = this._gridRange._cells[this._row - 1][this._column - 1].numberFormat;
    return format;
  }

  /**
   * Returns the number or date formats for the cells in the range.
   */
  public getNumberFormats() : string[][] {
    const formats = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const format = this._gridRange._cells[this._row + r - 1][this._column + c - 1].numberFormat;
        formats[r][c] = format;
      }

    return formats;
  }

  /**
   * Returns the row position for this range.
   *
   * @returns {number}
   */
  public getRow() : number {
    return this._row;
  }

  /**
   * Returns the row position for this range.
   *
   * @returns {number}
   */
  public getRowIndex() : number {
    return this._row;
  }

  /**
   * Returns the vertical alignment (top/middle/bottom) of the cell in the top-left corner of the range.
   */
  public getVerticalAlignment() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].verticalAlignment;
  }

  /**
   * Returns the vertical alignments of the cells in the range.
   *
   * @returns alignments
   */
  public getVerticalAlignments() : string[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].verticalAlignment;
        values[r][c] = value;
      }

    return values;
  }

  /**
   * Returns the width of the range in columns.
   *
   * @returns the number of columns in the range
   */
  public getWidth() : number {
    return this.columnWidth;
  }

  /**
   * Sets the background color of all cells in the range in CSS notation (like '#ffffff' or 'white').
   *
   * @param color
   */
  public setBackground(color: string) : void {
    this.applyProperty(Fields.BACKGROUND, color);
  }

  /**
   * Sets a rectangular lib of background colors (must match dimensions of this range).
   * The colors are in CSS notation (like '#ffffff' or 'white').
   *
   * @param colors
   */
  public setBackgrounds(colors: string[][]): void {
    this.applyStringPropertyValues(Fields.BACKGROUND, colors);
  }

  /**
   * Sets one data-validation rule for all cells in the range.
   *
   * @param rule
   * @returns {Range}
   */
  public setDataValidation(rule: DataValidation) : Range {
    this.applyProperty(Fields.DATA_VALIDATION, rule);

    return this;
  }

  /**
   * Sets the data-validation rules for all cells in the range.
   *
   * @param rules
   * @returns {Range}
   */
  public setDataValidations(rules	: DataValidation[][]) : Range {
    this.applyDataValidationPropertyValues(Fields.DATA_VALIDATION, rules);

    return this;
  }

  /**
   * Sets the font color in CSS notation (like '#ffffff' or 'white').
   *
   * @param color
   */
  public setFontColor(color: string) : void {
    this.applyProperty(Fields.FONT_COLOR, color);
  }

  /**
   * Sets a rectangular lib of font colors (must match dimensions of this range).
   * The colors are in CSS notation (like '#ffffff' or 'white').
   *
   * @param colors
   */
  public setFontColors(colors: string[][]): void {
    this.applyStringPropertyValues(Fields.FONT_COLOR, colors);
  }

  /**
   * Sets the font family, such as "Arial" or "Helvetica".
   *
   * @param fontFamily
   */
  public setFontFamily(fontFamily: string) : void {
    this.applyProperty(Fields.FONT_FAMILY, fontFamily);
  }

  /**
   * Sets a rectangular lib of font families (must match dimensions of this range).
   * Examples of font families are "Arial" or "Helvetica".
   *
   * @param fontFamilies
   */
  public setFontFamilies(fontFamilies: string[][]): void {
    this.applyStringPropertyValues(Fields.FONT_FAMILY, fontFamilies);
  }

  /**
   * Sets the font line style of the given range ('underline', 'line-through', or 'none').
   *
   * @param fontLine
   */
  public setFontLine(fontLine: string) : void {
    this.applyProperty(Fields.FONT_LINE, fontLine);
  }

  /**
   * Sets a rectangular lib of line styles (must match dimensions of this range).
   *
   * @param fontLines
   */
  public setFontLines(fontLines: string[][]): void {
    this.applyStringPropertyValues(Fields.FONT_LINE, fontLines);
  }

  /**
   * Sets the font size, with the size being the point size to use.
   *
   * @param fontSize
   */
  public setFontSize(fontSize: number) : void {
    this.applyProperty(Fields.FONT_SIZE, fontSize);
  }

  /**
   * Sets a rectangular lib of font sizes (must match dimensions of this range).
   * The sizes are in points.
   *
   * @param fontSizes
   */
  public setFontSizes(fontSizes: Array<number>[]): void {
    this.applyNumberPropertyValues(Fields.FONT_SIZE, fontSizes);
  }

  /**
   * Set the font style for the given range ('italic' or 'normal').
   *
   * @param fontStyle
   */
  public setFontStyle(fontStyle: string) : void {
    this.applyProperty(Fields.FONT_STYLE, fontStyle);
  }

  /**
   * Sets a rectangular lib of font styles (must match dimensions of this range).
   *
   * @param fontStyles
   */
  public setFontStyles(fontStyles: string[][]): void {
    this.applyStringPropertyValues(Fields.FONT_STYLE, fontStyles);
  }

  /**
   * Set the font weight for the given range (normal/bold).
   *
   * @param fontWeight
   */
  public setFontWeight(fontWeight: string) : void {
    this.applyProperty(Fields.FONT_WEIGHT, fontWeight);
  }

  /**
   * Sets a rectangular lib of font weights (must match dimensions of this range).
   * An example of a font weight is "bold".
   *
   * @param fontWeights
   */
  public setFontWeights(fontWeights: string[][]): void {
    this.applyStringPropertyValues(Fields.FONT_WEIGHT, fontWeights);
  }

  /**
   * Updates the formula for this range.
   *
   * @param formula
   */
  public setFormula(formula: string) : void {
    this.applyProperty(Fields.FORMULA, formula);
  }

  /**
   * Updates the formula for this range.
   *
   * @param formula
   */
  public setFormulaR1C1(formula: string) : void {
    this.setFormula(formula);
  }

  /**
   * Sets a rectangular lib of formulas (must match dimensions of this range)
   *
   * @param formulas
   */
  public setFormulas(formulas: string[][]): void {
    this.applyStringPropertyValues(Fields.FORMULA, formulas);
  }

  /**
   * Sets a rectangular lib of formulas (must match dimensions of this range)
   *
   * @param formulas
   */
  public setFormulasR1C1(formulas: string[][]): void {
    this.setFormulas(formulas);
  }

  /**
   * Set the horizontal (left to right) alignment for the given range (left/center/right).
   *
   * @param horizontalAlignment
   */
  public setHorizontalAlignment(horizontalAlignment: string) : void {
    this.applyProperty(Fields.HORIZONTAL_ALIGNMENT, horizontalAlignment);
  }

  /**
   * Sets a rectangular lib of horizontal alignments.
   *
   * @param horizontalAlignments
   */
  public setHorizontalAlignments(horizontalAlignments: string[][]): void {
    this.applyStringPropertyValues(Fields.HORIZONTAL_ALIGNMENT, horizontalAlignments);
  }

  /**
   * Sets the note to the given value.
   *
   * @param note
   */
  public setNote(note: string) : void {
    this.applyProperty(Fields.NOTE, note);
  }

  /**
   * Sets a rectangular lib of notes (must match dimensions of this range).
   *
   * @param notes
   */
  public setNotes(notes: string[][]): void {
    this.applyStringPropertyValues(Fields.NOTE, notes);
  }

  /**
   * Sets the number or date format to the given formatting string.
   *
   * @param numberFormat
   */
  public setNumberFormat(numberFormat: string) : void {
    this.applyProperty(Fields.NUMBER_FORMAT, numberFormat);
  }

  /**
   * Sets a rectangular lib of number or date formats (must match dimensions of this range).
   *
   * @param numberFormats
   */
  public setNumberFormats(numberFormats: string[][]): void {
    this.applyStringPropertyValues(Fields.NUMBER_FORMAT, numberFormats);
  }

  /**
   * Set the vertical (left to right) alignment for the given range (left/center/right).
   *
   * @param verticalAlignment
   */
  public setVerticalAlignment(verticalAlignment: string) : void {
    this.applyProperty(Fields.VERTICAL_ALIGNMENT, verticalAlignment);
  }

  /**
   * Sets a rectangular lib of vertical alignments.
   *
   * @param horizontalAlignments
   */
  public setVerticalAlignments(verticalAlignments: string[][]): void {
    this.applyStringPropertyValues(Fields.VERTICAL_ALIGNMENT, verticalAlignments);
  }

  /**
   * Sets the value of the range. The value can be numeric, string, boolean or date.
   * If it begins with '=' it is interpreted as a formula.
   *
   * TODO: formula
   *
   * @param value
   */
  public setValue(value: CellValue): void {
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    this._parent.setLastRow(this._row + rowCount);
    this._parent.setLastColumn(this._column + columnCount);

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++) {
        // current cell
        const cell: Cell = this._gridRange._cells[this._row - 1 + r][this._column - 1 + c];
        // data valication
        var criteria: DataValidationCriteria = cell.dataValidation.getCriteriaType();
        if (criteria !== DataValidationCriteria.ANY) {
          validators[criteria](this._row + r, this._column + c, value, cell.dataValidation);
        }
        // assign value
        cell.value = value;
      }

  }

  /**
   * Sets a rectangular lib of values (must match dimensions of this range).
   *
   * @param values
   */
  public setValues(values: Array<CellValue>[]): void {
    checkValues(values, this);
    const rowCount = values.length;
    const columnCount = values[0].length;

    const cells = this._gridRange ? this._gridRange._cells : this._cells;

    this._parent.setLastRow(this._row + rowCount);
    this._parent.setLastColumn(this._column + columnCount);

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++) {
        // current cell
        const cell: Cell = cells[this._row + r - 1][this._column + c - 1];
        // data valication
        var criteria: DataValidationCriteria = cell.dataValidation.getCriteriaType();
        if (criteria !== DataValidationCriteria.ANY) {
          validators[criteria](this._row + r, this._column + c, values[r][c], cell.dataValidation);
        }
        // assign value
        cell.value = values[r][c];
      }
  }

  /**
   * Inserts a column after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnAfter(afterPosition: number, howMany?: number) : void {
    const cells = this._cells;
    const rowCount = cells.length;
    const toAddNum: number = howMany ? howMany : 1;

    // + lastRow
    const lastColumn: number = this._parent.getLastColumn();

    for (var r = 0; r < rowCount; r++) {
      const extraColumn: Cell[]= [];
      for(var i = 0; i < toAddNum; i++) extraColumn.push(NewCell(r, i + afterPosition, null));

      cells[r].splice(afterPosition, 0, ...extraColumn);
    }

    this._parent.setLastColumn(lastColumn + 1);
  }

  /**
   * Inserts a number of columns after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertColumnBefore(afterPosition: number, howMany?: number) : void {
    const cells = this._cells;
    const rowCount = cells.length;
    const toAddNum: number = howMany ? howMany : 1;

    // + lastRow
    const lastColumn: number = this._parent.getLastColumn();

    for (var r = 0; r < rowCount; r++) {
      const extraColumn: Cell[]= [];
      for(var i = 0; i < toAddNum; i++) extraColumn.push(NewCell(r, i + afterPosition, null));

      cells[r].splice(afterPosition - 1, 0, ...extraColumn);
    }

    this._parent.setLastColumn(lastColumn + 1);
  }

  /**
   * Inserts a row after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowAfter(afterPosition: number, howMany?: number) : void {
    const cells = this._cells;
    const columnCount = cells[0].length;
    const toAddNum: number = howMany ? howMany : 1;

    const lastRow: number = this._parent.getLastRow();

    for (var i = 0; i < toAddNum; i++) {
      const row: Cell[] = [];
      for (var col=0;col < columnCount; col++) {
        row.push(NewCell(afterPosition, col, null));
      }

      cells.splice(afterPosition, 0, row);
    }

    this._parent.setLastRow(lastRow + toAddNum);
  }

  /**
   * Inserts a number of rows after the given column position.
   *
   * @param afterPosition
   * @param howMany
   */
  public insertRowBefore(afterPosition: number, howMany?: number) : void {
    const cells = this._cells;
    const columnCount = cells[0].length;
    const toAddNum: number = howMany ? howMany : 1;

    const lastRow: number = this._parent.getLastRow();

    for (var i = 0; i < toAddNum; i++) {
      const row: Cell[] = [];
      for (var col=0;col < columnCount; col++) {
        row.push(NewCell(afterPosition, col, null));
      }

      cells.splice(afterPosition-1, 0, row);
    }

    this._parent.setLastRow(lastRow + toAddNum);
  }

  public get values() : any[][]
  {
    const values = createArray(this._rowHeight, this._columnWidth);

    const cells = this._gridRange ? this._gridRange._cells : this._cells;

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = cells[this._row + r - 1][this._column + c - 1].value;
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        values[r][c] = clone;
      }

    return values;
  }

  /**
   * Merges the cells in the range together into a single block.
   */
  public merge() : void {

    if (this._rowHeight == 1 && this._columnWidth == 1) {
      throw new Error(Errors.INCORRECT_MERGE_SINGLE_CELL);
    }

    const mergeIndexesToBeRemoved = new Array<number>();

    for (var i=0; i<this._mergedRanges.length; i++) {
      const range: Range = this._mergedRanges[i];

      const isOverlapping = this.overlap(range);

      if (isOverlapping) {
        let error: boolean = false;

        // check contained
        const isContained = this.contains(range);

        if (!isContained) throw new Error(Errors.INCORRECT_MERGE);

        // remove contained range
        if (!error && isContained) {
          mergeIndexesToBeRemoved.push(i);
        }
      }
    }

    // Removed englobed merge
    for (let i=0; i<mergeIndexesToBeRemoved.length; i++) {
      this._mergedRanges.splice(mergeIndexesToBeRemoved[i]-1, 1);
    }

    this._gridRange._cells[this._row - 1][this._column -1].merged = true;

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        if (!(r == 0 && c == 0)) {
          const cell = this._gridRange._cells[this._row - 1 + r][this._column -1 + c];
          cell.merged = true;
          cell.value = null;
        }
      }

    this._mergedRanges.push(this);
  }

  /**
   * Merge the cells in the range across the columns of the range.
   *
   * @param vertically
   */
  public mergeAcross(vertically?: boolean) : void {

    if (this._rowHeight == 1 && this._columnWidth == 1) {
      throw new Error(Errors.INCORRECT_MERGE_SINGLE_CELL);
    }

    const mergeIndexesToBeRemoved = new Array<number>();

    // check fhis range contains unmergeable ranges

    for (var i=0; i<this._mergedRanges.length; i++) {
      const range: Range = this._mergedRanges[i];

      const isOverlapping = this.overlap(range);

      if (isOverlapping) {
        let error: boolean = false;

        // check contained
        const isContained = this.contains(range);

        if (!isContained) throw new Error(Errors.INCORRECT_MERGE);

        // merged range not on one row only
        const isOnMultipleRowsOrColumns = !vertically ? range.rowHeight > 1 : range.columnWidth > 1;

        if (isOnMultipleRowsOrColumns) {
          if (!vertically) {
            throw new Error(Errors.INCORRECT_MERGE_HORIZONTAL);
          }
          throw new Error(Errors.INCORRECT_MERGE_VERTICALLY);
        }

        // remove contained range
        if (!error && !isOnMultipleRowsOrColumns) {
          mergeIndexesToBeRemoved.push(i);
        }
      }
    }

    // Removed englobed merge
    for (let i=0; i<mergeIndexesToBeRemoved.length; i++) {
      this._mergedRanges.splice(mergeIndexesToBeRemoved[i]-1, 1);
    }

    // otherwise merge line by line
    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const firstValueCell = vertically ? r == 0 : c == 0;
        if (!firstValueCell) {
          const cell = this._gridRange._cells[this._row - 1 + r][this._column -1 + c];
          cell.merged = true;
          cell.value = null;
        }
      }

    this._mergedRanges.push(this);
  }

  /**
   * Merges the cells in the range together.
   */
  public mergeVertically() : void {
    this.mergeAcross(true);
  }

  public offset(rowOffset: number, columnOffset: number): void {
    // todo limits

    this._row += rowOffset;
    this._column += columnOffset;
  }

  /**
   * Cut and paste (both format and values) from this range to the target range.
   *
   * @param target
   */
  public moveTo(target: Range) : void {
    if (this.getMergedRanges().length > 0) {
      throw new Error(Errors.INCORRECT_MOVE_TO);
    }

    // Source range attributes
    const sourceValues: Object[][] = this.values;
    const sourceFormats: string[][] = this.getNumberFormats();

    for (var r = 0; r < this._rowHeight; r++) {
        const originRow: number = this._row + r - 1;
        const targetRow: number = target._row + r - 1;
        for (var c = 0; c < this._columnWidth; c++) {
          const originCol: number = this._column + c - 1;
          const targetCol: number = target._column + c - 1;
          // assign target
          const targetCell: Cell = this._gridRange._cells[targetRow][targetCol];
          targetCell.numberFormat = sourceFormats[r][c];
          targetCell.value = sourceValues[r][c];

          // clean source
          this._gridRange._cells[originRow][originCol].value = null;
          this._gridRange._cells[originRow][originCol].numberFormat = DefaultEmptyString;

        }
    }
  }

  overlap(other: Range) {
    const right = this._column + this._columnWidth - 1;
    const otherRight = other._column + other._columnWidth - 1;
    const bottom = this._row + this._rowHeight - 1;
    const otherBottom = other._row + other._rowHeight - 1;

    return !(other._column > right ||
    otherRight < this._column ||
    other._row > bottom ||
    otherBottom < this._row);
  }

  contains(other: Range) {
    const right = this._column + this._columnWidth - 1;
    const otherRight = other._column + other._columnWidth - 1;
    const bottom = this._row + this._rowHeight - 1;
    const otherBottom = other._row + other._rowHeight - 1;

    return (
      this._column <= other._column && // left
      right >= otherRight && // right
      this._row <= other._row && // top
      bottom >= otherBottom // bottom
    );
  }

  public get rowHeight() {
    return this._rowHeight;
  }

  public get cells() {
    return this._cells;
  }

  public set cells(value) {
    this._cells = value;
  }

  public get columnWidth() {
    return this._columnWidth;
  }

  public set cellValue(value: CellValue | undefined) {
    if (!isNullUndefined(value))
      this._cellValue = value;
  }

  private applyProperty(propName: string, propValue : any) : void {
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row + r - 1][this._column + c - 1][propName] = propValue;
  }

  private applyDataValidationPropertyValues(propName: string, propValues : DataValidation[][]) : void {
    check2DimArrayOfDataValidation(propValues, this);
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row + r - 1][this._column + c - 1][propName] = propValues[r][c];
  }

  private applyStringPropertyValues(propName: string, propValues : string[][]) : void {
    check2DimArrayOfStrings(propValues, this);
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row + r - 1][this._column + c - 1][propName] = propValues[r][c];
  }

  private applyNumberPropertyValues(propName: string, propValues : Array<number>[]) : void {
    check2DimArrayOfNumbers(propValues, this);
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row + r - 1][this._column + c - 1][propName] = propValues[r][c];
  }
}

/**
 * Verify 2 dimensional array values.
 *
 * @param obj values
 * @param range current range
 */
const checkValues = (obj: any, range: Range) : void => {
  isTwoDimArray(obj);
  isTwoDimArrayCorrectDimensions(obj, range.rowHeight, range.columnWidth);
};

const check2DimArrayOfStrings = (obj: any, range: Range) : void => {
  isTwoDimArray(obj);
  isTwoDimArrayOfString(obj, range.rowHeight, range.columnWidth);
};

const check2DimArrayOfDataValidation = (obj: any, range: Range) : void => {
  isTwoDimArray(obj);
  isTwoDimArrayOfDataValidation(obj, range.rowHeight, range.columnWidth);
};

const check2DimArrayOfNumbers = (obj: any, range: Range) : void => {
  isTwoDimArray(obj);
  isTwoDimArrayOfNumber(obj, range.rowHeight, range.columnWidth);
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