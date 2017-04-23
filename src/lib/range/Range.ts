// Import
import  { Cell, NewCell } from '../Cell'
import {Errors} from '../Error'
import Fields from '../Fields'
import { decode } from '../decoder/index'
import { isTwoDimArray, isTwoDimArrayCorrectDimensions, isTwoDimArrayOfString, isTwoDimArrayOfNumber } from '../../utils/mix'
import {Sheet} from '../sheet/Sheet';

// Interfaces

// Range cell value type
export interface CellValue { }

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
  private _cells: any[][];

  /**
   * Constructor
   *
   * @param {number} row Row coordinate
   * @param {number} rowHeight Number of rows
   * @param {number} row Column coordinate
   * @param {number} columnWidth Number of columns
   * @param {Array} cells Cell values
   */
  constructor(row: number = 0, rowHeight: number, column: number = 0, columnWidth: number, parent: Sheet | null, cells?: Array<Array<Cell>>) {
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
        const clone = bg == null || "object" != typeof bg ? bg : Object.assign({}, bg);
        bgs[r][c] = clone;
      }

    return bgs;
  }

  /**
   * Returns the background color of the top-left cell in the range.
   */
  public getCell(row: number, column: number) : any {
    return this._gridRange._cells[row - 1][column - 1];
  }

  /**
   * Returns the starting column position for this range.
   */
  public getColumn() : number {
    return this._column;
  }

  /**
   * Returns the data-validation rule for the top-left cell in the range. If data validation has not been set on the cell, this method returns null.
   *
   * // TODO
   */
  public getDataValidation() : any {
    return null;
  }

  /**
   * Returns the data-validation rule for the top-left cell in the range. If data validation has not been set on the cell, this method returns null.
   *
   * // TODO
   */
  public getDataValidations() : any {
    return null;
  }

  /**
   * Returns the displayed value of the top-left cell in the range. The value will be of type String.
   *
   * @returns display value
   */
  public getDisplayValue() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].value;
  }

  /**
   * Returns the displayed value of the top-left cell in the range. The value will be of type String.
   *
   * @returns display value
   */
  public getDisplayValues() : string[][] {
    const values = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].value;
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        colors[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        families[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        lines[r][c] = clone;
      }

    return lines;
  }

  /**
   * Returns the font size in point size of the cell in the top-left corner of the range.
   */
  public getFontSize() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontSize;
  }

  /**
   * Returns the font sizes of the cells in the range.
   *
   * @returns font sizes
   */
  public getFontSizes() : string[][] {
    const sizes = createArray(this._rowHeight, this._columnWidth);

    for (var r = 0; r < this._rowHeight; r++)
      for (var c = 0; c < this._columnWidth; c++) {
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontSize;
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        sizes[r][c] = clone;
      }

    return sizes;
  }

  /**
   * Returns the font style ('italic' or 'normal') of the cell in the top-left corner of the range.
   */
  public getFontStyle() : string {
    return this._gridRange._cells[this._row - 1][this._column - 1].fontSize;
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
        const value = this._gridRange._cells[this._row + r - 1][this._column + c - 1].fontSize;
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        sizes[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        weights[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        formulas[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        formulas[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        values[r][c] = clone;
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
        const clone = value == null || "object" != typeof value ? value : Object.assign({}, value);
        values[r][c] = clone;
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
        const clone = format == null || "object" != typeof format ? format : Object.assign({}, format);
        formats[r][c] = clone;
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
  public setBackgrounds(colors: Array<string>[]): void {
    this.applyStringPropertyValues(Fields.BACKGROUND, colors);
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
  public setFontColors(colors: Array<string>[]): void {
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
  public setFontFamilies(fontFamilies: Array<string>[]): void {
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
  public setFontLines(fontLines: Array<string>[]): void {
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
  public setFontStyle(fontStyle: number) : void {
    this.applyProperty(Fields.FONT_STYLE, fontStyle);
  }

  /**
   * Sets a rectangular lib of font styles (must match dimensions of this range).
   *
   * @param fontStyles
   */
  public setFontStyles(fontStyles: Array<string>[]): void {
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
  public setFontWeights(fontWeights: Array<string>[]): void {
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
  public setFormulas(formulas: Array<string>[]): void {
    this.applyStringPropertyValues(Fields.FORMULA, formulas);
  }

  /**
   * Sets a rectangular lib of formulas (must match dimensions of this range)
   *
   * @param formulas
   */
  public setFormulasR1C1(formulas: Array<string>[]): void {
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
  public setHorizontalAlignments(horizontalAlignments: Array<string>[]): void {
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
  public setNotes(notes: Array<string>[]): void {
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
  public setNumberFormats(numberFormats: Array<string>[]): void {
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
  public setVerticalAlignments(verticalAlignments: Array<string>[]): void {
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
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row - 1 + r][this._column - 1 + c].value = value;
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
      for (var c = 0; c < columnCount; c++)
        cells[this._row + r - 1][this._column + c - 1].value = values[r][c];
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
      throw new Error("can not merge single cell dude !");
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

  public moveTo(target: Range) : void {
    if (this.getMergedRanges().length > 0) {
      throw new Error(Errors.INCORRECT_MOVE_TO);
    }

    const sourceValues: Object[][] = this.values;

    const sourceFormats: Object[][] = this.getNumberFormats();

    for (var r = 0; r < this._rowHeight; r++) {
        const originRow: number = this._row + r - 1;
        const targetRow: number = target._row + r - 1;
        for (var c = 0; c < this._columnWidth; c++) {
          const originCol: number = this._column + c - 1;
          const targetCol: number = target._column + c - 1;
          // assign target
          this._gridRange._cells[targetRow][targetCol].numberFormat = sourceFormats[r][c];
          this._gridRange._cells[targetRow][targetCol].value = sourceValues[r][c];
          // clean source
          this._gridRange._cells[originRow][originCol].value = null;
          this._gridRange._cells[originRow][originCol].numberFormat = null;
        }
    }
  }

  public breakApart() : void {

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

  private applyProperty(propName: string, propValue : any) : void {
    const rowCount = this._rowHeight;
    const columnCount = this._columnWidth;

    for (var r = 0; r < rowCount; r++)
      for (var c = 0; c < columnCount; c++)
        this._gridRange._cells[this._row + r - 1][this._column + c - 1][propName] = propValue;
  }

  private applyStringPropertyValues(propName: string, propValues : Array<string>[]) : void {
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