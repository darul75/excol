import { CellValue } from './range/Range';
import { DataValidation } from './validation/DataValidation';
import { DataValidationCriteria } from './validation/DataValidationCriteria';

export interface Cell {
    row: number;
    column: number;
    background: string;
    dataValidation: DataValidation;
    numberFormat: string;
    fontColor: string;
    fontFamily: string;
    fontLine: string;
    fontSize: number;
    fontStyle: string;
    fontWeight: string;
    formula: string;
    horizontalAlignment: string,
    merged: boolean,
    note: string,
    value: any;
    verticalAlignment: string
}

export const DefaultFontLine: string = 'none';
export const DefaultFontSize: number = 10;
export const DefaultFontStyle: string = 'normal';
export const DefaultFontWeight: string = 'normal';
export const DefaultEmptyString: string = '';

export const NewCell = (row: number, column: number, value: any) : Cell => {
  return {
      row: row, column: column,
      background: DefaultEmptyString,
      dataValidation: new DataValidation(true, '', DataValidationCriteria.ANY, []),
      fontColor: DefaultEmptyString, fontFamily: DefaultEmptyString, fontStyle: DefaultFontStyle, fontLine: DefaultFontLine, fontSize: DefaultFontSize, fontWeight: DefaultFontWeight,
      formula: DefaultEmptyString,
      horizontalAlignment: 'left',
      merged: false,
      note: DefaultEmptyString,
      numberFormat: DefaultEmptyString,
      value: value,
      verticalAlignment: 'top'
  };
};

export const clearContent = (cell: Cell, cellValue: CellValue | undefined) : void => {
  cell.value = cellValue;
};

export const clearFormat = (cell: Cell) : void => {
  cell.fontColor = DefaultEmptyString;
  cell.fontFamily = DefaultEmptyString;
  cell.fontLine = DefaultFontLine;
  cell.fontSize = DefaultFontSize;
  cell.fontStyle = DefaultFontStyle;
  cell.fontWeight = DefaultFontWeight;
};

