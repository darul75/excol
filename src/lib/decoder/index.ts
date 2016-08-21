// Interfaces
export interface EncodeFunc {
  (columnName: string): number;
}

export interface DecodeFunc {
  (columnNumber: number): string;
}

// Constants
const A_CHAR_CODE = 65;
const CHAR_COUNT = 26;

/**
 * Encodes a column name and transforms it into a column number.
 *
 * @param {string} columnName The column name
 *
 * @return {number} the column number
 *
 * @throws Error if the column name is null or empty
 */
export const encode: EncodeFunc = (columnName: string): number => {
  if (columnName == null || columnName.length == 0) {
    throw new Error('column name is empty');
  }

  const last = columnName.length - 1;
  let offset = last;
  let columnNumber = 1;
  let power = 1;

  while (offset >= 0) {
    let x = columnName.charCodeAt(offset) - A_CHAR_CODE;
    if (offset != last) {
      x += 1;
    }
    columnNumber += (power * x);
    power *= CHAR_COUNT;
    offset -= 1;
  }

  return columnNumber;
};

/**
 * Decodes a column number and transforms it into a column name.
 *
 * @param {number} columnNumber The column number
 *
 * @return {string} the column name
 *
 * @throws Error if the column number is negative or greater than 26^4
 */
export const decode:DecodeFunc = (columnNumber: number): string=> {

  if (columnNumber < 1 || columnNumber > Math.pow(CHAR_COUNT, 4)) {
    throw new Error(`column number out of range: ${columnNumber}, limit is ${Math.pow(CHAR_COUNT, 4)}`);
  }

  const columnName = ['', '', '', ''];
  let offset = columnName.length;
  let dividend = columnNumber;
  let modulo: number;

  while (dividend > 0) {
    modulo = (dividend - 1) % CHAR_COUNT;
    columnName[--offset] =  String.fromCharCode(A_CHAR_CODE + modulo);
    dividend = Math.round((dividend - modulo) / CHAR_COUNT);
  }

  return columnName.join('');
};
