// Imports
import { encode } from './../decoder/index'
import { A1Type, MatrixConverterFunc } from './Types'

/*

 Detail	                    A1	           Matrix

 A1	                        A1	           [ [ '1,1' ] ]
 Multiple A1                A1,B5          [ [ '1,1' ], [ '5,2' ] ]
 A1 through B5	            A1:B5	         [ [ '1,1', '5,2' ] ]
 A multiple-area selection	A1:B5,A1:B5	   [ [ '1,1', '5,2' ], [ '1,1', '5,2' ] ]
 Column A	                  A:A	           [ [ 'A', 'A' ] ]
 Row 1	                    1:1	           [ [ '1', '1' ] ]
 Rows 1 through 5	          1:5	           [ [ '1', '5' ] ]
 Columns A through C	      A:C	           [ [ 'A', 'C' ] ]
 Rows 1, 3, and 8	          1:1,3:3,8:8	   [ [ '1', '1' ], [ '3', '3' ], [ '8', '8' ] ]
 Columns A, C, and F	      A:A,C:C,F:F	   [ [ 'A', 'A' ], [ 'C', 'C' ], [ 'F', 'F' ] ]
 Multiple mixed             A1,1:1,A1:B5,C:F   [ [ '1,1' ], [ '1', '1' ], [ '1,1', '5,2' ], [ 'F', 'F' ] ]

*/



// Constants
const COMMA = ',';
const COLUMN = ':';
const COLUMN_REGEXP = /^[A-z]+/;
const EMPTY = '';
const SPACE_REGEXP = / /g;

/**
 * Encodes A1 notation to matrix notation.
 *
 *
 * @throws Error if the column name is null or empty
 */
export const toMatrix: MatrixConverterFunc = ({ A1 }: A1Type) : string[][] => {

  if (A1 == null || A1.length == 0) {
    throw new Error('A1 notation is empty');
  }

  // Remove spaces
  A1 = A1.replace(SPACE_REGEXP, EMPTY);

  // Build matrix representation from A1
  return A1.split(COMMA).map(elt => {
    return elt.split(COLUMN).map(rangeElt => {
      return convertA1Notation(rangeElt);
    });
  });

};

const convertA1Notation = (A1: string) => {
  const matches = A1.match(COLUMN_REGEXP);
  const columnName =  matches && matches.length ? matches[0] : '';
  const rowIndex = A1.replace(columnName, '');

  if (rowIndex === '') { return columnName; }

  if (columnName === '') { return rowIndex; }

  return rowIndex + COMMA + encode(columnName);
};
