// Imports
import { encode } from './../decoder/index'
import { toMatrix } from './ToMatrix'
import { A1Type, CoordinatesConverterFunc } from './Types'

/*
 Detail	                    A1	           Matrix

 A1	                        A1	           [ '1,1' ]
 Multiple A1                A1,B5          [ [ '1,1' ], [ '5,2' ] ]
 A1 through B5	            A1:B5	         [ [ '1,1', '5,2' ] ]
 A multiple-area selection	A1:B5,A1:B5	   [ [ '1,1', '5,2' ], [ '1,1', '5,2' ] ]
 Column A	                  A:A	           [ [ '-1', '1' ] ]
 Row 1	                    1:1	           [ [ '1', '1' ] ]
 Rows 1 through 5	          1:5	           [ [ '1', '5' ] ]
 Columns A through C	      A:C	           [ [ 'A', 'C' ] ]
 Rows 1, 3, and 8	          1:1,3:3,8:8	   [ [ '1', '1' ], [ '3', '3' ], [ '8', '8' ] ]
 Columns A, C, and F	      A:A,C:C,F:F	   [ [ '-1', '1' ], [ '-1', '3' ], [ '-1', '6' ] ]
 Multiple mixed             A1,1:1,A1:B5,C:F   [ [ '1,1' ], [ '1', '1' ], [ '1,1', '5,2' ], [ 'F', 'F' ] ]
 */

// Constants
const COMMA = ',';

// Small validation
export const A1_VALIDATOR_REGEXP = /^([A-Z]*[0-9]*,?:?)*$/;

export const validateA1 = ({ A1 }: A1Type) => A1_VALIDATOR_REGEXP.test(A1);

export const toCoordinates: CoordinatesConverterFunc = ({ A1 }: A1Type) : any[] => {

  if (A1 == null || A1.length === 0) {
    throw new Error('A1 notation is empty');
  }

  const matrixView: string[][] = toMatrix({A1: A1});

  return matrixView.map((elt: string[]) => {
    const firstElt = elt[0];

    // Single cell
    if (elt.length === 1) {
      return firstElt.split(COMMA).map(toInt);
    }

    // Range
    if (elt.length === 2) {
      const secondElt = elt[1];

      if (~firstElt.indexOf(COMMA) || ~secondElt.indexOf(COMMA)) {

        const startRangeRowColumns = firstElt.split(COMMA);
        const endRangeRowColumns = secondElt.split(COMMA);

        let firstCoordinate: Array<number | number[]> = [];
        let secondCoordinate: Array<number | number[]> = [];

        // check whole row or column [ 'A' ] [ '1' ]
        if (startRangeRowColumns.length === 1) {

          const item = startRangeRowColumns[0];
          if (isNan(item)) {
            firstCoordinate.push(-1, toInt(item));
          }
          else {
            firstCoordinate.push(toInt(item), -1);
          }
        }
        // normal case [ '1', '2' ] row,column
        else {
          firstCoordinate = firstElt.split(COMMA).map((item) => toInt(item));
        }

        if (endRangeRowColumns.length === 1) {
          // check whole row or column
          const item = endRangeRowColumns[0];
          if (isNan(item)) {
            secondCoordinate.push(-1, toInt(item));
          }
          else {
            secondCoordinate.push(toInt(item), -1);
          }
        }
        else {
          secondCoordinate = secondElt.split(COMMA).map((item) => toInt(item));
        }

        return [ firstCoordinate, secondCoordinate ];
      }
      // Row or column
      else
      {
        const nan = +firstElt;
        if (Number.isNaN(nan)) { // column names
          const colValue = encode(firstElt);
          if (firstElt === secondElt) {
            return [-1, colValue];
          }
          else {
            const colValue2 = encode(secondElt);
            // range of row or column
            return [ [-1, colValue], [-1, colValue2] ];
          }
        }
        else { // rows
          if (firstElt === secondElt) {
            return [toInt(firstElt), -1];
          }
          else {
            // range of row or column
            return [ [toInt(firstElt), -1], [toInt(secondElt), -1] ];
          }
        }

      }

    }
  });
};

const isNan = (elt: string) : boolean => {
  const nan = +elt;
  return Number.isNaN(nan);
};

/**
 * Int converter.
 *
 * @param elt Number or String (column)
 * @returns {number} Intege number type.
 */
const toInt = (elt: string) : number => {
  const nan = +elt;
  if (Number.isNaN(nan)) {
    return encode(elt);
  }

  return parseInt(elt, 10);
};
