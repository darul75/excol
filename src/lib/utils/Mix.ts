import { Errors } from './Error'

export const isTwoDimArray = (o : any) : void => {
    if (!(o.constructor === Array && o[0] && o[0].constructor === Array)) {
        throw new Error(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof o));
    }
};

export const isTwoDimArrayCorrectDimensions = (o : any, rows: number, columns: number) : void => {
    isTwoDimArray(o);

    const numRows = o.length;
    const numColumns = o[0].length;

    if (numRows !== rows) {
        throw new Error(Errors.INCORRECT_RANGE_HEIGHT(numRows, rows));
    }

    if (numColumns !== columns) {
        throw new Error(Errors.INCORRECT_RANGE_WIDTH(numColumns, columns));
    }
};


export const isTwoDimArrayOfString = (o : any, rows: number, columns: number) : boolean => {
    isTwoDimArrayCorrectDimensions(o, rows, columns);

    return true;
};

export const isTwoDimArrayOfNumber = (o : any, rows: number, columns: number) : boolean => {
    isTwoDimArrayCorrectDimensions(o, rows, columns);

    return true;
};

export const isTwoDimArrayOfDataValidation = (o : any, rows: number, columns: number) : boolean => {
  isTwoDimArrayCorrectDimensions(o, rows, columns);

  // TO BE FINISHED

  return true;
};

export const isNullUndefined = (o : any) : boolean => o == null || o == undefined;
