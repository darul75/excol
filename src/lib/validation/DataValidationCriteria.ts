import { Errors } from '../Error'
import { decode } from '../decoder/index'
import { Cell } from '../Cell';
import { DataValidation } from '../validation/DataValidation';

export enum DataValidationCriteria {
  ANY,
  DATE_AFTER,
  DATE_BEFORE,
  DATE_BETWEEN,
  DATE_EQUAL_TO,
  DATE_IS_VALID_DATE,
  DATE_NOT_BETWEEN,
  DATE_ON_OR_AFTER,
  DATE_ON_OR_BEFORE,
  NUMBER_BETWEEN,
  NUMBER_EQUAL_TO,
  NUMBER_GREATER_THAN,
  NUMBER_GREATER_THAN_OR_EQUAL_TO,
  NUMBER_LESS_THAN,
  NUMBER_LESS_THAN_OR_EQUAL_TO,
  NUMBER_NOT_BETWEEN,
  NUMBER_NOT_EQUAL_TO,
  TEXT_CONTAINS,
  TEXT_DOES_NOT_CONTAIN,
  TEXT_EQUAL_TO,
  TEXT_IS_VALID_EMAIL,
  TEXT_IS_VALID_URL,
  VALUE_IN_LIST,
  VALUE_IN_RANGE,
  CUSTOM_FORMULA
};

export const validators = {
  10 : function(row, column, cellValue: Object, validation : DataValidation) {
    if (cellValue !== validation.getCriteriaValues()[0]) {
      throw new Error(Errors.INCORRECT_RANGE_DATA_VALIDATION(getA1Notation(row, column)));
    }
  }
};

const getA1Notation = (row: number, column: number) => decode(column) + row;
