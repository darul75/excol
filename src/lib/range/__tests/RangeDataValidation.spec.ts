import { test } from 'ava';
import {Sheet, SheetConfig, DataValidationBuilder, DataValidation, DataValidationCriteria} from 'excol';
import { Errors } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should set array/one of datavalidation for range A2:B3', t => {

  const sheet = new Sheet(cfg);
  const range = sheet.getRange({A1: 'A2:B3'});

  var validation: DataValidation = new DataValidationBuilder().requireDate().build();

  range.setDataValidation(validation);

  const vals = range.getDataValidations();
  const val = range.getDataValidation();

  t.is(vals.length, 2);
  t.is(vals[0][0].getCriteriaType(), DataValidationCriteria.DATE_IS_VALID_DATE);
  t.is(val.getCriteriaType(), DataValidationCriteria.DATE_IS_VALID_DATE);

});

