import { test } from 'ava'
import {Sheet, SheetConfig, DataValidationBuilder, DataValidation} from 'excol'
import { Errors } from 'excol'

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};


test('should handle getAllowInvalid', t => {

  const validation = new DataValidationBuilder().requireNumberEqualTo(2).setAllowInvalid(false).build();

  t.is(validation.getAllowInvalid(), false);

});

test('should handle getHelpText', t => {

  const validation = new DataValidationBuilder()
    .requireNumberEqualTo(2)
    .setAllowInvalid(false)
    .setHelpText('help')
      .build();

  t.is(validation.getHelpText(), 'help');

});


test('should throw error on bad NUMBER_EQUAL_TO', t => {

  const grid = new Sheet(cfg);
  const notation = 'B1';
  const range = grid.getRange(notation);

  const validation = new DataValidationBuilder().requireNumberEqualTo(2).build();

  range.setDataValidation(validation);

  const expected = Errors.INCORRECT_RANGE_DATA_VALIDATION(notation);

  const actual = t.throws(() => {
    range.setValue(3);
  }, Error);

  t.is(actual.message, expected);

});

test('should throw error on bad multiple NUMBER_EQUAL_TO', t => {

  const grid = new Sheet(cfg);
  const notation = 'A1:B2';
  const range = grid.getRange(notation);

  const validation = new DataValidationBuilder().requireNumberEqualTo(2).build();

  range.setDataValidation(validation);

  const expected = Errors.INCORRECT_RANGE_DATA_VALIDATION('B2');

  const actual = t.throws(() => {
    range.setValues([
      [2,2],
      [2,3]
    ]);
  }, Error);

  t.is(actual.message, expected)

});

