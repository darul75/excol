import { test } from 'ava';
import { Sheet, SheetConfig } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

/* TODO check
test('should throw error on bad value types', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});

  const expectedErr = Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof '');
  const expectedErr2 = Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof 0);

  /* tslint:disable:"type-def" */
  /*
   expect(function(){range.setValues('toto')}).to.throw(expectedErr);
   expect(function(){range.setValues(0)}).to.throw(expectedErr2);
   */
  /* tslint:enable:"type-def" */
//});

test('should set array of values for range A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});

  range.setValues([
    ['A2', 'B2'],
    ['A3', 'B3']
  ]);

  const rangeValue = range.values;

  const expected = [
    ['A2', 'B2'],
    ['A3', 'B3']
  ];

  t.deepEqual(rangeValue, expected);
});

test('should get display values for range A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});

  range.setValues([
    ['A2', 'B2'],
    ['A3', 'B3']
  ]);

  const displayVal = range.getDisplayValue();
  const displayVals = range.getDisplayValues();

  const expected = [
    ['A2', 'B2'],
    ['A3', 'B3']
  ];

  t.is(displayVal, 'A2');
  t.deepEqual(displayVals, expected);

  const expected2 = [
    ['1', '2'],
    ['3', '4']
  ];

  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  const displayVal2 = range.getDisplayValue();
  const displayVals2 = range.getDisplayValues();

  t.deepEqual(displayVals2, expected2);
  t.is(displayVal2, '1');

});

