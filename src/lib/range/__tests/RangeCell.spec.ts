import {test} from 'ava';

import {Sheet, SheetConfig} from 'excol';
import {Errors} from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should handle getCell', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B7:C8'});

  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  const cellRange = range.getCell(2,1);
  const cellRange2 = range.getCell(1,2);
  const cellRange3 = range.getCell(2,2);

  t.deepEqual(cellRange.values, [[3]]);
  t.deepEqual(cellRange2.values, [[2]]);
  t.deepEqual(cellRange3.values, [[4]]);

});

test('should getCell throw out of range', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B7:C8'});

  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  const expected = Errors.INCORRECT_RANGE_CELL;

  const actual = t.throws(() => {
    range.getCell(3,1);
  }, Error);

  t.is(actual.message, expected);

  const actual2 = t.throws(() => {
    range.getCell(1,3);
  }, Error);

  t.is(actual2.message, expected);

});
