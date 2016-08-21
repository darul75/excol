import { test } from 'ava';
import { Sheet, SheetConfig } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should handle custom object', t => {

  cfg.cellValue = {
    oldValue: 'OLD',
    newValue: 'NEW',
    value: 0
  };

  const grid = new Sheet(cfg);
  const res = grid.getRange({row: 1, column: 1});

  const value = res.values[0][0];

  t.deepEqual(value.value, 0);
  t.deepEqual(value.oldValue, 'OLD');
  t.deepEqual(value.newValue, 'NEW');
});

test('should handle number', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const res = grid.getRange({row: 1, column: 1});

  const value = res.values[0][0];

  t.deepEqual(value, 0);
});
