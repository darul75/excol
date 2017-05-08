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

test('should handle getColumn', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2:D4'});

  const col = range.getColumn();

  t.is(col, 2);

});

test('should handle getNumColumns', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2:D5'});

  const col = range.getNumColumns();

  t.is(col, 3);

});

test('should handle getRow', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2'});

  const col = range.getRow();

  t.is(col, 2);

});

test('should handle getNumRows', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2:D5'});

  const col = range.getNumRows();

  t.is(col, 4);

});
