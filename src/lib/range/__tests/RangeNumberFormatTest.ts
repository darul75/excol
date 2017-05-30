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

test('should get cells format from getNumberFormats', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const cellFormat = 'myFormat';

  range.setNumberFormat(cellFormat);

  const rangeFormats = range.getNumberFormats();

  const expected = [
    [cellFormat, cellFormat],
    [cellFormat, cellFormat]
  ];

  t.deepEqual(rangeFormats, expected);

  range.setNumberFormats([
    [cellFormat, cellFormat],
    [cellFormat, cellFormat]
  ]);

  const rangeFormats2 = range.getNumberFormats();

  t.deepEqual(rangeFormats2, expected);
});

test('should get cells format from getNumberFormat', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const cellFormat = 'myFormat';

  range.setNumberFormat(cellFormat);

  const rangeFormat = range.getNumberFormat();

  t.deepEqual(rangeFormat, cellFormat);

  range.setNumberFormats([
    [cellFormat, cellFormat],
    [cellFormat, cellFormat]
  ]);

  const rangeFormats2 = range.getNumberFormat();

  t.deepEqual(rangeFormats2, cellFormat);
});
