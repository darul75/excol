import { test } from 'ava';
import { Sheet, SheetConfig } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'sheet',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should append new line at the end', t => {

  const sheet = new Sheet(cfg);
  const range = sheet.getRange('A1:B2');
  ;
  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  t.deepEqual(range.getLastRow(), 3);

  sheet.appendRow([5, 6]);

  const range2 = sheet.getRange('A1:B3');

  t.deepEqual(range2.values,
    [
      [1, 2],
      [3, 4],
      [5, 6]
    ]
  );
});
