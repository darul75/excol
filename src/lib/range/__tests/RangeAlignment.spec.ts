import { test } from 'ava';
import { Sheet, SheetConfig } from 'excol';
import { Errors } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should handle horizontalAlignment for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const left = 'left';
  const right = 'right';

  range.setHorizontalAlignment(left);

  const h = range.getHorizontalAlignment();
  const hs = range.getHorizontalAlignments();

  const expected = [
    [left, left],
    [left, left]
  ];

  t.deepEqual(hs, expected);
  t.is(h, left);

  range.setHorizontalAlignments([
    [right, right],
    [right, right]
  ]);

  const expected2 = [
    [right, right],
    [right, right]
  ];

  const color2 = range.getHorizontalAlignment();
  const colors2 = range.getHorizontalAlignments();

  t.deepEqual(colors2, expected2);
  t.is(color2, right);

});

test('should handle verticalAlignment for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const left = 'left';
  const right = 'right';

  range.setVerticalAlignment(left);

  const h = range.getVerticalAlignment();
  const hs = range.getVerticalAlignments();

  const expected = [
    [left, left],
    [left, left]
  ];

  t.deepEqual(hs, expected);
  t.is(h, left);

  range.setVerticalAlignments([
    [right, right],
    [right, right]
  ]);

  const expected2 = [
    [right, right],
    [right, right]
  ];

  const color2 = range.getVerticalAlignment();
  const colors2 = range.getVerticalAlignments();

  t.deepEqual(colors2, expected2);
  t.is(color2, right);

});
