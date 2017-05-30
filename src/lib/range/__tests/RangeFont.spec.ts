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


test('should handle font colors for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const blue = 'blue';
  const red = 'red';

  range.setFontColor(blue);

  const color = range.getFontColor();
  const colors = range.getFontColors();

  const expected = [
    [blue, blue],
    [blue, blue]
  ];

  t.deepEqual(colors, expected);
  t.is(color, blue);

  range.setFontColors([
    [red, red],
    [red, red]
  ]);

  const expected2 = [
    [red, red],
    [red, red]
  ];

  const color2 = range.getFontColor();
  const colors2 = range.getFontColors();

  t.deepEqual(colors2, expected2);
  t.is(color2, red);

});

test('should handle font family for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const fa1 = 'fa1';
  const fa2 = 'fa2';

  range.setFontFamily(fa1);

  const fa = range.getFontFamily();
  const fas = range.getFontFamilies();

  const expected = [
    [fa1, fa1],
    [fa1, fa1]
  ];

  t.deepEqual(fas, expected);
  t.is(fa, fa1);

  range.setFontFamilies([
    [fa2, fa2],
    [fa2, fa2]
  ]);

  const expected2 = [
    [fa2, fa2],
    [fa2, fa2]
  ];

  const fa3 = range.getFontFamily();
  const fas2 = range.getFontFamilies();

  t.deepEqual(fas2, expected2);
  t.is(fa3, fa2);

});

test('should handle font line for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const fl1 = 'fl1';
  const fl2 = 'fl2';

  range.setFontLine(fl1);

  const fl = range.getFontLine();
  const fls = range.getFontLines();

  const expected = [
    [fl1, fl1],
    [fl1, fl1]
  ];

  t.deepEqual(fls, expected);
  t.is(fl, fl1);

  range.setFontLines([
    [fl2, fl2],
    [fl2, fl2]
  ]);

  const expected2 = [
    [fl2, fl2],
    [fl2, fl2]
  ];

  const fl3 = range.getFontLine();
  const fls2 = range.getFontLines();

  t.deepEqual(fls2, expected2);
  t.is(fl3, fl2);

});

test('should handle font size for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const fs1 = 14;
  const fs2 = 20;

  range.setFontSize(fs1);

  const fs = range.getFontSize();
  const fss = range.getFontSizes();

  const expected = [
    [fs1, fs1],
    [fs1, fs1]
  ];

  t.deepEqual(fss, expected);
  t.is(fs, fs1);

  range.setFontSizes([
    [fs2, fs2],
    [fs2, fs2]
  ]);

  const expected2 = [
    [fs2, fs2],
    [fs2, fs2]
  ];

  const fs3 = range.getFontSize();
  const fss2 = range.getFontSizes();

  t.deepEqual(fss2, expected2);
  t.is(fs3, fs2);

});

test('should handle font style for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const fs1 = 'fs1';
  const fs2 = 'fs2';

  range.setFontStyle(fs1);

  const fs = range.getFontStyle();
  const fss = range.getFontStyles();

  const expected = [
    [fs1, fs1],
    [fs1, fs1]
  ];

  t.deepEqual(fss, expected);
  t.is(fs, fs1);

  range.setFontStyles([
    [fs2, fs2],
    [fs2, fs2]
  ]);

  const expected2 = [
    [fs2, fs2],
    [fs2, fs2]
  ];

  const fs3 = range.getFontStyle();
  const fss2 = range.getFontStyles();

  t.deepEqual(fss2, expected2);
  t.is(fs3, fs2);

});

test('should handle font weight for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange('A2:B3');
  const fw1 = 'fw1';
  const fw2 = 'fw2';

  range.setFontWeight(fw1);

  const fw = range.getFontWeight();
  const fws = range.getFontWeights();

  const expected = [
    [fw1, fw1],
    [fw1, fw1]
  ];

  t.deepEqual(fws, expected);
  t.is(fw, fw1);

  range.setFontWeights([
    [fw2, fw2],
    [fw2, fw2]
  ]);

  const expected2 = [
    [fw2, fw2],
    [fw2, fw2]
  ];

  const fw3 = range.getFontWeight();
  const fws2 = range.getFontWeights();

  t.deepEqual(fws2, expected2);
  t.is(fw3, fw2);

});
