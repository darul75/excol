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

test('should clear erase all', t => {

  cfg.cellValue = 0;

  const grid = new Sheet(cfg);
  const range = grid.getRange('A1:B2');

  range.setBackground('black');
  range.setFontColor('blue');
  range.setFontFamily('arial');
  range.setFontLine('stroke');
  range.setFontSize(3);
  range.setFontStyle('italic');
  range.setFontWeight('bold');
  range.setNumberFormat('numberFormat');
  range.setValues([[1,2], [3,4]]);

  range.merge();

  range.clear();

  const emptyStrings = [ [ '', '' ], [ '', '' ] ];
  const defaultFontLines = [ [ 'none', 'none' ], [ 'none', 'none' ] ];
  const defaultFontSizes = [ [ 10, 10 ], [ 10, 10 ] ];
  const defaultFontStyle = [ [ 'normal', 'normal' ], [ 'normal', 'normal' ] ];
  const defaultValues = [ [ cfg.cellValue, cfg.cellValue ], [ cfg.cellValue, cfg.cellValue ] ];

  t.deepEqual(range.getBackgrounds(), emptyStrings);
  t.deepEqual(range.getFontColors(), emptyStrings);
  t.deepEqual(range.getFontFamilies(), emptyStrings);
  t.deepEqual(range.getFontLines(), defaultFontLines);
  t.deepEqual(range.getFontSizes(), defaultFontSizes);
  t.deepEqual(range.getFontStyles(), defaultFontStyle);
  t.deepEqual(range.getFontWeights(), defaultFontStyle);
  t.deepEqual(range.getNumberFormats(), emptyStrings);
  t.deepEqual(range.values, defaultValues);

});
