import {test} from 'ava';

import {
  clearContent, NewCell, Cell, clearFormat, DefaultEmptyString, DefaultFontLine, DefaultFontSize,
  DefaultFontStyle, DefaultFontWeight
} from 'excol';

test('should handle clearContent', t => {

  const cell: Cell = NewCell(1, 1, 4);

  clearContent(cell, undefined);

  t.is(cell.value, undefined);

  clearContent(cell, '5');

  t.is(cell.value, '5');

});

test('should handle clearFormat', t => {

  const cell: Cell = NewCell(1, 1, 4);

  cell.fontColor = 'blue';
  cell.fontFamily = 'family';
  cell.fontLine = 'test';
  cell.fontSize = 50;
  cell.fontStyle = 'style';
  cell.fontWeight = 'bold';

  clearFormat(cell);

  t.is(cell.fontColor, DefaultEmptyString);
  t.is(cell.fontFamily, DefaultEmptyString);
  t.is(cell.fontLine, DefaultFontLine);
  t.is(cell.fontSize, DefaultFontSize);
  t.is(cell.fontStyle, DefaultFontStyle);
  t.is(cell.fontWeight, DefaultFontWeight);

});
