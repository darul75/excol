import {test} from 'ava';

import { Spreadsheet } from 'excol';

const spreadSheet = new Spreadsheet({name:'sheet1', numRows: 5, numColumns: 5});
const range = spreadSheet.getRange('A1:B5');

const namedRangeName = 'myNamedRange';
const namedRange = spreadSheet.addNamedRanges(namedRangeName, range);

test('should handle getName', t => {

  t.is(namedRange.getName(), namedRangeName);

});

test('should handle getRange', t => {

  t.is(namedRange.getRange(), range);
  t.is(namedRange.getRange().getA1Notation(), 'A1:B5');

});

test('should handle remove', t => {

  t.is(spreadSheet.getNamedRanges().length, 1);

  namedRange.remove();

  t.is(spreadSheet.getNamedRanges().length, 0);

});

test('should handle setName', t => {

  namedRange.setName('newName');

  t.is(namedRange.getName(), 'newName');

});

test('should handle setRange', t => {

  const range = spreadSheet.getRange('A5:B5');

  namedRange.setRange(range);

  t.is(namedRange.getRange(), range);
  t.is(namedRange.getRange().getA1Notation(), 'A5:B5');

});
