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

test('should throw error on single cell', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B2'});

  const expected = Errors.INCORRECT_MERGE_SINGLE_CELL;

  const actual = t.throws(() => {
    range1.mergeAcross();
  }, Error);

  t.is(actual.message, expected);

});

test('should merge B2:C2 width D2:E2', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B2:C2'});
  const range2 = grid.getRange({A1: 'D2:E2'});
  const range4 = grid.getRange({A1: 'B2:E2'});
  range4.setValues([[0, 1, 2, 3]]);

  range1.mergeAcross();
  range2.mergeAcross();

  const expected = [
    [0, null, 2, null]
  ];

  t.deepEqual(range4.values, expected);
});

test('should merge B2:C5', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B2:C5'});

  range1.setValues([
    [0, 1],
    [2, 3],
    [4, 5],
    [6, 7]
  ]);

  range1.mergeAcross();

  const expected = [
    [0, null],
    [2, null],
    [4, null],
    [6, null]
  ];

  t.deepEqual(range1.values, expected);
});

test('can not merge cells with bigger merged range height inside already', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B4:C6'});
  const range2 = grid.getRange({A1: 'B3:C6'});

  range1.merge();

  const fn = () => {
    range2.mergeAcross();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE_HORIZONTAL);

});

test('should throw error if not overlaps but not contains', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B6:C7'});
  const range2= grid.getRange({A1: 'A6:B8'});

  range1.mergeAcross();

  const expected = Errors.INCORRECT_MERGE;

  const actual = t.throws(() => {
    range2.mergeAcross();
  }, Error);

  t.is(actual.message, expected);

});

test('should throw error if merge horizontally across an existing vertically merged section', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B7:B9'});
  const range2= grid.getRange({A1: 'A7:B9'});

  range1.mergeVertically();

  const expected = Errors.INCORRECT_MERGE_HORIZONTAL;

  const actual = t.throws(() => {
    range2.mergeAcross();
  }, Error);

  t.is(actual.message, expected);

});

test('should throw error if vertically across an existing horizontally merged section', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'A7:C8'});
  const range2= grid.getRange({A1: 'A6:C9'});

  range1.mergeAcross();

  const expected = Errors.INCORRECT_MERGE_VERTICALLY;

  const actual = t.throws(() => {
    range2.mergeAcross(true);
  }, Error);

  t.is(actual.message, expected);

});

test('should remove existing merge when containing it', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B5:C5'});
  const range2= grid.getRange({A1: 'B5:C6'});

  range1.mergeAcross();

  t.is(range2.getMergedRanges().length, 1);
  range2.mergeAcross();

  t.is(range2.getMergedRanges().length, 1);

});

/* TODO check
test('can merge cells with same merged range height inside already', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B4:C4'});
  const range2 = grid.getRange({A1: 'B5:C5'});
  const range3 = grid.getRange({A1: 'B3:C8'});

  range1.mergeAcross();
  range2.mergeAcross();
  range3.mergeAcross();

});
*/
