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

test('should merge B2:B3 width C2:C3', t => {

  const grid = new Sheet(cfg);
  cfg.cellValue = 0;

  const range1 = grid.getRange({A1: 'B2:B3'});
  const range2 = grid.getRange({A1: 'C2:C3'});
  const range4 = grid.getRange({A1: 'B2:C3'});
  range4.setValues([
    [0, 1],
    [2, 3]
  ]);

  range1.mergeVertically();
  range2.mergeVertically();

  const expected = [
    [0, 1], [null, null]
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

  range1.mergeVertically();

  const expected = [
    [0, 1],
    [null, null],
    [null, null],
    [null, null]
  ];

  t.deepEqual(range1.values, expected);
});

test('can not merge cells with bigger merged range height inside already', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B4:C6'});
  const range2 = grid.getRange({A1: 'B3:C6'});

  range1.merge();

  const fn = () => {
    range2.mergeVertically();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE_VERTICALLY);

});

/* TODO check
test('can merge cells with same merged range height inside already', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B4:B5'});
  const range2 = grid.getRange({A1: 'C4:C5'});
  const range3 = grid.getRange({A1: 'B2:C7'});

  range1.mergeVertically();
  range2.mergeVertically();

});*/
