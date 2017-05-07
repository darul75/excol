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

test('should range B3:C4 be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'B2:C4'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, true);
});

test('should range A2:C5 be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'A2:C5'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, true);
});

test('should range A2:D5 be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'A2:D5'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, true);
});

test('should range A2:B4 not be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'A2:B4'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, false);
});

test('should range A2:B4 not be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'A2:B4'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, false);
});

test('should range B3:B4 not be contained in B2:C4', t => {

  const grid = new Sheet(cfg);
  const outRange = grid.getRange({A1: 'B3:B4'});
  const inRange = grid.getRange({A1: 'B3:C4'});

  const isContained = outRange.contains(inRange);

  t.deepEqual(isContained, false);
});

test('should not merge single range', t => {

  const grid = new Sheet(cfg);
  const res = grid.getRange({A1: 'A1'});

  const fn = () => {
    res.merge();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE_SINGLE_CELL);

});

test('should merge small range', t => {

  const grid = new Sheet(cfg);
  const res = grid.getRange({A1: 'A1:B2'});

  //expect(function(){res.merge()}).to.throw("can not merge single cell dude !");
});

test('can not merge uncompatible cells from merge from left', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'A3:B3'});

  const fn = () => {
    range2.merge();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE);


});

test('can merge cells from left', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'A2:C4'});

  range2.merge();

});

test('can not merge uncompatible cells from merge from right', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'C3:D3'});

  const fn = () => {
    range2.merge();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE);
});

test('can merge cells from right', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'B2:D4'});

  range2.merge();

});

test('can not merge uncompatible cells from merge from top', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'B2:B4'});

  const fn = () => {
    range2.merge();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE);
});

test('can merge cells from top', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'B2:C6'});

  range2.merge();

});

test('can not merge uncompatible cells from merge from bottom', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'B2:B5'});

  const fn = () => {
    range2.merge();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE);
});

test('can merge from bigger', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'A2:D5'});

  range2.merge();

});

test('can merge and erase multiple contained merges', t => {

  const grid = new Sheet(cfg);
  const range1 = grid.getRange({A1: 'B3:C4'});

  range1.merge();

  const range2 = grid.getRange({A1: 'C5:D7'});

  range2.merge();

  const range3 = grid.getRange({A1: 'A2:E13'});

  range3.merge();

  const oneBigMerge = range3.getMergedRanges().length == 1;

  t.deepEqual(oneBigMerge, true);
});

test('can merge and verify values', t => {

  cfg.cellValue = 0;
  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B3:C4'});

  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  range.merge();

  const expected = [
    [1, null],
    [null, null]
  ];

  t.deepEqual(range.values, expected);
});


test('for single merged range', t => {

  cfg.cellValue = 0;
  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B3:C4'});

  range.setValues([
    [1, 2],
    [3, 4]
  ]);

  range.merge();
  range.breakApart();

  const expected = [
    [1, null],
    [null, null]
  ];

  t.deepEqual(range.values, expected);
});

test('should throw error for wrong single merged range', t => {

  cfg.cellValue = 0;
  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B3:C4'});

  range.merge();

  const range2 = grid.getRange({A1: 'B3:B4'});

  const fn = () => {
    range2.breakApart();
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_MERGE);


});

test('for multiple ranges merged', t => {

  cfg.cellValue = 0;
  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B3:C4'});
  const range2 = grid.getRange({A1: 'D3:E4'});
  const range3 = grid.getRange({A1: 'B1:E4'});

  range.merge();
  range2.merge();

});
