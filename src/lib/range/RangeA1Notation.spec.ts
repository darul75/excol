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

test('should throw error in A1 notation for A[]A....z', t => {

  const grid = new Sheet(cfg);

  const fn = () => {
    grid.getRange({A1: 'A[]A....z'});
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, Errors.INCORRECT_A1_NOTATION);

});

test('should get A1 notation for B2', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2'});

  const expected = 'B2';

  t.deepEqual(range.getA1Notation(), expected);
});

test('should get A1 notation for A:A', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A:A'});

  const expected = 'A1:A20';

  t.deepEqual(range.getA1Notation(), expected);
});

test('should get A1 notation for A1:B5', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A1:B5'});

  const expected = 'A1:B5';

  t.deepEqual(range.getA1Notation(), expected);
});

test('should get A1 notation for B5:C', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B5:C'});

  const expected = 'B5:C20';

  t.deepEqual(range.getA1Notation(), expected);
});

test('should get A1 notation for 2:4', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: '2:4'});

  const expected = 'A2:T4';

  t.deepEqual(range.getA1Notation(), expected);
});
