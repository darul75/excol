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

test('should set array of backgrounds for range A2:B3', t => {

  const sheet = new Sheet(cfg);
  const range = sheet.getRange('A2:B3');

  range.setBackgrounds([
    ['A2', 'B2'],
    ['A3', 'B3'],
  ]);

  const bgs = range.getBackgrounds();
  const bg = range.getBackground();

  const expected = [
    ['A2', 'B2'],
    ['A3', 'B3']
  ];

  t.deepEqual(bgs, expected);
  t.is(bg, 'A2');
});

/*
 test.only('should throw error on bad colors types', t => {

 const grid = new Sheet(cfg);
 const range = grid.getRange({A1: 'A2:B3'});

 const expectedErr = Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof '');
 const expectedErr2 = Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof 0);


 ///* tslint:disable */

/*
 const actual = t.throws(() => {
 range.setBackgrounds('toto');
 }, Error);

 t.is(actual.message, expectedErr);


 //expect(function(){range.setBackgrounds(0)}).to.throw(expectedErr2);

 /* tslint:enable */


//});
