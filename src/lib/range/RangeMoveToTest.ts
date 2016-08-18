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

test('should move range A1:B2 to C1:D2', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A1:B2'});
  const target = grid.getRange({A1: 'C1:D2'});

  range.setValues([
    ['A1', 'B1'],
    ['A2', 'B2']
  ]);

  range.setNumberFormats([
    ['N1', 'N2'],
    ['N2', 'N4']
  ]);

  range.moveTo(target);

  const rangeValue = target.values;

  const expected = [
    ['A1', 'B1'],
    ['A2', 'B2']
  ];

  t.deepEqual(rangeValue, expected);
});

test('should move range values B2:C3 to C1:D2', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'B2:C3'});
  const target = grid.getRange({A1: 'C1:D2'});

  range.setValues([
    ['B2', 'C2'],
    ['B3', 'C3']
  ]);

  range.setNumberFormats([
    ['N1', 'N2'],
    ['N2', 'N4']
  ]);

  range.moveTo(target);

  const rangeValue = range.values;
  const targetValues = target.values;

  const expectedSource = [
    [null, 'B3'],
    [null, null]
  ];

  const expectedTarget = [
    ['B2', 'C2'],
    ['B3', 'C3']
  ];

  t.deepEqual(rangeValue, expectedSource);
  t.deepEqual(targetValues, expectedTarget);

});
