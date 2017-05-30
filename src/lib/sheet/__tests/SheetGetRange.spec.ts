import {test} from 'ava';
import {Sheet, SheetConfig} from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test Sheet',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should get address for A1', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('A1');

  t.deepEqual(res.values[0][0].value, '0-0');
});

test('should get address for A:A', t => {

  const sheet = new Sheet(cfg);

  const res = sheet.getRange('A:A');

  t.deepEqual(res.values.length, 20);
  t.deepEqual(res.values[0].length, 1);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[19][0].value, '19-0');

});

test('should get address for B:D', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B:D');

  t.deepEqual(res.values.length, 20);
  t.deepEqual(res.values[0].length, 3);
  t.deepEqual(res.values[0][0].value, '0-1');
  t.deepEqual(res.values[19][2].value, '19-3');

});

test('should get address for D:B', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('D:B');

  t.deepEqual(res.values.length, 20);
  t.deepEqual(res.values[0].length, 3);
  t.deepEqual(res.values[0][0].value, '0-1');
  t.deepEqual(res.values[19][2].value, '19-3');

});

test('should get address for 8:8', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('8:8');

  t.deepEqual(res.values.length, 1);
  t.deepEqual(res.values[0].length, 20);
  t.deepEqual(res.values[0][0].value, '7-0');
  t.deepEqual(res.values[0][19].value, '7-19');

});

test('should get address for 1:5', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('1:5');

  t.deepEqual(res.values.length, 5);
  t.deepEqual(res.values[0].length, 20);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[4][19].value, '4-19');

});

test('should get address for 5:1', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('1:5');

  t.deepEqual(res.values.length, 5);
  t.deepEqual(res.values[0].length, 20);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[4][19].value, '4-19');

});

test('should get address for A1:B5', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('A1:B5');

  t.deepEqual(res.values.length, 5);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[4][1].value, '4-1');

});

test('should get address for B5:A1', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B5:A1');

  t.deepEqual(res.values.length, 5);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[4][1].value, '4-1');

});

test('should get address for B2:C5', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B2:C5');

  t.deepEqual(res.values.length, 4);
  t.deepEqual(res.values[0][0].value, '1-1');
  t.deepEqual(res.values[3][1].value, '4-2');

});

test('should get address for C5:B2', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B2:C5');

  t.deepEqual(res.values.length, 4);
  t.deepEqual(res.values[0][0].value, '1-1');
  t.deepEqual(res.values[3][1].value, '4-2');

});

test('should get address for A1:A2', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('A1:A2');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[1][0].value, '1-0');

});

test('should get address for A2:A1', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('A2:A1');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '0-0');
  t.deepEqual(res.values[1][0].value, '1-0');

});

test('should get address for B5:C', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B5:C');

  t.deepEqual(res.values.length, 16);
  t.deepEqual(res.values[0][0].value, '4-1');
  t.deepEqual(res.values[15][0].value, '19-1');

});

test('should get address for C:B5', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('C:B5');

  t.deepEqual(res.values.length, 16);
  t.deepEqual(res.values[0][0].value, '4-1');
  t.deepEqual(res.values[15][0].value, '19-1');

});

test('should get address for A7:8', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('A7:8');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '6-0');
  t.deepEqual(res.values[1][19].value, '7-19');

});

test('should get address for 8:A7', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('8:A7');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '6-0');
  t.deepEqual(res.values[1][19].value, '7-19');

});

test('should get address for B7:8', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('B7:8');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '6-1');
  t.deepEqual(res.values[1][18].value, '7-19');

});

test('should get address for 8:B7', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange('8:B7');

  t.deepEqual(res.values.length, 2);
  t.deepEqual(res.values[0][0].value, '6-1');
  t.deepEqual(res.values[1][18].value, '7-19');

});

test('should get multiple address for C5:D9,G9:H16', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRanges({A1: 'C5:D9,G9:H16'});

  t.deepEqual(res.length, 2);

  const firstRange = res[0];
  const secondRange = res[1];

  t.deepEqual(firstRange.values.length, 5);
  t.deepEqual(firstRange.values[0].length, 2);
  t.deepEqual(firstRange.values[0][0].value, '4-2');
  t.deepEqual(firstRange.values[4][1].value, '8-3');

  t.deepEqual(secondRange.values.length, 8);
  t.deepEqual(secondRange.values[0].length, 2);
  t.deepEqual(secondRange.values[0][0].value, '8-6');
  t.deepEqual(secondRange.values[7][1].value, '15-7');

});

test('should get multiple address for C5:D9,F:F', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRanges({A1: 'C5:D9,F:F'});

  t.deepEqual(res.length, 2);

  const firstRange = res[0];
  const secondRange = res[1];

  t.deepEqual(firstRange.values.length, 5);
  t.deepEqual(firstRange.values[0].length, 2);
  t.deepEqual(firstRange.values[0][0].value, '4-2');
  t.deepEqual(firstRange.values[4][1].value, '8-3');

  t.deepEqual(secondRange.values.length, 20);
  t.deepEqual(secondRange.values[0].length, 1);
  t.deepEqual(secondRange.values[0][0].value, '0-5');
  t.deepEqual(secondRange.values[19][0].value, '19-5');

});

test('should get multiple address for C5:D9,1:5', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRanges({A1: 'C5:D9,1:5'});

  t.deepEqual(res.length, 2);

  const firstRange = res[0];
  const secondRange = res[1];

  t.deepEqual(firstRange.values.length, 5);
  t.deepEqual(firstRange.values[0].length, 2);
  t.deepEqual(firstRange.values[0][0].value, '4-2');
  t.deepEqual(firstRange.values[4][1].value, '8-3');

  t.deepEqual(secondRange.values.length, 5);
  t.deepEqual(secondRange.values[0].length, 20);
  t.deepEqual(secondRange.values[0][0].value, '0-0');
  t.deepEqual(secondRange.values[4][19].value, '4-19');

});

test('should get multiple address for A:A,C:C', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRanges({A1: 'A:A,C:C'});

  t.deepEqual(res.length, 2);

  const firstRange = res[0];
  const secondRange = res[1];

  t.deepEqual(firstRange.values.length, 20);
  t.deepEqual(firstRange.values[0].length, 1);
  t.deepEqual(firstRange.values[0][0].value, '0-0');
  t.deepEqual(firstRange.values[19][0].value, '19-0');

  t.deepEqual(secondRange.values.length, 20);
  t.deepEqual(secondRange.values[0].length, 1);
  t.deepEqual(secondRange.values[0][0].value, '0-2');
  t.deepEqual(secondRange.values[19][0].value, '19-2');
});

test('should get address for row1, column1', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange(1, 1);

  t.deepEqual(res.values[0][0].value, '0-0');
});

test('should get address for row1, column1, numRows4', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange(1, 1, 4);

  t.deepEqual(res.values.length, 4);
  t.deepEqual(res.values[3][0].value, '3-0');
});

test('should get address for row1, column1, numRows4, numColumns2', t => {

  const sheet = new Sheet(cfg);
  const res = sheet.getRange(1, 1, 4, 2);

  t.deepEqual(res.values.length, 4);
  t.deepEqual(res.values[3][1].value, '3-1');
});
