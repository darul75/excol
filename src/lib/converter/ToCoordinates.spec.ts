import {test} from 'ava';

import {toCoordinates} from 'excol';
import {toMatrix} from 'excol';

const singleA1 = {A1: 'A1'};
const singleMatrix = [['1,1']];
const singleCoordinate = [[1, 1]];

const multipleA1 = {A1: 'A1,B5'};
const multipleMatrix = [['1,1'], ['5,2']];
const multipleCoordinate = [[1, 1], [5, 2]];

const rangeA1 = {A1: 'A1:B5'};
const rangeMatrix = [['1,1', '5,2']];
const rangeCoordinate = [[[1, 1], [5, 2]]];

const multipleRangeA1 = {A1: 'A1:B5, A1:B5'};
const multipleRangeMatrix = [['1,1', '5,2'], ['1,1', '5,2']];
const multipleRangeCoordinate = [[[1, 1], [5, 2]], [[1, 1], [5, 2]]];

const singleColumnA1 = {A1: 'A:A'};
const singleColumnMatrix = [['A', 'A']];
const singleColumnCoordinate = [[-1, 1]];

const rangeColumnA1 = {A1: 'A:C'};
const rangeColumnMatrix = [['A', 'C']];
const rangeColumnCoordinate = [[[-1, 1], [-1, 3]]];

const singleRowA1 = {A1: '1:1'};
const singleRowMatrix = [['1', '1']];
const singleRowCoordinate = [[1, -1]];

const rangeRowA1 = {A1: '1:5'};
const rangeRowMatrix = [['1', '5']];
const rangeRowCoordinate = [[[1, -1], [5, -1]]];

const multipleRangeSingleColumnA1 = {A1: 'A:A,C:C,F:F'};
const multipleRangeSingleColumnMatrix = [['A', 'A'], ['C', 'C'], ['F', 'F']];
const multipleRangeSingleColumnCoordinate = [[-1, 1], [-1, 3], [-1, 6]];

const multipleSingleRowA1 = {A1: '1:1,3:3,8:8'};
const multipleSingleRowMatrix = [['1', '1'], ['3', '3'], ['8', '8']];
const multipleSingleRowCoordinate = [[1, -1], [3, -1], [8, -1]];

const multipleMixedA1 = {A1: 'A1,1:1,A1:B5,F:F'};
const multipleMixedMatrix = [['1,1'], ['1', '1'], ['1,1', '5,2'], ['F', 'F']];
const multipleMixedCoordinate = [[1, 1], [1, -1], [[1, 1], [5, 2]], [-1, 6]];

const format = JSON.stringify;

test(`convert ${format(singleA1)} to ${format(singleMatrix)}`, t => {
  const res = toMatrix(singleA1);
  t.deepEqual(res, singleMatrix);
});

test(`convert ${format(rangeA1)} to ${format(rangeMatrix)}`, t => {
  const res = toMatrix(rangeA1);
  t.deepEqual(res, rangeMatrix);
});

test(`convert ${format(multipleA1)} to ${format(multipleMatrix)}`, t => {
  const res = toMatrix(multipleA1);
  t.deepEqual(res, multipleMatrix);
});

test(`convert ${format(multipleRangeA1)} to ${format(multipleRangeMatrix)}`, t => {
  const res = toMatrix(multipleRangeA1);
  t.deepEqual(res, multipleRangeMatrix);
});

test(`convert ${format(singleColumnA1)} to ${format(singleColumnMatrix)}`, t => {
  const res = toMatrix(singleColumnA1);
  t.deepEqual(res, singleColumnMatrix);
});

test(`convert ${format(rangeColumnA1)} to ${format(rangeColumnMatrix)}`, t => {
  const res = toMatrix(rangeColumnA1);
  t.deepEqual(res, rangeColumnMatrix);
});

test(`convert ${format(singleRowA1)} to ${format(singleRowMatrix)}`, t => {
  const res = toMatrix(singleRowA1);
  t.deepEqual(res, singleRowMatrix);
});

test(`convert ${format(rangeRowA1)} to ${format(rangeRowMatrix)}`, t => {
  const res = toMatrix(rangeRowA1);
  t.deepEqual(res, rangeRowMatrix);
});

test(`convert ${format(multipleRangeSingleColumnA1)} to ${format(multipleRangeSingleColumnMatrix)}`, t => {
  const res = toMatrix(multipleRangeSingleColumnA1);
  t.deepEqual(res, multipleRangeSingleColumnMatrix);
});

test(`convert ${format(multipleSingleRowA1)} to ${format(multipleSingleRowMatrix)}`, t => {
  const res = toMatrix(multipleSingleRowA1);
  t.deepEqual(res, multipleSingleRowMatrix);
});

test(`convert ${format(multipleMixedA1)} to ${format(multipleMixedMatrix)}`, t => {
  const res = toMatrix(multipleMixedA1);
  t.deepEqual(res, multipleMixedMatrix);
});

test('should error when wrong A1 notation', t => {

  const fn = () => {
    toMatrix({A1: ''});
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, 'A1 notation is empty');
});

// test('Matrix to coordinates', t => {

test(`coordinates ${format(singleA1)} to ${format(singleCoordinate)}`, t => {
  const res = toCoordinates(singleA1);

  t.deepEqual(res, singleCoordinate);
});

test(`coordinates ${format(multipleA1)} to ${format(multipleCoordinate)}`, t => {
  const res = toCoordinates(multipleA1);
  t.deepEqual(res, multipleCoordinate);
});

test(`coordinates ${format(multipleRangeA1)} to ${format(multipleRangeCoordinate)}`, t => {
  const res = toCoordinates(multipleRangeA1);
  t.deepEqual(res, multipleRangeCoordinate);
});

test(`coordinates ${format(rangeA1)} to ${format(rangeCoordinate)}`, t => {
  const res = toCoordinates(rangeA1);
  t.deepEqual(res, rangeCoordinate);
});

test(`coordinates ${format(singleColumnA1)} to ${format(singleColumnCoordinate)}`, t => {
  const res = toCoordinates(singleColumnA1);
  t.deepEqual(res, singleColumnCoordinate);
});

test(`coordinates ${format(rangeColumnA1)} to ${format(rangeColumnCoordinate)}`, t => {
  const res = toCoordinates(rangeColumnA1);
  t.deepEqual(res, rangeColumnCoordinate);
});

test(`coordinates ${format(singleRowA1)} to ${format(singleRowCoordinate)}`, t => {
  const res = toCoordinates(singleRowA1);
  t.deepEqual(res, singleRowCoordinate);
});

test(`coordinates ${format(rangeRowA1)} to ${format(rangeRowCoordinate)}`, t => {
  const res = toCoordinates(rangeRowA1);
  t.deepEqual(res, rangeRowCoordinate);
});

test(`coordinates ${format(multipleRangeSingleColumnA1)} to ${format(multipleRangeSingleColumnCoordinate)}`, t => {
  const res = toCoordinates(multipleRangeSingleColumnA1);
  t.deepEqual(res, multipleRangeSingleColumnCoordinate);
});

test(`coordinates ${format(multipleSingleRowA1)} to ${format(multipleSingleRowCoordinate)}`, t => {
  const res = toCoordinates(multipleSingleRowA1);
  t.deepEqual(res, multipleSingleRowCoordinate);
});

test(`coordinates ${format(multipleMixedA1)} to ${format(multipleMixedCoordinate)}`, t => {
  const res = toCoordinates(multipleMixedA1);
  t.deepEqual(res, multipleMixedCoordinate);
});

test('should error when wrong A1 notation', t => {

  const fn = () => {
    toCoordinates({A1: ''});
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, 'A1 notation is empty');

});
