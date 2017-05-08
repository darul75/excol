import { test } from 'ava';
import {isTwoDimArray, isTwoDimArrayCorrectDimensions, Errors, isTwoDimArrayOfDataValidation} from "excol";

test('should throw on isTwoDimArray with bad parameters', t => {

  const o = 200;

  const expected = Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof o);

  const actual = t.throws(() => {
    isTwoDimArray(o);
  }, Error);

  t.is(actual.message, expected);

});

test('should throw on isTwoDimArrayCorrectDimensions with bad parameters', t => {

  const o1 = [
    [1, 2],
    [3, 4]
  ];

  const expected = Errors.INCORRECT_RANGE_HEIGHT(2, 3);

  const actual = t.throws(() => {
    isTwoDimArrayCorrectDimensions(o1, 3, 2);
  }, Error);

  t.is(actual.message, expected);

  const expected2 = Errors.INCORRECT_RANGE_WIDTH(2, 3);

  const actual2 = t.throws(() => {
    isTwoDimArrayCorrectDimensions(o1, 2, 3);
  }, Error);

  t.is(actual2.message, expected2);

});

test('should throw on isTwoDimArrayOfDataValidation with bad parameters', t => {

  const o1 = [
    [1, 2],
    [3, 4]
  ];

  const expected = Errors.INCORRECT_RANGE_HEIGHT(2, 3);

  const actual = t.throws(() => {
    isTwoDimArrayOfDataValidation(o1, 3, 2);
  }, Error);

  t.is(actual.message, expected);

  const expected2 = Errors.INCORRECT_RANGE_WIDTH(2, 3);

  const actual2 = t.throws(() => {
    isTwoDimArrayOfDataValidation(o1, 2, 3);
  }, Error);

  t.is(actual2.message, expected2);

});





