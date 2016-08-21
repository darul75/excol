import { test } from 'ava';
import { isTwoDimArray, isTwoDimArrayCorrectDimensions } from '../utils/mix'
import { Errors } from '../lib/Error'

//describe('Mix', () => {
  //  describe('isTwoDimArray method ', () => {
        test('should test common types and 2 dimensional array', () => {
          /*
            expect(function(){isTwoDimArray(0)}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof 0));
            expect(function(){isTwoDimArray('')}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof ''));
            expect(function(){isTwoDimArray(false)}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof false));
            expect(function(){isTwoDimArray([false])}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof [false]));
            expect(function(){isTwoDimArray({})}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof {}));
            expect(function(){isTwoDimArray([])}).to.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof []));
            expect(function(){isTwoDimArray([[]])}).to.not.throw(Errors.INCORRECT_SET_VALUES_PARAMETERS(typeof [[]]));
            */
        });
        test('should test 2 dimensional array dimensions rows', () => {
            const emptyArray = [[]];
            const rightRowsArray = [[], []];
          /*
            expect(function(){isTwoDimArrayCorrectDimensions(emptyArray, 2, 2)}).to.throw(Errors.INCORRECT_RANGE_HEIGHT(1, 2));
            expect(function(){isTwoDimArrayCorrectDimensions(rightRowsArray, 2, 2)}).to.not.throw(Errors.INCORRECT_RANGE_HEIGHT(1, 2));
            */
        });
        test('should test 2 dimensional array dimensions columns', () => {
            const emptyArray = [[]];
            const rightRowsArray = [[], []];
          /*
            expect(function(){isTwoDimArrayCorrectDimensions(emptyArray, 1, 1)}).to.throw(Errors.INCORRECT_RANGE_WIDTH(0, 1));
            expect(function(){isTwoDimArrayCorrectDimensions(rightRowsArray, 1, 1)}).to.not.throw(Errors.INCORRECT_RANGE_WIDTH(0, 1));
            */
        });

