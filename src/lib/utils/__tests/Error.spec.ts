import {test} from 'ava';

import { Errors } from 'excol';

test('should handle error INCORRECT_RANGE_HEIGHT', t => {

  const error = Errors.INCORRECT_RANGE_HEIGHT(5, 8);

  t.is(error, 'Incorrect range height, was 5 but should be 8.');

});

test('should handle error INCORRECT_RANGE_WIDTH', t => {

  const error = Errors.INCORRECT_RANGE_WIDTH(5, 8);

  t.is(error, 'Incorrect range width, was 5 but should be 8.');

});

test('should handle error INCORRECT_SET_VALUES_PARAMETERS', t => {

  const error = Errors.INCORRECT_SET_VALUES_PARAMETERS('number');

  t.is(error, 'Cannot find method setValues(number).');

});
