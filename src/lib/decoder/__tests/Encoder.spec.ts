import {test} from 'ava';

import {encode, decode} from 'excol';

test('should encode column name in alpha to number', t => {
  const res = encode('AB');
  t.deepEqual(res, 28);
});

test('should error encoding wrong column name', t => {
  const fn = () => {
    encode('');
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, 'column name is empty');

});

test('should decode column index to alpha', t => {
  const res = decode(100);
  t.deepEqual(res, 'CV');
});

test('should error encoding wrong column index', t => {

  const fn = () => {
    decode(800000);
  };
  const error = t.throws(() => {
    fn();
  }, Error);

  t.is(error.message, 'column number out of range: 800000, limit is 456976');

});
