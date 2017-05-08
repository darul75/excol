import { test } from 'ava';
import { DataValidationBuilder, DataValidation, DataValidationCriteria, Range } from 'excol';

test('should handle requireDate', t => {

  const builder = new DataValidationBuilder();

  builder.requireDate();
  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_IS_VALID_DATE);

});

test('should handle requireDateAfter', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.requireDateAfter(date);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_AFTER);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

test('should handle requireDateBefore', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.requireDateBefore(date);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_BEFORE);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

test('should handle requireDateBetween', t => {

  const builder = new DataValidationBuilder();
  const start = new Date();
  const end = new Date();

  builder.requireDateBetween(start, end);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_BETWEEN);
  t.is(validation.getCriteriaValues().length, 2);
  t.is(validation.getCriteriaValues()[0], start);
  t.is(validation.getCriteriaValues()[1], end);

});

test('should handle requireDateEqualTo', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.requireDateEqualTo(date);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

test('should handle requireDateNotBetween', t => {

  const builder = new DataValidationBuilder();
  const start = new Date();
  const end = new Date();

  builder.requireDateNotBetween(start, end);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_NOT_BETWEEN);
  t.is(validation.getCriteriaValues().length, 2);
  t.is(validation.getCriteriaValues()[0], start);
  t.is(validation.getCriteriaValues()[1], end);

});

test('should handle requireDateOnOrAfter', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.requireDateOnOrAfter(date);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_ON_OR_AFTER);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

test('should handle requireDateOnOrBefore', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.requireDateOnOrBefore(date);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_ON_OR_BEFORE);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

test('should handle requireFormulaSatisfied', t => {

  const builder = new DataValidationBuilder();
  const formula = 'myFormula';

  builder.requireFormulaSatisfied('myFormula');

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.CUSTOM_FORMULA);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], formula);

});

test('should handle requireNumberBetween', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberBetween(1,5);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_BETWEEN);
  t.is(validation.getCriteriaValues().length, 2);
  t.is(validation.getCriteriaValues()[0], 1);
  t.is(validation.getCriteriaValues()[1], 5);

});

test('should handle requireNumberEqualTo', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberEqualTo(2);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 2);

});

test('should handle requireNumberGreaterThan', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberGreaterThan(1);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_GREATER_THAN);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 1);

});

test('should handle requireNumberGreaterThanOrEqualTo', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberGreaterThanOrEqualTo(1);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 1);

});

test('should handle requireNumberLessThan', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberLessThan(3);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_LESS_THAN);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 3);

});

test('should handle requireNumberLessThanOrEqualTo', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberLessThanOrEqualTo(1);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_LESS_THAN_OR_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 1);

});

test('should handle requireNumberNotBetween', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberNotBetween(1,3);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_NOT_BETWEEN);
  t.is(validation.getCriteriaValues().length, 2);
  t.is(validation.getCriteriaValues()[0], 1);
  t.is(validation.getCriteriaValues()[1], 3);

});

test('should handle requireNumberNotEqualTo', t => {

  const builder = new DataValidationBuilder();

  builder.requireNumberNotEqualTo(5);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.NUMBER_NOT_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 5);

});

test('should handle requireTextContains', t => {

  const builder = new DataValidationBuilder();

  builder.requireTextContains('t');

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.TEXT_CONTAINS);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 't');

});

test('should handle requireTextDoesNotContain', t => {

  const builder = new DataValidationBuilder();

  builder.requireTextDoesNotContain('t');

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.TEXT_DOES_NOT_CONTAIN);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 't');

});

test('should handle requireTextEqualTo', t => {

  const builder = new DataValidationBuilder();

  builder.requireTextEqualTo('t');

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.TEXT_EQUAL_TO);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], 't');

});

test('should handle requireTextIsEmail', t => {

  const builder = new DataValidationBuilder();

  builder.requireTextIsEmail();

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.TEXT_IS_VALID_EMAIL);
  t.is(validation.getCriteriaValues().length, 0);

});

test('should handle requireTextIsUrl', t => {

  const builder = new DataValidationBuilder();

  builder.requireTextIsUrl();

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.TEXT_IS_VALID_URL);
  t.is(validation.getCriteriaValues().length, 0);

});

test('should handle requireValueInList', t => {

  const builder = new DataValidationBuilder();

  builder.requireValueInList(['a','b']);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.VALUE_IN_LIST);
  t.is(validation.getCriteriaValues().length, 1);
  t.deepEqual(validation.getCriteriaValues()[0], ['a', 'b']);

});

test('should handle requireValueInRange', t => {

  const builder = new DataValidationBuilder();
  const range = new Range(1,5,1,5, null);

  builder.requireValueInRange(range);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.VALUE_IN_RANGE);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], range);

});

test.todo('setAllowInvalid');

test.todo('setHelpText');

test('should handle withCriteria', t => {

  const builder = new DataValidationBuilder();
  const date = new Date();

  builder.withCriteria(DataValidationCriteria.DATE_AFTER, [date]);

  var validation: DataValidation = builder.build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_AFTER);
  t.is(validation.getCriteriaValues().length, 1);
  t.is(validation.getCriteriaValues()[0], date);

});

