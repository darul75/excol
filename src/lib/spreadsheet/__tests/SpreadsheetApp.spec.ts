import { test } from 'ava';
import {
  Range, Sheet, Spreadsheet, SpreadsheetApp, DataValidationBuilder,
  DataValidation, DataValidationCriteria
} from 'excol';

const name = 'newSpreadsheet';
const name1 = 'newSpreadsheet1';

test('should handle create new spreadsheet', t => {

  const spreadsheet = SpreadsheetApp.create(name);
  const sheets = spreadsheet.getSheets();

  t.true(sheets.length === 1);
  t.true(sheets[0].getName() === 'Sheet1');

});

test('should handle get active spreadsheet', t => {

  const spreadsheet = SpreadsheetApp.create(name);


  t.true(SpreadsheetApp.getActive().getName() === name);

});

test('should handle get active range', t => {

  const spreadsheet = SpreadsheetApp.create(name);

  let range: Range = spreadsheet.getActiveRange();

  t.true(range.getLastRow() === 1001);
  t.true(range.getLastColumn() === 1001);

  const newActiveRange = spreadsheet.getActiveSheet().getRange('A1:B3');

  spreadsheet.setActiveRange(newActiveRange);

  range = spreadsheet.getActiveRange();

  t.true(range.getLastRow() === 4);
  t.true(range.getLastColumn() === 3);

});

test('should handle get active sheet', t => {

  const spreadsheet = SpreadsheetApp.create(name);

  const secondSheet: Sheet = spreadsheet.insertSheet();

  let sheet: Sheet = spreadsheet.getActiveSheet();

  t.true(sheet.getName() === secondSheet.getName());

  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);

  t.true(spreadsheet.getActiveSheet().getName() === spreadsheet.getSheets()[0].getName());

});

test('should handle get active spreadsheet', t => {

  const spreadsheet = SpreadsheetApp.create(name);
  SpreadsheetApp.create(name1);

  let ssheet: Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  t.true(ssheet.getName() === name1);

  SpreadsheetApp.setActiveSpreadsheet(spreadsheet);

  t.true(SpreadsheetApp.getActiveSpreadsheet().getName() === name);

});

test('should handle flush', t => {

  SpreadsheetApp.flush();

  t.true(1 === 1);

});

test('should handle newDataValidation', t => {

  const builder: DataValidationBuilder = SpreadsheetApp.newDataValidation();

  const validation: DataValidation = builder.requireDate().build();

  t.is(validation.getCriteriaType(), DataValidationCriteria.DATE_IS_VALID_DATE);

});

test('should handle setActiveSheet', t => {

  const activeSheet = SpreadsheetApp.getActiveSheet();

  SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  SpreadsheetApp.getActiveSpreadsheet().insertSheet();

  SpreadsheetApp.setActiveSheet(activeSheet);

  t.is(activeSheet.getName(), 'Sheet1');

});

test('should handle setActiveRange', t => {

  const activeSheet = SpreadsheetApp.getActiveSheet();

  SpreadsheetApp.setActiveRange(activeSheet.getRange('A1:B3'));

  const A1 = SpreadsheetApp.getActiveRange().getA1Notation();

  t.is(A1, 'A1:B3');

});

test('should handle getUi', t => {

  SpreadsheetApp.getUi();

  t.is(2,2);

});
