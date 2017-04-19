import { test } from 'ava';
import { Range, Sheet, Spreadsheet, SpreadsheetApp } from 'excol';

const name = 'newSpreadsheet';
const name1 = 'newSpreadsheet1';

test('should handle create new spreadsheet', t => {

  const spreadsheetApp = new SpreadsheetApp({});

  const spreadsheet = spreadsheetApp.create(name);
  const sheets = spreadsheet.getSheets();

  t.true(sheets.length === 1);
  t.true(sheets[0].getName() === 'Sheet1');

});

test('should handle get active spreadsheet', t => {

  const spreadsheetApp = new SpreadsheetApp({});

  const spreadsheet = spreadsheetApp.create(name);


  t.true(spreadsheetApp.getActive().getName() === name);

});

test('should handle get active range', t => {

  const spreadsheetApp = new SpreadsheetApp({});

  const spreadsheet = spreadsheetApp.create(name);

  let range: Range = spreadsheet.getActiveRange();

  t.true(range.getLastRow() === 1001);
  t.true(range.getLastColumn() === 1001);

  const newActiveRange = spreadsheet.getActiveSheet().getRange({A1: 'A1:B3'});

  spreadsheet.setActiveRange(newActiveRange);

  range = spreadsheet.getActiveRange();

  t.true(range.getLastRow() === 4);
  t.true(range.getLastColumn() === 3);

});

test('should handle get active sheet', t => {

  const spreadsheetApp = new SpreadsheetApp({});

  const spreadsheet = spreadsheetApp.create(name);

  const secondSheet: Sheet = spreadsheet.insertSheet();

  let sheet: Sheet = spreadsheet.getActiveSheet();

  t.true(sheet.getName() === secondSheet.getName());

  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);

  t.true(spreadsheet.getActiveSheet().getName() === spreadsheet.getSheets()[0].getName());

});

test('should handle get active spreadsheet', t => {

  const spreadsheetApp = new SpreadsheetApp({});

  const spreadsheet = spreadsheetApp.create(name);
  spreadsheetApp.create(name1);

  let ssheet: Spreadsheet = spreadsheetApp.getActiveSpreadsheet();

  t.true(ssheet.getName() === name1);

  spreadsheetApp.setActiveSpreadsheet(spreadsheet);

  t.true(spreadsheetApp.getActiveSpreadsheet().getName() === name);

});


test.todo('flush');
test.todo('newDataValidation');
test.todo('getUi');
