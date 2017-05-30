import { test } from 'ava';
import {Range, Sheet, Spreadsheet, SpreadsheetConfig, Errors} from 'excol';
import { User } from 'excol';

const email: string = 'test@com.net';

const cfg: SpreadsheetConfig = {
  name: 'test lib'
};

const cfg2: SpreadsheetConfig = {
  name: 'test lib',
  numRows: 4,
  numColumns: 4
};

test('should handle addEditor ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const user: User = new User(email);

  spreadsheet.addEditor(user);

  t.deepEqual(spreadsheet.getEditors().length, 1);

  spreadsheet.addEditor(email);

  t.deepEqual(spreadsheet.getEditors().length, 2);

});

test('should handle addEditors ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  spreadsheet.addEditors([email, email]);

  t.deepEqual(spreadsheet.getEditors().length, 2);

});

test('should handle addViewer ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const user: User = new User(email);

  spreadsheet.addViewer(user);

  t.deepEqual(spreadsheet.getViewers().length, 1);

  spreadsheet.addViewer(email);

  t.deepEqual(spreadsheet.getViewers().length, 2);

});

test('should handle addViewers ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  spreadsheet.addViewers([email, email]);

  t.deepEqual(spreadsheet.getViewers().length, 2);

});

test('should handle appendRow', t => {

  const spreadsheet = new Spreadsheet(cfg);

  t.true(spreadsheet.getLastRow() == 1);

  spreadsheet.appendRow([email, email]);

  const range: Range = spreadsheet.getRange('D6:E8');

  range.setValues([
    [1,2],
    [3,4],
    [5,6]
  ]);

  t.true(spreadsheet.getLastRow() === 9);
  t.true(spreadsheet.getLastColumn() === 6);

});

test('should handle deleteColumn', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.deleteColumn(2);

  const values = spreadsheet.getActiveSheet().getSheetValues(1, 1, 4, 4);

  t.true(values[0][0].value === '0-0');
  t.true(values[0][1].value === '0-2');
  t.true(values[1][0].value === '1-0');
  t.true(values[1][1].value === '1-2');

});

test('should handle deleteColumns', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.deleteColumns(2, 2);

  const values = spreadsheet.getActiveSheet().getSheetValues(1, 1, 4, 4);

  t.true(values[0][0].value === '0-0');
  t.true(values[0][1].value === '0-3');
  t.true(values[1][0].value === '1-0');
  t.true(values[1][1].value === '1-3');

});

test('should handle deleteRow', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.deleteRow(2);

  const values = spreadsheet.getActiveSheet().getSheetValues(1, 1, 4, 4);

  t.true(values[0][0].value === '0-0');
  t.true(values[1][0].value === '2-0');

});

test('should handle deleteRows', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.deleteRows(2, 2);

  const values = spreadsheet.getActiveSheet().getSheetValues(1, 1, 4, 4);

  t.true(values[0][0].value === '0-0');
  t.true(values[1][0].value === '3-0');

});

test('should handle deleteSheet', t => {

  const spreadsheet = new Spreadsheet(cfg);

  spreadsheet.insertSheet();

  const firstSheet: Sheet = spreadsheet.getSheets()[0];

  spreadsheet.deleteSheet(firstSheet);

  t.deepEqual(spreadsheet.getSheets().length, 1);
  t.deepEqual(spreadsheet.getSheets()[0].getName(), 'Sheet2');

});

test('should handle getDataRange', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  const sheet: Sheet = spreadsheet.getActiveSheet();

  const range: Range = sheet.getRange(3, 2, 2, 2);

  range.setValues([
    [1,2],
    [3,4]
  ]);

  const values = spreadsheet.getDataRange().values;

  t.deepEqual(values[2][1], 1);
  t.deepEqual(values[2][2], 2);

});

test('should handle getNumSheets', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.insertSheet();
  spreadsheet.insertSheet();
  spreadsheet.insertSheet();

  t.true(spreadsheet.getNumSheets() === 4);

});

test('should handle getSheetByName', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.insertSheet();
  spreadsheet.insertSheet();
  spreadsheet.insertSheet();

  const sheetName = 'Sheet1';

  const sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet !== null)
    t.true(sheet.getName() === sheetName);
  else
    t.fail();

  const sheet2 = spreadsheet.getSheetByName('NONEXISTING');

  t.is(sheet2, null);

});

test('should handle insertColumnAfter', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertColumnAfter(2);

  const values = spreadsheet.getSheetValues(1, 1, 4, 5);

  t.true(values[0][0] !== null);
  t.true(values[0][2] === null);
  t.true(spreadsheet.getLastColumn() === 6);
});

test('should handle insertColumnBefore', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertColumnBefore(2);

  const values = spreadsheet.getSheetValues(1, 1, 4, 5);

  t.true(values[0][0] !== null);
  t.true(values[0][1] === null);
  t.true(spreadsheet.getLastColumn() === 6);
});

test('should handle insertColumnsAfter', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertColumnsAfter(2, 2);

  const values = spreadsheet.getSheetValues(1, 1, 4, 5);

  t.true(values[0][0] !== null);
  t.true(values[0][2] === null);
  t.true(values[0][3] === null);
  t.true(spreadsheet.getLastColumn() === 6);
});

test('should handle insertColumnsBefore', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertColumnsBefore(2, 2);

  const values = spreadsheet.getSheetValues(1, 1, 4, 5);

  t.true(values[0][0] !== null);
  t.true(values[0][1] === null);
  t.true(values[0][2] === null);
  t.true(spreadsheet.getLastColumn() === 6);
});

test('should handle insertRowAfter', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertRowAfter(2);

  const values = spreadsheet.getSheetValues(1, 1, 5, 4);

  t.true(values[0][0] !== null);
  t.true(values[2][0] === null);
  t.true(spreadsheet.getLastRow() === 6);
});

test('should handle insertRowsAfter', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertRowsAfter(2, 2);

  const values = spreadsheet.getSheetValues(1, 1, 5, 4);

  t.true(values[0][0] !== null);
  t.true(values[2][0] === null);
  t.true(values[3][0] === null);
  t.true(spreadsheet.getLastRow() === 7);
});

test('should handle insertRowBefore', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertRowBefore(2);

  const values = spreadsheet.getSheetValues(1, 1, 5, 4);

  t.true(values[0][0] !== null);
  t.true(values[1][0] === null);
  t.true(spreadsheet.getLastRow() === 6);
});

test('should handle insertRowsBefore', t => {

  const spreadsheet = new Spreadsheet(cfg2);

  spreadsheet.getRange('A1:D4').setValues([
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4],
    [1,2,3,4]
  ]);

  t.true(spreadsheet.getLastColumn() === 5);

  spreadsheet.insertRowsBefore(2, 2);

  const values = spreadsheet.getSheetValues(1, 1, 5, 4);

  t.true(values[0][0] !== null);
  t.true(values[1][0] === null);
  t.true(values[2][0] === null);
  t.true(spreadsheet.getLastRow() === 7);
});

test('should get one sheet created on initialization', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const sheets = spreadsheet.getSheets();

  t.deepEqual(sheets.length, 1);
  t.deepEqual(sheets[0].getName(), 'Sheet1');

});

test('should handle insertSheet ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const newSheet: Sheet = spreadsheet.insertSheet();

  t.deepEqual(spreadsheet.getSheets().length, 2);
  t.deepEqual(newSheet.getName(), 'Sheet2');

});

test('should handle getId ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  t.true(typeof spreadsheet.getId() === 'number');

});

test('should throw error on getActiveCell', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const expected = Errors.NOT_IMPLEMENTED_YET;

  const actual = t.throws(() => {
    spreadsheet.getActiveCell();
  }, Error);

  t.is(actual.message, expected);

});

test.todo('should delete second column');
