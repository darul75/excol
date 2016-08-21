import { test } from 'ava';
import { Sheet, Spreadsheet, SpreadsheetConfig } from 'excol';
import { User } from 'excol';

const email: string = 'test@com.net';

const cfg: SpreadsheetConfig = {
  name: 'test lib'
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

test('should handle deleteSheet ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  const newSheet: Sheet = spreadsheet.insertSheet();

  const firstSheet: Sheet = spreadsheet.getSheets()[0];

  spreadsheet.deleteSheet(firstSheet);

  t.deepEqual(spreadsheet.getSheets().length, 1);
  t.deepEqual(spreadsheet.getSheets()[0].getName(), 'Sheet2');

});

test('should handle getId ', t => {

  const spreadsheet = new Spreadsheet(cfg);

  t.true(typeof spreadsheet.getId() === 'number');

});

test.todo('should delete second column');
