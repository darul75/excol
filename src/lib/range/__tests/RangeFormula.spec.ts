import { test } from 'ava';
import { Sheet, SheetConfig } from 'excol';
import { Errors } from 'excol';

const DIMENSION = 20;

const cfg: SheetConfig = {
  id: 0,
  name: 'test lib',
  numRows: DIMENSION,
  numColumns: DIMENSION
};

test('should handle formula for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});
  const for1 = '=for1';
  const for2 = '=for2';

  range.setFormula(for1);

  const form = range.getFormula();
  const forms = range.getFormulas();

  const expected = [
    [for1, for1],
    [for1, for1]
  ];

  t.deepEqual(forms, expected);
  t.is(form, for1);

  range.setFormulas([
    [for2, for2],
    [for2, for2]
  ]);

  const expected2 = [
    [for2, for2],
    [for2, for2]
  ];

  const form2 = range.getFormula();
  const forms2 = range.getFormulas();

  t.deepEqual(forms2, expected2);
  t.is(form2, for2);

});

test('should handle formula R1C1 for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});
  const for1 = '=for1';
  const for2 = '=for2';

  range.setFormulaR1C1(for1);

  const form = range.getFormulaR1C1();
  const forms = range.getFormulasR1C1();

  const expected = [
    [for1, for1],
    [for1, for1]
  ];

  t.deepEqual(forms, expected);
  t.is(form, for1);

  range.setFormulasR1C1([
    [for2, for2],
    [for2, for2]
  ]);

  const expected2 = [
    [for2, for2],
    [for2, for2]
  ];

  const form2 = range.getFormulaR1C1();
  const forms2 = range.getFormulasR1C1();

  t.deepEqual(forms2, expected2);
  t.is(form2, for2);

});

//notes
test('should handle note(s) for A2:B3', t => {

  const grid = new Sheet(cfg);
  const range = grid.getRange({A1: 'A2:B3'});
  const n1 = 'n1';
  const n2 = 'n2';

  range.setNote(n1);

  const n = range.getNote();
  const ns = range.getNotes();

  const expected = [
    [n1, n1],
    [n1, n1]
  ];

  t.deepEqual(ns, expected);
  t.is(n, n1);

  range.setFormulasR1C1([
    [n2, n2],
    [n2, n2]
  ]);

  const expected2 = [
    [n2, n2],
    [n2, n2]
  ];

  const n3 = range.getFormulaR1C1();
  const ns2 = range.getFormulasR1C1();

  t.deepEqual(ns2, expected2);
  t.is(n3, n2);

});
