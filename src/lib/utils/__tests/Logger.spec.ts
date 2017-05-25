import { test } from 'ava';
import { Logger } from "excol";

test.beforeEach(t => {
  Logger.clear();
});

test('should handle clear log', t => {

  Logger.log("test");
  Logger.clear();

  const log = Logger.getLog();

  t.deepEqual('', log);

});

test('should handle log of data', t => {

  Logger.log("test");

  const log = Logger.getLog();

  t.deepEqual("test", log);

});

test.todo('should handle log of data and append it')

  /*, t => {

  Logger.log("test");
  Logger.log(" ok !");

  const log = Logger.getLog();

  //t.deepEqual("test\r\n ok !\r\n", log);

});*/

test('should handle log of data with format', t => {

  Logger.log('You are a member of %s Google Groups.', 1);

  const log = Logger.getLog();

  t.deepEqual('You are a member of 1 Google Groups.', log);

});

test('should handle log of data with format', t => {

  Logger.log('You are a member of %s %s Google Groups.', 1, 'dude');

  const log = Logger.getLog();

  t.deepEqual('You are a member of 1 dude Google Groups.', log);

});
