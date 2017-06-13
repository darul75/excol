# excol

[![Build Status](https://travis-ci.org/darul75/excol.svg?branch=master)](https://travis-ci.org/darul75/excol) [![codecov](https://codecov.io/gh/darul75/excol/branch/master/graph/badge.svg)](https://codecov.io/gh/darul75/excol)

Library to mock excel(-like) solutions in Javascript.

## Why

You may (should) write your own tests before deploying your code to solution like **Google sheet** or **Office 365**.

## Idea

In case of google app script, you can use their own editor to play with simple scripts but for bigger project it will become very complex to maintain.

What you may need is a tool to work locally and push back your changes to Google.

- Starter project https://github.com/danthareja/node-google-apps-script 
- Mock everything with **Excol** 

## How to use

Google app script language is Javascript, a kind of, not ECMA 5 but an old one with most the basic necessary native features.

With starter kit provided before, you can named your files with .js extensions and play with it in a NodeJS environment (Tests).

Using a CJS manner, you just need to enhance your files a little bit this way.

*MyFunctions.js*

```javascript

/**
 * NodeJS part for tests
 */
if (typeof exports === 'object' && typeof module !== 'undefined') {
  var excol = require('excol');
  var SpreadsheetApp = excol.SpreadsheetApp;
}

function read() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var range = activeSheet.getRange("A1:B2");
}

function write() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var range2 = activeSheet.getRange(1,1,2,2);
    range2.setValues([
    [1,2],
    [3,4]
  ]);
}

/**
 * NodeJS part for tests
 **/
if (typeof exports === 'object' && typeof module !== 'undefined') {
  module.exports = {
    read: read,
    write: write
  };
  //
  // module.exports = read; 
}
```

*Main.js*

```javascript
/**
 * NodeJS part for tests
 */
if (typeof exports === 'object' && typeof module !== 'undefined') {
  var MyFunctions = require('MyFunctions');
}

function main() {
  MyFunctions.read();
  MyFunctions.write();
}

```
