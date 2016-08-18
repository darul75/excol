# excol

| Detail   |      A1      |  Matrix |
|----------|:-------------:|------:|
| A1 	   |  A1 		   | '1;1' |
| A1 through B5 |    A1:B5 | [ '1,1' , '5;2' ] |
| A multiple-area selection | C5:D9,G9:H16 | [ [ '5,3','9,4' ], [ '9,7','16,8' ] ]  |
| Column A | A:A |  'A' |
| Row 1 | 1:1 | '1' |
| Rows 1 through 5 | 1:5 | '1-5' |
| Columns A through C | A:C | 'A-C' | 
| Rows 1, 3, and 8 | 1:1,3:3,8:8 | ['1', '3', '8'] |
| Columns A, C, and F | "A:A,C:C,F:F" | ['A', 'C', 'F'] |

##Methods

### Column encoders

#### colEncod

Encode column name in alpha notation to column number.

**Usage**
 
*colEncod('A')*

**Result**
```json
1
```

#### colDecod

Decode column number to alpha notation.

**Usage**
*colEncod(1)*

**Result**

```json
A
```

### Grid

#### Initialization

Initiate a new grid with specified range dimensions.

**Usage**
new Grid('A2:B3');
new Grid([ '2,1' , '5;2' ]);

#### Values

Rectangular grid of values for this grid.

**Usage**
Grid.values()

**Result**
```json
[ 
  [
    {"value": 0}, {"value": 0}
  ], 
  [
    {"value": 0}, {"value": 0}
  ] 
]
```

#### GetRange(row, column)

Return range with the top left cell at the given coordinates.

**Usage**
Grid.getRange(2, 3);
  
getValue('A3') => ''
getValue('B2') => ''
getValue('C2') => Error
getValue('A1') => Error
getRange(1,1)
getRange(row, column)	range with the top left cell at the given coordinates.
getRange(row, column, numRows)	range with the top left cell at the given coordinates, and with the given number of rows.
getRange(row, column, numRows, numColumns)	range with the top left cell at the given coordinates with the given number of rows and columns.
getRange(a1Notation)



