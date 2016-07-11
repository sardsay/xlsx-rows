# xlsx-rows

Parses an *.xlsx file into rows

### Usage

All you have to do is pass it an Excel (`*.xlsx`) file and you get rows of information:

``` js
  var xlsxRows = require('xlsx-rows');
  var sheetIndex = 0;
  var rows = xlsxRows('my-workbook.xlsx', sheetIndex);
  console.dir(rows); // yay rows of things!
```
### Example output
```
	[ [ A: { value: 'test1', column: 'A1', format: '14/01/2010' },
    B: { value: 'test2', column: 'B1', format: '14/01/2010' },
    C: { value: 'test3', column: 'C1', format: '14/01/2010' },
    D: { value: 'test4', column: 'D1', format: '14/01/2010' },
    E: { value: 'test5', column: 'E1', format: '14/01/2010' },
    F: { value: 'test6', column: 'F1', format: '14/01/2010' } ]]
```

#### Author: [Charlie Robbins](http://github.com/indexzero)
#### License: MIT
