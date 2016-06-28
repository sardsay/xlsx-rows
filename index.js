var xlsx = require('xlsx'),
    range = require('r...e');

//
// Regular expression to get the cell header
//
var letters    = range('A', 'Z').toArray(),
    cellHeader = /^([A-Za-z]+)\d{1,}/,
    isCell     = /^[A-Za-z]+\d{1,}/;
var letters2 = range('A', 'Z').toArray();
    letters.forEach(function (lett1) {
      letters2.forEach(function (lett2) {
          letters.push(lett1 + lett2);
        })
    })
//
// ### function xlsxRows (options)
// #### @options {string|Object} Options for reading rows
// ####   file      {string} Name of the file to read.
// ####   sheetname {string} Name of the file to read.
// ####   format    {string} Format of the file to read.
// Reads the rows from the XLSX.
//
module.exports = function (options, sheetIndex) {
  var file   = options,
      rows   = [],
      row    = [],
      start  = /^A\d{1,}/,
      format = 'v',
      workbook,
      sheetname,
      sheet;

  if (typeof options !== 'string') {
    file      = options.file;
    sheetname = options.sheetname;
    format    = options.format || 'w';
  }

  workbook  = xlsx.readFile(file);
  sheetname = sheetname || workbook.SheetNames[sheetIndex];
  sheet     = workbook.Sheets[sheetname];

  if (!sheet) {
    throw new Error('No sheet with name: ' + sheetname);
  }

  //
  // Pushes the next row onto the `rows`
  //
  var rawObj = []
  function pushRow() {
    //
    // Fill the row since we prefer the empty string
    // to a value of undefined.
    //
    rows.push(rawObj);
    rawObj = []
    ckRow = true;
  }
  var ckRow = true;// check end row : true = start row
  var currentRow
  Object.keys(sheet).forEach(function (cell) {
    if (!isCell.test(cell)) {
      return;
    }
    var currentLetter = getLetter(cell);
    var strCell = cell;
    var currentCell = strCell.substr(currentLetter.length)
    // check first row parameter
      if (ckRow) {
        currentRow = currentCell
        ckRow = false;
      }else{
        // check data is new row push data in array
        if (currentRow != currentCell) {
          pushRow()
        }
      }
      rawObj[getLetter(cell)] = {value :sheet[cell].v,column:cell,format :sheet[cell].w}
  });
  pushRow();
  return rows;
};

//
// ### function rowIndex (cell)
// Returns the row index for the cell
//
function rowIndex (cell) {
  var header = cellHeader.exec(cell),
      length;

  if (!header) {
    throw new Error('Bad cell header for: ' + cell);
  }

  //
  // TODO: Actually do something with the length
  // to support multi-character headers.
  //
  header = header[1];
  length = header.length;
  return letters.indexOf(header);
}
function getLetter(cell){
  var letterHeader = cellHeader.exec(cell);
  return letterHeader[1];
}