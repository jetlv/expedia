var fs = require('fs');

var ew = require('node-xlsx');




// /* original data */
var data = [[new Date('1990/05/11'),new Date('1990/05/10'),3],[true, false, null, "sheetjs"],["foo","bar","0.3"], ["baz", null, "qux"]]
// var ws_name = "SheetJS";
var sheet = {name : 'test', data : data};

var buffer = ew.build([sheet]);

fs.writeFileSync('test.xlsx', buffer);

var sheets = ew.parse(fs.readFileSync('test.xlsx'));
console.log(new Date(new Date().getFullYear() + 70, 0 ,0, 0, 0,0,0));
sheets[0].data[0][0] = new Date(parseFloat(sheets[0].data[0][0]) * 24 * 60 * 60 * 1000 - (new Date(new Date().getFullYear() + 70, 0 ,0, 0, 0,0,0).getMilliseconds() - new Date().getMilliseconds())); 
// sheets[0].data[0][0] = new Date(0); 

var buffer = ew.build(sheets);

fs.writeFileSync('test.xlsx', buffer);
// /* require XLSX */
// var XLSX = require('./node_modules/node-xlsx/node_modules/xlsx');

// /* set up workbook objects -- some of these will not be required in the future */
// var wb = {}
// wb.Sheets = {};
// wb.Props = {};
// wb.SSF = {};
// wb.SheetNames = [];

// /* create worksheet: */
// var ws = {}

// /* the range object is used to keep track of the range of the sheet */
// var range = {s: {c:0, r:0}, e: {c:0, r:0 }};

// /* Iterate through each element in the structure */
// for(var R = 0; R != data.length; ++R) {
//   if(range.e.r < R) range.e.r = R;
//   for(var C = 0; C != data[R].length; ++C) {
//     if(range.e.c < C) range.e.c = C;

//     /* create cell object: .v is the actual data */
//     var cell = { v: data[R][C] };
//     if(cell.v == null) continue;

//     /* create the correct cell reference */
//     var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

//     /* determine the cell type */
//     if(typeof cell.v === 'number') cell.t = 'n';
//     else if(typeof cell.v === 'boolean') cell.t = 'b';
//     else cell.t = 's';

//     /* add to structure */
//     ws[cell_ref] = cell;
//   }
// }


// ws['!ref'] = XLSX.utils.encode_range(range);

// console.log(ws);

// /* add worksheet to workbook */
// wb.SheetNames.push(ws_name);
// wb.Sheets[ws_name] = ws;

// /* write file */
// XLSX.writeFile(wb, 'test.xlsx');