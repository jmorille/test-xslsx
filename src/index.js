const fs = require('fs')
const XLSX = require('xlsx');


// https://github.com/SheetJS/js-xlsx/blob/master/tests/write.js
var wb = XLSX.utils.book_new();

var ws_name = "SheetJS";

/* make worksheet */
var ws_data = [
    [ "S", "h", "e", "e", "t", "J", "S" ],
    [  1 ,  2 ,  3 ,  4 ,  5 ],
    [  1 ,  2 ,  3 ,  4 ,  5 ]
];
const ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the sheet name to the list */

XLSX.utils.book_append_sheet(wb, ws, ws_name);


ws['C1'].f = "A1+B1";
ws['C2'] = {t:'n', f:"A1+B1"};

XLSX.writeFile(wb, 'out.xlsb');