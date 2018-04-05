'use strict';
const Excel = require('exceljs');



function readExcelLine(worksheet, row) {
    const line = {
            [worksheet.getCell('A' + row).value]: {
                [worksheet.getCell('B' + row).value]: {
                    server: worksheet.getCell('C' + row).value
                }
            }
        }
    return line;
}

function mergeApp(apps, line) {
    return Object.entries(line).reduce( (acc, [name, data])=> {
        acc[name] = Object.assign({}, acc[name] , data);
        return acc;
    }, apps );
    return apps;
}

function readWorksheet(worksheet) {
    const maxRow = worksheet.rowCount;
    let apps = {};
    for (let row = 2; row <= maxRow; row++) {
        const line = readExcelLine(worksheet, row);
        apps = mergeApp(apps, line);
        console.log(apps);
    }
}


function main(filename) {
    const workbook = new Excel.Workbook();
    workbook.xlsx.readFile(filename).then(() => {
        const worksheet = workbook.getWorksheet(1);
        readWorksheet(worksheet);
    });
}


main('app.xlsx');