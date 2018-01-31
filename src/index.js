const fs = require('fs')
const XLSX = require('xlsx');
 
const sappData = require('../SAPP-data');


function convertMeasures(data) {
   return data.measures.reduce((accRoot,mes) => {
       return mes.history.reduce( (acc, hist) => {
            const key = hist.date.slice(0,10);
            acc[key] = Object.assign({}, acc[key], { [mes.metric]: hist.value});
            return acc;
        }, accRoot);
    }, {});
}

function convertDataAsSheetAoa(data) {
    const cols = ['date', 'bugs', 'code_smells', 'vulnerabilities', 'ncloc', 'duplicated_lines_density'];
    const lines = Object.entries(data).map( ([date, merics]) => {
        const line = [date];
        Object.entries(merics).forEach( ([key, val]) => {
            const idx = cols.indexOf(key);
            if (idx>0) {
                line[idx]=val;
            }
        });
        return line;
    });

    return [ cols, ...lines]
}
function convertWbSheet(data) {
    const lines = convertDataAsSheetAoa(data);
    const ws = XLSX.utils.aoa_to_sheet(lines);
    // columns
    ws["!cols"] = ws["!cols"] || [];
    [15, 10, 10, 10].forEach((wch, idx)=> {
        ws["!cols"][idx] = Object.assign({}, ws["!cols"][idx], {wch});
    });
    // cells
    console.log(ws['A1']);
    return ws;
}

function  main() {
    const dataApp =  convertMeasures(sappData);
    // https://github.com/SheetJS/js-xlsx/blob/master/tests/write.js
    const wb = XLSX.utils.book_new();
    const ws = convertWbSheet(  dataApp);
    XLSX.utils.book_append_sheet(wb, ws, 'appName');
    //console.log(JSON.stringify(dataApp));
    XLSX.writeFile(wb, 'out.xlsb');
}

main();