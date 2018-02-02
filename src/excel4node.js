const xl = require('excel4node');

const sappData = require('../SAPP-data');

function parseHistValue(metric, value) {
    switch (metric) {
        case 'date':
            return new Date(value);

        default:
            return Number.parseInt(value);
    }
}



function convertMeasures(data) {
    return data.measures.reduce((accRoot,mes) => {
        return mes.history.reduce( (acc, hist) => {
            const key = hist.date.slice(0,10);
            if (!acc[key]) {
                acc[key] = {date: parseHistValue('date', key)};
            }
            acc[key] = Object.assign({}, acc[key], { [mes.metric]: parseHistValue(mes.metric, hist.value) });
            return acc;
        }, accRoot);
    }, {});
}

function fillWorksheetMetrics(ws, data, styles) {
    const cols = ['date', 'bugs', 'code_smells', 'vulnerabilities', 'ncloc', 'duplicated_lines_density'];
    // Headers
    cols.forEach( (colKey, colIdx) => {
        const cell = ws.cell(1, colIdx+1);
        cell.string(colKey);
        cell.style(styles.headerStyle);
    });

    // Values
    Object.entries(data).forEach( ([key, line], idx)=> {
       const rowIdx = idx+2;
       cols.forEach( (colKey, colIdx) => {
           const cell =ws.cell(rowIdx, colIdx+1);
//           console.log(colIdx, ' == ', colKey, '--->', xl.getExcelCellRef(rowIdx, colIdx+1));
           if (colKey == 'date') {
               cell.date(line[colKey]);
           } else {
               cell.number(line[colKey]);
           }

//           console.log(cell.excelRefs);
       });
    });
    ws.row(1).setHeight(60);
    ws.row(1).freeze();
   // ws.column(1).freeze();
   // ws.row(1).filter();
    return ws;
}

function createStyle(wb) {
    const headerStyle = wb.createStyle({
        font: {
            bold: true
        },
        alignment: {
            wrapText: true,
            horizontal: 'left',
            textRotation: 30

        }
    });
    return {headerStyle}
}

function  main() {
    const dataApp = convertMeasures(sappData);
    const wb = new xl.Workbook({
        jszip: {
            compression: 'DEFLATE'
        },
        defaultFont: {
            size: 12,
            name: 'Calibri'
        },
        dateFormat: 'dd/mm/yyyy'
    });
    const styles = createStyle(wb);

    const ws = wb.addWorksheet("sApp");
    fillWorksheetMetrics(ws, dataApp, styles);

    wb.write('ExcelFile.xlsx');
}

main();