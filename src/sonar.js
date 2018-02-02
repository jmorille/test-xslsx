const Excel = require('exceljs');

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


function convertWbSheet(worksheet,data) {
    worksheet.views = [
        {state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2'}
    ];

    const colStyle = {
        border: {
            top: {style:'thin'},
            left: {style:'thin'},
            bottom: {style:'thin'},
            right: {style:'thin'}
        }
    };
    const colStyleEnd = Object.assign({}, colStyle );
    colStyleEnd.border.right.style='double';
    worksheet.columns = [
        { header: 'Date', key: 'date', style:colStyle , width: 10, outlineLevel: 1 },
        { header: 'Bugs', key: 'bugs', style:colStyle , width: 10 },
        { header: 'code_smells', key: 'code_smells', style:colStyle , width: 10  },
        { header: 'Vulnerabilities', key: 'vulnerabilities', style:colStyle , width: 10  },
        { header: 'Duplication', key: 'duplicated_lines_density', style:colStyle , width: 10  },
        { header: 'Lignes de code', key: 'ncloc', style:colStyleEnd , width: 10    }
    ];
    Object.entries(data).forEach( ([key, line])=> {
        worksheet.addRow(line);
    });
    // format
    const headers = worksheet.getRow(1);
    headers.height = 60;

    headers.eachCell((cell, colNumber)=> {
        cell.font = { bold: true, color: { argb: '00000000'} };
        cell.fill = {
            type: 'gradient',
            gradient: 'angle',
            degree: 0,
            stops: [
                {position:0, color:{argb:'FF0000FF'}},
                {position:0.5, color:{argb:'FFFFFFFF'}},
                {position:1, color:{argb:'FF0000FF'}}
            ]
        };
        cell.alignment = { horizontal: 'left', textRotation: 30, wrapText: true  };
        //console.log('Cell ' + colNumber + ' = ' + cell.value);
    });

    //worksheet.autoFilter = 'A1:F1';

    // Finished adding data. Commit the worksheet
    //worksheet.commit();
}


function  main() {
    const dataApp = convertMeasures(sappData);

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('sApp', {
        pageSetup:{paperSize: 9, orientation:'portrait'},
        properties:{tabColor:{argb:'FF00FF00'}}
    });


    const ws = convertWbSheet( worksheet,  dataApp);
    // write
    const filename = "sonar.xlsx";
    workbook.xlsx.writeFile(filename)
        .then(() => {
            // done
            console.log("Write ", filename);
        });

}

main();