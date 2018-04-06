'use strict';
const fs = require('fs');
const Excel = require('exceljs');


function convetStringAsArray(str) {
    if (!str) return undefined;
    return str.split(/,|\/|\r?\n|\s|\t/).filter(value => {
        return value && value.trim().length > 0;
    });
}

function readCellStr(ws, cell) {
    let val = ws.getCell(cell).value;
    if (val) {
        val = val.trim();
    }
    return val;
}

function readMaven(ws, row) {
    const groupId = readCellStr(ws, 'F' + row);
    const artifactId = readCellStr(ws, 'G' + row);
    if (groupId && artifactId) {
        return {maven: {groupId, artifactId}}
    }
}

function readServer(ws, row) {
    let lan = convetStringAsArray(ws.getCell('D' + row).value);
    let dmz = convetStringAsArray(ws.getCell('E' + row).value);
    if (lan || dmz) {
        return {lan, dmz};
    }
}

function readEnv(ws, row) {
    const envLabel = readCellStr(ws,'C' + row);
    switch (envLabel) {
        case "Production":
            return 'prod';
        case "Recette":
            return 'rec';
        case "Qualif":
            return 'qa';
    }
    return envLabel;
}


function readAppName(ws, row) {
    return readCellStr(ws,'B' + row);
}


function readExcelLine(ws, row) {
    const name = readAppName(ws, row);
    const env = readEnv(ws, row);
    // test read
    if (!env || (env === 'usine')) return undefined;
    // Read dats
    const maven = readMaven(ws, row);
    const servers = readServer(ws, row);
    // Line Structure
    let line = { [name]: {} };
    if (maven) {
        line[name].maven = maven;
    }
    line[name][env]= {servers};


    return line;
}

function mergeApp(apps, line) {
    return Object.entries(line).reduce((acc, [name, data]) => {
        if (data) {
            acc[name] = Object.assign({}, acc[name], data);
        }
        return acc;
    }, apps);
    return apps;
}

function readWorksheet(worksheet) {
    const maxRow = worksheet.rowCount;
    let apps = {};
    for (let row = 2; row <= maxRow; row++) {
        const line = readExcelLine(worksheet, row);
        apps = mergeApp(apps, line);
    }
    return apps;
}


function main(filename) {
    const workbook = new Excel.Workbook();
    workbook.xlsx.readFile(filename).then(() => {
        const worksheet = workbook.getWorksheet(1);
        const apps = readWorksheet(worksheet);
        console.log(JSON.stringify(apps, null, 2));
        fs.writeFile("dory.json", JSON.stringify(apps, undefined, 2), () => {
            console.log("File is write for ", Object.keys(apps));
        });
    });
}


main('app.xlsx');