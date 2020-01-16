import program from 'commander';
import excel from 'exceljs';
import fs from 'fs';
import path from 'path';

interface Options {
    column: number;
    row: number;
    cut: boolean;
    sheet?: number;
}

const dirname = ('pkg' in process) ?
    path.dirname((process as NodeJS.Process).execPath) :
    __dirname;

program.option(
    '--column <index>',
    'index of column to begin processing (starts at 1)',
    (column) => Math.max(1, parseInt(column, 10)) || 1,
    1
);

program.option(
    '--row <index>',
    'index of row to begin processing (starts at 1)',
    (row) => Math.max(1, parseInt(row, 10)) || 1,
    1
);

program.option(
    '--sheet <index>',
    'index of sheet to process (starts at 1)',
    (row) => Math.max(1, parseInt(row, 10)),
);

program.option(
    '--cut',
    'cut cells with null values',
    false
);

program.parse(process.argv);

const options: Options = {
    column: 1,
    row: 1,
    cut: false,
    ...program.opts()
};

async function app () {
    const filename = program.args[0] && path.join(dirname, program.args[0]);

    if (!filename) {
        throw `File not specified`;
    }

    if (!fs.existsSync(filename)) {
        throw `File ${filename} not found`;
    }

    const workbook = await new excel.Workbook().xlsx.readFile(filename);

    workbook.eachSheet(processSheet);

    let outputName = program.args[1] && path.join(dirname, program.args[1]) || filename;

    await workbook.csv.writeFile(outputName);
}

function processSheet(worksheet: excel.Worksheet, worksheetId: number): void {
    if (options.sheet && options.sheet !== worksheetId) {
        return;
    }

    worksheet.eachRow(processRow);
}

function processRow(row: excel.Row, rowId: number): void {
    if (rowId <= options.row) {
        return;
    }

    const values: excel.CellValue[] = [];

    for (let i = options.column + 1; i <= row.cellCount; i++) {
        values.push(row.getCell(i).value)
    }

    const filteredValues = options.cut ?
        filterDuplicates(values) :
        nullDuplicates(values);

    row.splice(options.column + 1, values.length, ...filteredValues);

    row.commit();
}

function nullDuplicates(values: excel.CellValue[]) {
    return values.map((value, index) =>
        (values.indexOf(value) === index) ?
            value :
            null
    );
}

function filterDuplicates(values: excel.CellValue[]) {
    return values.filter((value, index) =>
        (values.indexOf(value) === index)
    );
}

app().catch((e) => {
    console.error(`Error: ${e}`);
    process.exit(-1);
});
