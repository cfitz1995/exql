#!/usr/bin/env node

'use strict';

// ------------------------------- NODE MODULES -------------------------------

const xlsx = require('xlsx');
const fs = require('mz/fs');

const argv = require('yargs')
    .command('run', 'Generate SQL', yargs =>  yargs
        .usage('Usage: exql run [options]')
        .example('exql -i C:/Documents/myExcelFile.xlsx -o C:/Documents/myOutputFile.sql', 'Process an Excel file and generate insert statements for each sheet')
        .alias('i', 'input')
        .nargs('i', 1)
        .describe('i', 'Path of the Excel file')
        .alias('o', 'output')
        .nargs('o', 1)
        .describe('o', 'Path to output the SQL file. Output will only be printed to the console if ommitted')
        .alias('q', 'quiet')
        .describe('q', 'Suppress output to the console')
        .alias('f', 'force')
        .describe('f', 'Overwrite the output file if it already exists')
        .demandOption(['i'])
        .help('h')
        .alias('h', 'help')
        .argv)
    .demandCommand(1, 'You must provide the input file path')
    .argv;

// ------------------------------ CUSTOM MODULES ------------------------------

// -------------------------------- VARIABLES ---------------------------------

const argsFilePath = argv.i || argv.input;
const outputFilePath = argv.o || argv.output;
const suppressOutput = argv.q || argv.quiet || false;
const forceOverwrite = argv.f || argv.false || false;
const AcharCode = 65;

// ----------------------------- FILE DEFINITION ------------------------------

/**
 * Split a cell address into row and column
 *
 * @param {string!} cell Cell address
 *
 * @returns {Object} A JSON object containing row and column
 */
const getColumnAndRow = cell => {
    const column = (/\D+/g).exec(cell)[0];
    const row = cell.substring(column.length);

    return {
        column,
        row
    };
};

/**
 * Get all populated cells on a row
 *
 * @param {Object!} worksheet Parsed Excel worksheet
 * @param {Number!} columnsCount Number of columns
 * @param {Number!} rowIndex Index of the row to get
 * @param {Number!} wrapInQuotes If true, wrap results in quotes
 * @param {Array<string>!} numericColumns An array of column headers which are not to be wrapped in quotes
 *
 * @returns {Array<Object>} An array with each cell as a JSON object
 */
const getRow = (worksheet, columnsCount, rowIndex, wrapInQuotes, numericColumns) => {
    const result = [];

    for (let charCode = 0; charCode < columnsCount; charCode++) {
        const cell = worksheet[`${String.fromCharCode(AcharCode + charCode)}${rowIndex}`];

        result[charCode] = cell ? cell.v : 'NULL';

        if (wrapInQuotes === true && result[charCode] !== 'NULL' && !numericColumns[charCode]) {
            result[charCode] = `'${result[charCode]}'`;
        }
    }

    return result;
};

/**
 * Get the max row populated in the excel sheet
 *
 * @param {Object!} worksheet Parsed excel worksheet
 *
 * @returns {Number} Maximum row number
 */
const getMaxRowAndColumn = worksheet => {
    const maxCell = worksheet['!ref'].split(':')[1];

    return getColumnAndRow(maxCell);
};

/**
 * Reads in an Excel file and produces insert statements for each sheet
 *
 * @param {string!} excelFilePath Full path of the Excel file to process
 * @param {string?} outputPath Full path of where to write the output
 * @param {boolean?} quiet True to suppress console output
 * @param {boolean?} force True to overwrite the output file if it already exists
 *
 * @returns {void}
 */
const processWorkbook = (excelFilePath, outputPath, quiet = false, force = false) => {
    fs.exists(excelFilePath)
        .then(result => {
            if (!result) {
                throw new Error(`${excelFilePath} is not a valid file path`);
            }

            return outputPath ? fs.exists(outputPath) : Promise.resolve(false);
        })
        .then(result => {
            if (result && !force) {
                throw new Error(`${outputPath} already exists. Provide the force (f) flag to overwrite`);
            }

            const workbook = xlsx.readFile(excelFilePath);

            let outputStr = '';

            for (let sheetIndex = 0; sheetIndex < workbook.SheetNames.length; sheetIndex++) {
                const sheetName = workbook.SheetNames[sheetIndex];

                let commentLength = 100;
                const comTitleLength = (commentLength / 2) - 1;

                commentLength += sheetName.length;

                outputStr += `${'='.repeat(commentLength)}\n`;
                outputStr += `${'='.repeat(comTitleLength)} ${sheetName} ${'='.repeat(comTitleLength)}\n`;
                outputStr += `${'='.repeat(commentLength)}\n`;

                if (!quiet) {
                    console.log(`Generating insert statements for ${sheetName}...`);
                    console.log('-------------------------------------------------------------------------------------------------');
                }

                const worksheet = workbook.Sheets[sheetName];

                if (!worksheet.A1) {
                    throw new Error(`${sheetName} is invalid. Please ensure the spreadsheet begins at A1, with column headers along row 1`);
                }

                const maxValues = getMaxRowAndColumn(worksheet);

                const maxRow = maxValues.row;
                const maxColumn = maxValues.column.charCodeAt(0);
                const columnsCount = maxColumn - AcharCode + 1;

                const headers = getRow(worksheet, columnsCount, 1);

                const numericColumns = [];

                for (let headerIndex = 0; headerIndex < headers.length; headerIndex++) {
                    if (headers[headerIndex].startsWith('#')) {
                        headers[headerIndex] = headers[headerIndex].substring(1);

                        numericColumns[headerIndex] = headers[headerIndex];
                    }
                }

                const results = [];

                for (let rowIndex = 2; rowIndex < maxRow; rowIndex++) {
                    const currentRow = getRow(worksheet, columnsCount, rowIndex, true, numericColumns);

                    results.push(currentRow);
                }

                const maxColumnLengths = [];

                for (let resultIndex = 0; resultIndex < results.length; resultIndex++) {
                    for (let columnIndex = 0; columnIndex < results[resultIndex].length; columnIndex++) {
                        if (maxColumnLengths[columnIndex]) {
                            if (results[resultIndex][columnIndex].length > maxColumnLengths[columnIndex]) {
                                maxColumnLengths[columnIndex] = results[resultIndex][columnIndex].length;
                            }
                        } else {
                            maxColumnLengths[columnIndex] = results[resultIndex][columnIndex].length;
                        }
                    }
                }

                const sqlTemplate = `INSERT INTO ${sheetName} (${headers.join(', ')}) VALUES (#values);`;

                for (let resultIndex = 0; resultIndex < results.length; resultIndex++) {
                    for (let columnIndex = 0; columnIndex < results[resultIndex].length; columnIndex++) {
                        results[resultIndex][columnIndex] += ' '.repeat(maxColumnLengths[columnIndex] - `${results[resultIndex][columnIndex]}`.length);
                    }

                    const output = sqlTemplate.replace('#values', results[resultIndex].join(', '));

                    if (!quiet) {
                        console.log(output);
                    }

                    outputStr += `${output}\n`;
                }

                outputStr += '\n';

                if (!quiet) {
                    console.log('-------------------------------------------------------------------------------------------------');
                }
            }

            if (outputPath) {
                fs.writeFile(outputPath, outputStr)
                    .then(() => {
                        console.log(`Output successfully written to ${outputPath}`);
                    })
                    .catch(err => console.log(`Failed to write output to ${outputPath} - ${err}`));
            } else if (quiet) {
                console.log('Nothing to output. Please provide an output path or remove the quiet flag');
            }
        })
        .catch(err => {
            console.log(err);
        });
};

module.exports = processWorkbook;

module.exports(argsFilePath, outputFilePath, suppressOutput, forceOverwrite);