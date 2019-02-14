#!/usr/bin/env node

import commander from 'commander';
import fs from 'fs';
import path from 'path';
import { promisify } from 'util';
import xlsx from 'xlsx';

import { IConvertOptions } from './src/types/covnert_options.interface';
import { IExcelParserOptions } from './src/types/excel_parser_options.interface';

const writeFileAsync = promisify(fs.writeFile);

export async function convert(
    excelPath: string,
    partialOptions: Partial<IConvertOptions> = {},
): Promise<string | boolean> {
    try {
        const options: IConvertOptions = formOptions(excelPath, partialOptions);

        const csv = excelToJson(excelPath, {
                sheetIndex: options.sheetIndex,
                sheetName: options.sheetName,
            });

        if (options.writeCsv) {
            await writeFileAsync(options.csvPath, csv);
            return true;
        }

        return csv;
    } catch (error) {
        throw error;
    }
}

function formOptions(excelPath: string, partialOptions: Partial<IConvertOptions>): IConvertOptions {
    const excelFilename: string = path.parse(excelPath).name;
    const options: IConvertOptions = {
        ...{
            csvPath: `${excelFilename}.csv`,
            writeCsv: false,
        },
        ...partialOptions,
    };

    return options;
}

function excelToJson(excelPath: string, options: Partial<IExcelParserOptions> = {}): string {
    const workbook: xlsx.WorkBook = xlsx.readFile(excelPath);
    const sheetName: string = getExcelSheetName(workbook.SheetNames, options);
    const sheet: xlsx.WorkSheet = workbook.Sheets[sheetName];

    return xlsx.utils.sheet_to_csv(sheet);
}

function getExcelSheetName(sheetNames: string[], options: Partial<IExcelParserOptions>): string {
    if (options.sheetIndex && options.sheetIndex < sheetNames.length) {
        return sheetNames[options.sheetIndex];
    } else if (options.sheetName && sheetNames.indexOf(options.sheetName) > -1) {
        return options.sheetName;
    } else {
        return sheetNames[0];
    }
}

function main(): void {
    if (require.main !== module) { // required as module
        return;
    }

    commander
        .version('0.1.0', '-v, --version')
        .arguments('<excel_file>')
        .option('-n, --sheet-index <sheet_name>', 'Sheet index to convert (defaults to 0)')
        .option('-s, --sheet-name <sheet_name>', 'Sheet name to convert')
        .option('-o, --output <csv_file>', 'Output CSV file path (defaults to input file name)')
        .action((excelFile?: string) => {
            if (excelFile) {
                const options: Partial<IConvertOptions> = { writeCsv: true };
                if (commander.output) {
                    options.csvPath = commander.output;
                }

                if (Number(commander.sheetIndex)) {
                    options.sheetIndex = Number(commander.sheetIndex);
                } else if (commander.sheetName) {
                    options.sheetName = commander.sheetName;
                }

                convert(excelFile, options);
            }
        })
        .parse(process.argv);

    return;
}

main();

export { IConvertOptions };
