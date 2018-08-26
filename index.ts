#!/usr/bin/env node

import commander from 'commander';
import csvStringify from 'csv-stringify';
import fs from 'fs';
import path from 'path';
import { promisify } from 'util';
import xlsx from 'xlsx';

import { IConvertOptions } from './src/types/covnert_options.interface';

const csvStringifyAsync = promisify<csvStringify.Input, csvStringify.Options, string>(csvStringify);
const writeFileAsync = promisify(fs.writeFile);

export async function convert(
    excelPath: string,
    partialOptions: Partial<IConvertOptions> = {},
): Promise<string | boolean> {
    try {
        const options: IConvertOptions = formOptions(excelPath, partialOptions);
        const items: any[] = await excelToJson(excelPath);

        const csv = await csvStringifyAsync(items, { header: true });

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

async function excelToJson(excelPath: string): Promise<any[]> {
    const workbook: xlsx.WorkBook = xlsx.readFile(excelPath);
    const sheetName: string = workbook.SheetNames[0];
    const sheet: xlsx.WorkSheet = workbook.Sheets[sheetName];

    const items = xlsx.utils.sheet_to_json(sheet);

    return items;
}

function main(): void {
    if (require.main !== module) { // required as module
        return;
    }

    commander
        .version('0.1.0', '-v, --version')
        .arguments('<excel_file>')
        .option('-o, --output <csv_file>', 'Output CSV file path')
        .action((excelFile?: string) => {
            if (excelFile) {
                convert(excelFile, {
                    csvPath: commander.output,
                    writeCsv: true,
                });
            }
        })
        .parse(process.argv);

    return;
}

main();
