# excel2csv

Convert Excel files to CSV

## Installation

To install locally:

```sh
npm install excel2csv
```

To install globally:

```sh
npm install -g excel2csv
```

## Standalone Use Case

To use as a standalone application, install globally. To get help

```sh
excel2csv -h
```

To convert a `xlsx` or `xls` file to `csv`:

```sh
excel2csv -o output.csv input.xlsx
```

If output filename is not provided via `-o` option, the same filename as the input will be used with `.csv` extension.

## Local Use Case

To use locally, install the package in the local directory. Then, the package can be imported as:

```js
const excel2csv = require('excel2csv');
```

The package includes `convert` function with the following arguments:

```js
excel2csv.convert(excelPath, options);
```

`excelPath` is a string path to the input Excel file. `options` object is optional and has the following format:

```js
options = {
    csvPath, // string path to the output CSV file
    writeCsv, // if true, the output will be written to a file, otherwise will be returned by the function
}
```

The `convert` function returns a promise. If the `writeCsv` option is `true`, the function returns a boolean promise, which is `true` if file is successfully written, and `false` otherwise. If `writeCsv` is set to `false` (default value), the `convert` function returns a string promise containing the CSV output.
