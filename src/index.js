const FileManager = require('./services/FileManager');
const SpreadsheetReader = require('./services/SpreadsheetReader');

const filename = '_planilhas/duplas.xlsx';

const result = new SpreadsheetReader(
    filename,
    JSON.parse(FileManager.getFile(`${__dirname}\\config.json`))
);

// eslint-disable-next-line
console.log(result);
