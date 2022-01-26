const XLSX = require('xlsx');
const REG = require('../utils/Regex');
const CHAR = require('../utils/Char');

module.exports = class SpreadsheetReader {
    constructor(filename, configs = {}) {
        this.filename = filename;
        this.configs = configs;
        this.finalContent = false;
        this.init();
    }

    init() {
        try {
            this.readSpreadsheet();
            this.organizeColumns();
            this.setData();
            this.prepareJSON();
            return this;
        } catch {
            return false;
        }
    }

    readSpreadsheet() {
        this.pages = {};
        const workbook = XLSX.readFile(this.filename);

        Object.entries(workbook.Sheets).forEach(([index, page]) => {
            if (
                typeof this.configs.pages !== 'undefined' &&
                this.configs.pages.length &&
                this.configs.pages.indexOf(index.toLowerCase()) < 0
            )
                return;

            const columns = [];
            const datas = [];
            let previous = false;

            Object.entries(page).forEach(([key, value]) => {
                if (typeof value.w === 'undefined') return;
                let data = value.w.toString();
                const letter = key.replace(REG.onlyLetters, '');

                if (columns.length < this.configs.columns.length) {
                    data = data.toLowerCase();

                    const found = this.configs.columns.find(
                        (column) => column.searchIndex === data
                    );

                    if (typeof found === 'object') {
                        if (previous) {
                            previous.letter.end = CHAR.previousChar(letter);
                        }

                        previous = {
                            index: found.searchIndex,
                            letter: {
                                start: letter,
                                end: 'Z'
                            },
                            newIndexValue: found.newIndexValue,
                            data: [],
                            subColumns: found.subColumns || false
                        };

                        columns.push(previous);

                        return;
                    }

                    previous = false;
                } else {
                    const found = columns.find(
                        (column) => letter >= column.letter.start && letter <= column.letter.end
                    );
                    if (typeof found === 'object') {
                        if (found.subColumns) {
                            if (typeof found.subColumns[data] !== 'undefined') {
                                found.subColumns[data] = Object.assign(found.subColumns[data], {
                                    index: found.subColumns[data].horizontal
                                        ? CHAR.nextChar(key)
                                        : key,
                                    data: []
                                });
                            } else {
                                datas.push({ key, value: data });
                            }
                        } else {
                            found.data[key.replace(REG.onlyNumbers, '')] = data;
                        }
                    }
                }
            });

            this.pages[index] = { columns, datas };
        });
    }

    organizeColumns() {
        Object.entries(this.pages).forEach((page) => {
            page[1].columns.forEach((column) => {
                Object.entries(column.subColumns).sort((a) => (a[1].horizontal ? true : -1));
            });
        });
    }

    setData() {
        Object.entries(this.pages).forEach((page) => {
            page[1].datas.forEach((data) => {
                const letter = data.key.replace(REG.onlyLetters, '');
                const number = data.key.replace(REG.onlyNumbers, '');

                const columnFound = page[1].columns.find(
                    (column) => letter >= column.letter.start && letter <= column.letter.end
                );

                if (columnFound.subColumns) {
                    const subColumnFound = Object.entries(columnFound.subColumns).find(
                        (subColumn) =>
                            typeof subColumn[1].index !== 'undefined' &&
                            ((subColumn[1].horizontal && subColumn[1].index === data.key) ||
                                (!subColumn[1].horizontal &&
                                    subColumn[1].index.replace(REG.onlyLetters, '') === letter &&
                                    number > subColumn[1].index.replace(REG.onlyNumbers, '')))
                    );
                    if (subColumnFound) subColumnFound[1].data[number] = data.value;
                    return;
                }
                columnFound.data[number] = data.value;
            });
        });
    }

    prepareJSON() {
        const json = {};
        Object.entries(this.pages).forEach((page) => {
            json[page[0]] = {};
            page[1].columns.forEach((column) => {
                if (column.data.length) {
                    json[page[0]][column.newIndexValue] = column.data;
                } else {
                    json[page[0]][column.newIndexValue] = {};
                    Object.entries(column.subColumns).forEach((subColumn) => {
                        json[page[0]][column.newIndexValue][
                            subColumn[1].newIndexValue || subColumn[0]
                        ] = subColumn[1].data;
                    });
                }
            });
        });

        this.finalContent = json;
    }
};
