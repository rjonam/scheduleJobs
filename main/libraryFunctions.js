
const xlsx = require('xlsx');
const excelJS = require('exceljs');


class LibraryFunctions {

    getRowsBySheetName(sheetName, filePath) {
        let data;
        try {
            if (filePath.toString().endsWith('.xlsx')) {
                const workbook = xlsx.readFile(filePath);
                data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

            }
        }
        catch (e) {
            throw new Error('Exception in reading all rows from excel' + e);
        }
        return data;
    }

    async updateSheetWithCurrentValue(filePath,currentValue){
        let tempCompCount = 0;
        const newWorkbook = new excelJS.Workbook();
        await newWorkbook.xlsx.readFile(filePath);
        const newworksheet = newWorkbook.getWorksheet('Main');
        const rowCount = newworksheet.rowCount;
        const columnCount = newworksheet.actualColumnCount +1 ;
        newworksheet.getColumn(columnCount).header = new Date().
        toLocaleString('en-AU', { timeZone: 'Australia/Melbourne' });
        for(let i = 2;i <=rowCount; i++){
            if((newworksheet.getCell(i,1).value) != null){
                const cell = newworksheet.getCell(i,columnCount);
                cell.value = currentValue[tempCompCount];
                tempCompCount++;
            }
        }
        await newWorkbook.xlsx.writeFile(filePath);
    }


}

module.exports = LibraryFunctions;
