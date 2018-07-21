const Excel = require('exceljs');

class XL {
  static createWS(wsName) {
    const wb = new Excel.Workbook(),
          ws = wb.addWorksheet('test');

    return ws;
  }

  static readFile(filePath, sheetName) {
    const wb = new Excel.Workbook();
    return wb.xlsx.readFile(filePath).then(() => {
      return Promise.resolve(wb.getWorksheet(sheetName));
    });
  }
}

module.exports = XL;