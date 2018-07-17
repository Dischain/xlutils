const Excel = require('exceljs');

class XL {
  static createWS(wsName) {
    const wb = new Excel.Workbook(),
          ws = wb.addWorksheet('test');

    return ws;
  }
}

module.exports = XL;