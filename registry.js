'use strict';

const Doc = require('./lib/objects/Doc'),
      XL = require('./lib/utils/xl'),
      fs = require('fs');

function listFilesById(path, sep) {
  return fs.readdirSync(path).map((item) => {
    const splittedArr = item.split(sep);
    if (splittedArr.length >= 2)
      return splittedArr[0];
  });
}

const sourcePath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/Приказ о закреплении.xlsx';
const dir = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/24.08.2018';
const doc1 = new Doc(sourcePath, undefined, 'A4:A210');
let unmatched = [];
doc1.constructObjects('A:L')
.then(() => {
  XL.readFile('C:/Users/ShaytanovAI/Documents/work/формирование реестра/registry.xlsx').then((res) => {
    const { wb, ws } = res;

    let i = 3;
    
    listFilesById(dir, '_').forEach((item) => {
      let row = doc1.getPlainRows()[item].row;

      if (row != undefined) {
        row.getValuesRange().forEach((cell) => {
          ws.getCell(cell.getColumn() + i).value = cell.getValue();        
        });
        i ++;
      } else {
        unmatched.push(item);
      }
    });

    return XL.save('C:/Users/ShaytanovAI/Documents/work/формирование реестра/registry.xlsx', wb);
  })
  .then(() => { console.log(unmatched); })
});
