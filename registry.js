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

// const sourcePath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/Приказ о закреплении.xlsx';
// const dir = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/ДСОСС/31.08.2018';
// const doc1 = new Doc(sourcePath, undefined, 'A7:A231');
// let unmatched = [];
// doc1.constructObjects('plain', 'A:L')
// .then(() => {
//   XL.readFile('C:/Users/ShaytanovAI/Downloads/Реестр справок_шаблон.xlsx').then((res) => {
//     const { wb, ws } = res;

//     let i = 3;
    
//     listFilesById(dir, '_').forEach((item) => {
      
//       const plainRow = doc1.getPlainRows()[item];
      
//       if (plainRow != undefined) {
//         let row = plainRow.row;
//         row.getValuesRange().forEach((cell) => {
//           ws.getCell(cell.getColumn() + i).value = cell.getValue();        
//         });
//         i ++;
//       } else {
//         unmatched.push(item);
//       }
//     });

//     return XL.save('C:/Users/ShaytanovAI/Documents/work/формирование реестра/temp ДСОСС.xlsx', wb);
//   })
//   .then(() => { console.log(unmatched); })
// });

// Слить инфо по ответственной дирекции
// const fromPath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/Сводная форма по ГРБС актуальный_ЭКСПЕРИМЕНТ.xlsx',
//       toPath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/temp ДСОСС.xlsx';

// const from = new Doc(fromPath, undefined, 'A18:A267'),
//       to = new Doc(toPath, undefined, 'A3:A75'); // <----

// from.constructObjects('plain', 'G:H')
// .then(() => to.constructObjects('plain', 'M:N'))
// .then(() => from.buildFieldSet('G6:H6'))
// .then(() => to.buildFieldSet('M1:N1'))
// .then(() => {
//   from.merge(to);
//   return to.save('C:/Users/ShaytanovAI/Documents/work/формирование реестра/Реестр ДСОСС.xlsx')
// })
// .catch(console.log);

// Найти отсутствующие
const fromPath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/Сводная форма по ГРБС актуальный_ЭКСПЕРИМЕНТ.xlsx',
      toPath = 'C:/Users/ShaytanovAI/Documents/work/формирование реестра/Реестр ДСОСС.xlsx';

const from = new Doc(fromPath, undefined, 'A18:A267'),
      to = new Doc(toPath, undefined, 'A3:A75'); // <----

from.constructObjects('plain', 'G:J')
.then(() => to.constructObjects('plain', 'M:N'))
.then(() => {
  let objsOfInterest = {}
  let diff = {};
  for (let a in from.getPlainRows()) {
    let curRow = from.getPlainRows()[a];

    console.log(curRow.row.getName());

    if (curRow.row.getValueByColIndex(7) == 'ДСОСС' && 
        (curRow.row.getValueByColIndex(9) == 'конкурс СМР' || 
        curRow.row.getValueByColIndex(9) == 'СМР' || 
        curRow.row.getValueByColIndex(9) == 'СМР не ведутся'))
    {
      objsOfInterest[curRow.row.getName()] = curRow;
      console.log('added: ' + curRow.row.getName());
    }
  }
  console.log(Object.keys(objsOfInterest).length);

  for (let a in objsOfInterest) {
    if (to.getPlainRows()[a] == undefined) { 
      diff[a] = objsOfInterest[a];
    }
  }

  for (let a in diff) {
    console.log(a + ' ' + diff[a].row.getValueByColIndex(7));
  }
})
.catch(console.log);