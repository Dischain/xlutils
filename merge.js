'use strict';

const Doc = require('./lib/objects/Doc');

const to1Path   = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      from1Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18 — копия.xlsx',
      doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
      doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx',      
      sheet = undefined;

const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
      to1 = new Doc(to1Path, sheet, 'C15:C230'),
      doc3 = new Doc(doc3Path, sheet, 'C16:C211'),
      from1 = new Doc(from1Path, sheet, 'C15:C230');

/**
 * Слияние двух документов с дирекциями и ответственными исполнителями
 */
to1.constructObjects('AC:CQ', 'C15', 'C16')
.then(() => to1.constructObjects('AC:CQ', null, 'C16', 'C17'))
.then(() => from1.constructObjects('AC:CQ', 'C15', 'C16'))
.then(() => from1.constructObjects('AC:CQ', null, 'C16', 'C17'))
.then(() => to1.buildFieldSet('AC8:CQ8', 3))
.then(() => from1.buildFieldSet('AC8:CQ8', 3))
.then(() => {
  from1.merge(to1);

  // console.log(from1._ws.getCell('BM21').value);
  // console.log(from1._ws.getCell('BY24').value);
  // console.log(to1._ws.getCell('BM21').value);
  // console.log(to1._ws.getCell('BY24').value);
  return to1.save('C:/Users/ShaytanovAI/Documents/saved.xlsx');
})
.then(() => {
  console.log('saved');
})
.catch(console.log);

// from1.constructObjects('AC:CQ', 'C15', 'C16')
// .then(() => to1.constructObjects('AC:CQ', 'C15', 'C16'))
// .then(() => {
//   return Object.keys(from1.getTopRows()).reduce((acc, curTopObjName) => {
//     const curFromTopObj = from1.getTopRows()[curTopObjName],
//           curToTopObj = to1.getTopRows()[curTopObjName];

//     const curFrom1 = new Doc(from1Path, sheet, curFromTopObj.getChildrenRange()),
//           curTo1 = new Doc(to1Path, sheet, curToTopObj.getChildrenRange());
    
//     return acc
//       .then(() => curFrom1.constructObjects('AC:CQ', null, 'C16', 'C17')
//       .then(() => curTo1.constructObjects('AC:CQ', null, 'C16', 'C17'))    
//       .then(() => curFrom1.buildFieldSet('AC8:CQ8', 3))
//       .then(() => curTo1.buildFieldSet('AC8:CQ8', 3))
//       .then(() => curFrom1.merge(curTo1)))
//       .then(() => curTo1.save('C:/Users/ShaytanovAI/Documents/saved.xlsx'));
//   }, Promise.resolve())
// })
// .catch(console.log);