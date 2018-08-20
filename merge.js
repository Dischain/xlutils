'use strict';

const Doc = require('./lib/objects/Doc');

/*************************************************************************************
 * Слияние двух документов с дирекциями и ответственными исполнителями
 * Производить отдельно по дирекциям. Предварительно привести все ячейки
 * к одному облику путем копирования свойств (в том числе белых)
 ************************************************************************************/
// const to1Path   = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
//       from1Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18 — копия.xlsx',
//       sheet = undefined;

// const to1 = new Doc(to1Path, sheet, 'C15:C99'), // C100:C197
//       from1 = new Doc(from1Path, sheet, 'C15:C99');

// to1.constructObjects('AC:CQ', 'C15', 'C16')
// .then(() => to1.constructObjects('AC:CQ', null, 'C16', 'C17'))
// .then(() => from1.constructObjects('AC:CQ', 'C15', 'C16'))
// .then(() => from1.constructObjects('AC:CQ', null, 'C16', 'C17'))
// .then(() => to1.buildFieldSet('AC8:CQ8', 3))
// .then(() => from1.buildFieldSet('AC8:CQ8', 3))
// .then(() => {
//   from1.merge(to1);
//   return to1.save('C:/Users/ShaytanovAI/Documents/saved.xlsx');
// })
// .then(() => {
//   console.log('saved');
// })
// .catch(console.log);

/*************************************************************************************
 * Слияние двух документов - одного с дирекциями и ответственными исполнителями, 
 * второго - с ответственными исполнителями
 ************************************************************************************/
// const to2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx', 
//       from2Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018 — копия.xlsx';      

// const to2 = new Doc(to2Path, sheet, 'C15:C99'),
//       from2 = new Doc(from2Path, sheet, 'C16:C211');

// to2.constructObjects('AC:CQ', 'C15', 'C16')
// .then(() => to2.constructObjects('AC:CQ', null, 'C16', 'C17'))
// .then(() => from2.constructObjects('F:BP', null, 'C16', 'C17'))
// .then(() => to2.buildFieldSet('AC8:CQ8', 3))
// .then(() => from2.buildFieldSet('F8:BP8', 3))
// .then(() => {
//   from2.merge(to2);
//   return to2.save('C:/Users/ShaytanovAI/Documents/saved2.xlsx');
// })
// .then(() => {
//   console.log('saved');
// })
// .catch(console.log);

/*************************************************************************************
 * Слияние документов по объектам
 ************************************************************************************/
const to3Path   = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx', 
      from3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018 — копия.xlsx',
      sheet = undefined;
      
const to3 = new Doc(to3Path, sheet, 'C17:C230'),
      from3 = new Doc(from3Path, sheet, 'C17:C208');

from3.constructObjects('F:BP')
.then(() => to3.constructObjects('AC:CQ'))
.then(() => from3.buildFieldSet('F8:BP8', 3))
.then(() => to3.buildFieldSet('AC8:CQ8', 3))
.then(() => {
  from3.merge(to3);
  return to3.save('C:/Users/ShaytanovAI/Documents/saved3.xlsx');
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