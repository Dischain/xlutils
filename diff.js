'use strict';

const Doc = require('./lib/objects/Doc');

// const doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
//       doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
//       doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx',
//       sheet = undefined;

// let doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
//     doc2 = new Doc(doc2Path, sheet, 'C15:C230'),
//     doc3 = new Doc(doc3Path, sheet, 'C16:C211');

/*************************************************************************************
* Определение разницы по двум документам с учетом дирекций и ответственных
* исполнителей в одном документе и только ответстсвенных исполнителей в другом.
*************************************************************************************/
/*****************************************/
/* + вывод в файл добавленных документов */
/*****************************************/
// let doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
//     doc3 = new Doc(doc3Path, sheet, 'C16:C211');

// const diff1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС_шаблон.XLSX';
// // По умолчанию зададим количество строк в документе равное 200.
// let diff1Doc = new Doc(diff1Path, sheet, 'C14:C214');

// doc1.constructObjects('F:S', 'E15', 'E16')
// .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
// .then(() => doc3.constructObjects('D:CI', null, 'C16', 'C17'))
// .then(() => diff1Doc.constructObjects('F:S'))
// .then(() => {
//   const startRow = 14,
//         valueCol = 'E',
//         diff = doc1.diff(doc3);
  
//   let i = startRow;
//   for (let a in diff) {
//     diff1Doc._ws.getCell(valueCol + i).value = a;
//     i ++;
//     for (let b in diff[a]) {
//       diff1Doc._ws.getCell(valueCol + i).value = b;
//       i ++;
//       for (let c in diff[a][b]) {
//         diff1Doc._ws.getCell(valueCol + i).value = c;
//         i ++;
//       }
//     }
//   }
//   // Предварительное сохранение
//   return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff1saved.xlsx');  
// })
// .catch(console.log);
/*****************************************/
/* + заливка в этот файл данных          */
/*****************************************/
// const diff1Path = 'C:/Users/ShaytanovAI/Documents/diff1saved.xlsx';
// let doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
//     diff1Doc = new Doc(diff1Path, sheet, 'E14:E214');

// diff1Doc.constructObjects('F:P')
// // Повторное построение исходного документа
// .then(() => doc1.constructObjects('F:P'))
// .then(() => {
//   console.log('diffdoc plain rows:');
//   for (let a in diff1Doc.getPlainRows()) { console.log(a); }
//   return Promise.resolve();
// })
// // Построить поля
// .then(() => doc1.buildFieldSet('F6:P6', 1))
// .then(() => diff1Doc.buildFieldSet('F6:P6', 1))
// .then(() => {
//   const diff = diff1Doc.diff(doc1);
//   // const diff = diff1Doc.diff(doc1);
//   for (let a in diff) {
//     console.log('Дирекция: ' + a);
//     for (let b in diff[a]) {
//       console.log('Отв. исп.: ' + b);
//       for (let c in diff[a][b]) {
//         console.log('Объект: ' + c);
//       }
//     }
//   }
//   // Слить
//   doc1.merge(diff1Doc);
//   // Сохранить
//   return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff1saved.xlsx')
// })
// .catch(console.log);

/**************************************************************************************
* Определение разницы по двум документам с учетом дирекций и 
* ответственных исполнителей под ними и только ответственных исполнителей
* во втором.
*************************************************************************************/
// const doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
//       doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx';
      
// const diff1Path = 'C:/Users/ShaytanovAI/Downloads/Информация по заключенным договорам_шаблон.xlsx';
// let diff1Doc = new Doc(diff1Path, sheet, 'C14:C214');

// let doc2 = new Doc(doc2Path, sheet, 'C15:C230'),
//     doc3 = new Doc(doc3Path, sheet, 'C16:C211');

// doc3.constructObjects('F:BP', null, 'C16', 'C17')
// .then(() => doc2.constructObjects('AC:CM', 'C15', 'C16'))
// .then(() => doc2.constructObjects('AC:CM', null, 'C16', 'C17'))
// .then(() => diff1Doc.constructObjects('F:BP'))
// .then(() => {
//   const startRow = 14,
//         valueCol = 'C',
//         diff = doc3.diff(doc2);
  
//   let i = startRow;
//   for (let a in diff) {
//     console.log('Отв. исп.:' + a);
//     diff1Doc._ws.getCell(valueCol + i).value = a;
//     i ++;
//     for (let b in diff[a]) {
//       if (b == 'childrenRange') continue;
//       console.log('Объект.:' + b);
//       diff1Doc._ws.getCell(valueCol + i).value = b;
//       i ++;
//     }
//   }
//   // Предварительное сохранение
//   return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff2saved.xlsx');
// })
// .catch(console.log);

/*****************************************/
/* + заливка в этот файл данных          */
/*****************************************/
// const diff1Path = 'C:/Users/ShaytanovAI/Documents/diff2saved.xlsx';
// let diff1Doc = new Doc(diff1Path, sheet, 'C14:C214'),
//     doc3 = new Doc(doc3Path, sheet, 'C16:C211');

// diff1Doc.constructObjects('F:BT')
// .then(() => doc3.constructObjects('F:BT'))
// .then(() => diff1Doc.buildFieldSet('F8:BT8', 3))
// .then(() => doc3.buildFieldSet('F8:BT8', 3))
// .then(() => {
//   doc3.merge(diff1Doc);
//   return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff2saved.xlsx');
// })
// .catch(console.log);

/**************************************************************************************
* Определение разницы по двум документам с учетом только ответственных
* исполнителей.
*************************************************************************************/
// doc1.constructObjects('F:S', null, 'E16', 'empty')
//     .then(() => doc2.constructObjects('D:CM', null, 'C16', 'empty'))
//     .then(() => {
//       const diff = doc1.diff(doc2);
//       for (let a in diff) {
//         console.log('Отв. исп.: ' + a);
//         // console.log(diff[a]);
//         for (let b in diff[a]) {
//           console.log('Объект: ' + b);
//           // console.log(diff[a][b]);
//         }
//         console.log('----');
//       }
// })
// .catch(console.log);

/**************************************************************************************
 * Определение разницы между двумя документами только по обычным строкам
 *************************************************************************************/
// const fromPRPath = 'C:/Users/ShaytanovAI/Documents/temp/08.08.18/Список сотрудников ГКУ в АИС.xlsx',
//       toPRPath = 'C:/Users/ShaytanovAI/Documents/temp/08.08.18/UserView 20180808-1555.xlsx';

// const fromPR = new Doc(fromPRPath, 'UserView', 'B2:B137'),
//       toPR = new Doc(toPRPath, 'UserView', 'A2:A121');

// fromPR.constructObjects('plain', 'B:F')
// .then(() => toPR.constructObjects('plain', 'C:G'))
// .then(() => {
//   const diff = fromPR.diff(toPR);
//   console.log(Object.keys(diff).length);
//   for (let a in diff) { console.log(a); }
// })
// .catch(console.log);

const doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx';

const doc2 = new Doc(doc2Path, undefined, 'C15:C230'),
      doc3 = new Doc(doc3Path, undefined, 'C16:C211');

let topSigns = ['Дирекция'],
    midSigns = ['Министерство', 'Служба', 'Государственный комитет', 'Управление', 'Администрация'];

doc3.constructObjects('mid', 'F:BP', topSigns, midSigns, 'empty')
.then(() => doc2.constructObjects('top', 'AC:CM', topSigns, midSigns))
.then(() => doc2.constructObjects('mid', 'AC:CM', topSigns, midSigns, 'empty'))
.then(() => {
  const diff = doc3.diff(doc2);
  for (let a in diff) { 
    console.log(a); 
    for (let b in diff[a]) { console.log('  ' + b); }
  }
})
.catch(console.log);