'use strict';

const Doc = require('./lib/objects/Doc');

const doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
      doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx',
      sheet = undefined;

let doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
    doc2 = new Doc(doc2Path, sheet, 'C15:C230'),
    doc3 = new Doc(doc3Path, sheet, 'C16:C211');

const diff1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС_шаблон.XLSX';
// По умолчанию зададим количество строк в документе равное 200.
let diff1Doc = new Doc(diff1Path, sheet, 'C14:C214');
/*************************************************************************************
* Определение разницы по двум документам с учетом дирекций и ответственных
* исполнителей в одном документе и только ответстсвенных исполнителей в другом.
*************************************************************************************/
// doc1.constructObjects('F:S', 'E15', 'E16')
// .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
// .then(() => doc3.constructObjects('D:CI', null, 'C16', 'C17'))
// .then(() => {
//   const diff = doc1.diff(doc3);
//   for (let a in diff) {
//     console.log('Дирекция: ' + a);
//     // console.log(diff[a]);
//     for (let b in diff[a]) {
//       console.log('Отв. исп.: ' + b);
//       // console.log(diff[a][b]);
//       for (let c in diff[a][b]) {
//         console.log('Объект: ' + c);
//         // console.log(diff[a][b][c]);
//       }
//     }
//   }
// })
// .catch(console.log);
/*****************************************/
/* + вывод в файл добавленных документов */
/*****************************************/
doc1.constructObjects('F:S', 'E15', 'E16')
.then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
.then(() => doc3.constructObjects('D:CI', null, 'C16', 'C17'))
.then(() => diff1Doc.constructObjects('F:S'))
.then(() => {
  const startRow = 14,
        valueCol = 'E',
        diff = doc1.diff(doc3);
  
  let i = startRow;
  for (let a in diff) {
    diff1Doc._ws.getCell(valueCol + i).value = a;
    i ++;
    for (let b in diff[a]) {
      diff1Doc._ws.getCell(valueCol + i).value = b;
      i ++;
      for (let c in diff[a][b]) {
        diff1Doc._ws.getCell(valueCol + i).value = c;
        i ++;
      }
    }
  }
  // Предварительное сохранение
  return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff1saved.xlsx');  
})
.then(() => {
  diff1Doc = new Doc(diff1Path, sheet, 'E14:E214');
  return diff1Doc.constructObjects('F:P');
})
.then(() => {
  doc1 = new Doc(doc1Path, sheet, 'E15:E264')
  // Повторное построение исходного документа
  return doc1.constructObjects('F:P');
})
.then(() => {
  for (let a in diff1Doc.getPlainRows()) { console.log(a); }
  return Promise.resolve();
})
// Построить поля
.then(() => doc1.buildFieldSet('F6:P6', 1))
.then(() => diff1Doc.buildFieldSet('F6:P6', 1))
.then(() => {
  // Слить
  // const diff = doc1.diff(diff1Doc);
  const diff = diff1Doc.diff(doc1);
  for (let a in diff) {
    console.log('Дирекция: ' + a);
    for (let b in diff[a]) {
      console.log('Отв. исп.: ' + b);
      for (let c in diff[a][b]) {
        console.log('Объект: ' + c);
      }
    }
  }
  doc1.merge(diff1Doc);
  // Сохранить
  return diff1Doc.save('C:/Users/ShaytanovAI/Documents/diff1saved.xlsx')
})
.catch(console.log);
 
/**************************************************************************************
* Определение разницы по двум документам с учетом дирекций и 
* ответственных исполнителей под ними.
*************************************************************************************/
// doc1.constructObjects('F:S', 'E15', 'E16')
//     .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
//     .then(() => doc2.constructObjects('D:CM', 'C15', 'C16'))
//     .then(() => doc2.constructObjects('D:CM', null, 'C16', 'C17'))
//     .then(() => { 
//       const diff = doc1.diff(doc2);
//       for (let a in diff) {
//         console.log('Дирекция: ' + a);
//         // console.log(diff[a]);
//         for (let b in diff[a]) {
//             console.log('Отв. исп.: ' + b);
//             // console.log(diff[a][b]);
//           for (let c in diff[a][b]) {
//             console.log('Объект: ' + c);
//             // console.log(diff[a][b][c]);
//           }
//         }
//         console.log('----');
//       }
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

// fromPR.constructObjects('B:F')
// .then(() => toPR.constructObjects('C:G'))
// .then(() => {
//   const diff = fromPR.diff(toPR);
//   console.log(Object.keys(diff).length);
//   for (let a in diff) { console.log(a); }
// })
// .catch(console.log);

