'use strict';

const Doc = require('./lib/objects/Doc');

const doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
      doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx',
      sheet = undefined;

const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
      doc2 = new Doc(doc2Path, sheet, 'C15:C230'),
      doc3 = new Doc(doc3Path, sheet, 'C16:C211');

/**
 * Формирование документа с дирекциями
 */
// doc1.constructObjects('F:S', 'E15', 'E16')
//     .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
//     .then(() => {
//       for (let a in doc1.getTopRows()) {
//         console.log('Дирекция:' + a);
//         console.log(doc1.getTopRows()[a].getChildrenRange().getAddress().toString());
//         for (let b in doc1.getTopRows()[a].getChildren()) {
//           console.log('Отв. исп.: ' + b);
//           // console.log(doc1.getTopRows()[a].getChildren()[b])
//         }
//       }
//     })
// .catch(console.log);
// doc2.constructObjects('D:CM', 'C15', 'C16')
//     .then(() => doc2.constructObjects('D:CM', null, 'C16', 'C17'))
//     .then(() => {
//       for (let a in doc2.getTopRows()) {
//         console.log('Дирекция:' + a);
//         console.log(doc2.getTopRows()[a].getChildrenRange().getAddress().toString());
//         for (let b in doc2.getTopRows()[a].getChildren()) {
//           console.log('Отв. исп.: ' + b);
//           // console.log(doc2.getTopRows()[a].getChildren()[b])
//         }
//       }
//     })
// .catch(console.log);

/**
 * Формирование документа с ответственными исполнителями
 */
// doc1.constructObjects('F:S', null, 'E16', 'E17')
//     .then(() => {
//       for (let a in doc1.getMidRows()) {
//         console.log('Отв. исп.:' + a);
//         console.log(doc1.getMidRows()[a].getChildrenRange().getAddress().toString());
//         for (let b in doc1.getMidRows()[a].getChildren()) {
//           console.log('Объект: ' + b);
//           // console.log(doc1.getMidRows()[a].getChildren()[b])
//         }
//       }
//     })
// .catch(console.log);
// doc2.constructObjects('F:S', null, 'E16', 'E17')
//     .then(() => {
//       for (let a in doc2.getMidRows()) {
//         console.log('Отв. исп.:' + a);
//         console.log(doc2.getMidRows()[a].getChildrenRange().getAddress().toString());
//         for (let b in doc2.getMidRows()[a].getChildren()) {
//           console.log('Объект: ' + b);
//           // console.log(doc2.getMidRows()[a].getChildren()[b])
//         }
//       }
//     })
// .catch(console.log);