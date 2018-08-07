'use strict';

const Doc = require('./lib/objects/Doc');

const doc1Path = 'H:/работа/xlutils_integration/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
      doc2Path = 'H:/работа/xlutils_integration/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      sheet = undefined;

const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
      doc2 = new Doc(doc2Path, sheet, 'C15:C230');

// doc1.constructObjects('F:S', 'E15', 'E16')
//     .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
//     .then(() => doc2.constructObjects('D:CM', 'C15', 'C16'))
//     .then(() => doc2.constructObjects('D:CM', null, 'C16', 'C17'))
// .then(() => {      
//   const diff = doc1.diff(doc2);
//   for (let a in diff) {
//     console.log('Дирекция: ' + a);        
//     for (let b in diff[a]) {
//       console.log('Отв. исп.: ' + b);
//       for (let c in diff[a][b]) {
//         console.log('Объект: ' + c);
//       }
//     }
//     console.log('----');
//   }
// })
// .catch(console.log);


doc1.constructObjects('F:S', null, 'E16', 'empty')
    .then(() => doc2.constructObjects('D:CM', null, 'C16', 'empty'))
    .then(() => {      
      const diff = doc1.diff(doc2);
      for (let a in diff) {        
        console.log('Отв. исп.: ' + a);
        console.log(diff[a].childrenRange.getAddress().toString());
        
        for (let b in diff[a]) {
          console.log('Объект: ' + b);
          // for (let keys in diff[a][b]) { console.log(keys); }
        }
        console.log('----');
      }
    })
    .catch(console.log);