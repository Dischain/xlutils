'use strict';

const Doc = require('./lib/objects/Doc');

const doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
      doc2Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
      doc3Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма на 01.05.2018.xlsx',
      doc4Path = 'C:/Users/ShaytanovAI/Downloads/Информация_по_заключенным_договорам ТЗ_16.05.18 — копия.xlsx',
      sheet = undefined;

const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
      doc2 = new Doc(doc2Path, sheet, 'C15:C230'),
      doc3 = new Doc(doc3Path, sheet, 'C16:C211'),
      doc4 = new Doc(doc2Path, sheet, 'C15:C230');

/**
 * Слияние двух документов с дирекциями и ответственными исполнителями
 */
doc2.constructObjects('AC:CL', 'C15', 'C16')
.then(() => doc2.constructObjects('AC:CL', null, 'C16', 'C17'))
.then(() => doc4.constructObjects('AC:CL', 'C15', 'C16'))
.then(() => doc4.constructObjects('AC:CL', null, 'C16', 'C17'))
.then(() => doc2.buildFieldSet('AC8:CL8', 3))
.then(() => doc4.buildFieldSet('AC8:CL8', 3))
.then(() => {
  doc4.merge(doc2);
})
.catch(console.log);