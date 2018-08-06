'use strict';

const assert = require('assert'),
      Doc = require('../../lib/objects/Doc');

describe('Doc integration test', () => {
  describe('Work with document', () => {
    const doc1Path = 'H:/работа/xlutils_integration/Сводная форма по ГРБС на 02.08.2018_ЭКСПЕРИМЕНТ.XLSX',
          doc2Path = 'H:/работа/xlutils_integration/Информация_по_заключенным_договорам ТЗ_16.05.18.xlsx',
          sheet = undefined;

    const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
          doc2 = new Doc(doc2Path, sheet, 'C15:C230');

    it('should get the difference between two large documents with both top objects', (done) => {      
      doc1.constructObjects('F:S', 'E15', 'E16')
          .then(() => doc1.constructObjects('F:S', null, 'E16', 'E17'))
          .then(() => doc2.constructObjects('D:CM', 'C15', 'C16'))
          .then(() => doc2.constructObjects('D:CM', null, 'C16', 'C17'))
          .then(() => {
            // console.log(doc1.diff(doc2));
            done();
          })
          .catch(console.log);
    });
  });
});