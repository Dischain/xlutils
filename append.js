'use strict';

const Doc = require('./lib/objects/Doc');

/**
 * Добавление строки
 */
const doc1Path = '‪C:/Users/ShaytanovAI/Downloads/append_ГРБС.XLSX',
      doc2Path = 'C:/Users/ShaytanovAI/Downloads/append_МинЭк.xlsx',
      sheet = 'укрупненно';

const doc1 = new Doc(doc1Path, sheet, 'E15:E264'),
      doc2 = new Doc(doc2Path, sheet, 'C16:C211');

try {
  // doc1.appendRow(18, 'E', 'E:S', 'say hello');
  doc2.appendRow(18, 'C', 'F:BP', 'say hello');
} catch (e) { console.log(e); }