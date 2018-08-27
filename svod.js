'use strict';

const Doc = require('./lib/objects/Doc');

const doc1Path = 'C:/Users/ShaytanovAI/Downloads/Сводная форма по ГРБС актуальный_ЭКСПЕРИМЕНТ.xlsx',
      sheet = undefined;

let doc1 = new Doc(doc1Path, sheet, 'E16:E264');

doc1.constructObjects('F:P', 'E16', 'E17')
.then(() => doc1.constructObjects('F:P', null, 'E17', 'E18'))
//.then(() => doc1.buildFieldSet('F6:P6', 1))
.then(() => { 
  for (let a in doc1.getTopRows()) {
    console.log('Дирекция ' + a);
    let pir = []; let concursPir = []; let smr = []; let concursSmr = []; let smrStopped = [];
    let pirSKS = []; let concursPirSKS = []; let smrSKS = []; let concursSmrSKS = []; let smrStoppedSKS = [];

    for (let b in doc1.getTopRows()[a].getChildren()) { 
      console.log('Отв. исп.: ' + b);
      for (let c in doc1.getMidRows()[b].getChildren()) {
        let curRow = doc1.getMidRows()[b].getChildren()[c];
        
        if (curRow.getValueByColIndex(9) == 'СМР') {
          smr.push(1);
          if (curRow.getValueByColIndex(8) == 'Служба Капитального Строительства') {
            smrSks.push(1);
          }
        }

        if (curRow.getValueByColIndex(9) == 'Конкурс СМР') {
          concursSmr.push(1);
          if (curRow.getValueByColIndex(8) == 'Служба Капитального Строительства') {
            concursSmrSKS.push(1);
          }
        }

        if (curRow.getValueByColIndex(9) == 'ПИР') {
          pir.push(1);
          if (curRow.getValueByColIndex(8) == 'Служба Капитального Строительства') {
            pirSKS.push(1);
          }
        }

        if (curRow.getValueByColIndex(9) == 'Конкурс ПИР') {
          concursPir.push(1);
          if (curRow.getValueByColIndex(8) == 'Служба Капитального Строительства') {
            concursPirSKS.push(1);
          }
        }

        if (curRow.getValueByColIndex(9) == 'СМР не ведутся') {
          smrStopped.push(1);
          if (curRow.getValueByColIndex(8) == 'Служба Капитального Строительства') {
            smrStoppedSKS.push(1);
          }
        }
      }
    }
    console.log('Всего');
    console.log('СМР: ' + smr.length);
    console.log('Конкурс СМР: ' + concursSmr.length);
    console.log('ПИР: ' + pir.length);
    console.log('Конкурс ПИР: ' + concursPir.length);
    console.log('СМР не ведутся: ' + smrStopped.length);
  }
})
.catch(console.log);