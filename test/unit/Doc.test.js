'use strict';

const assert = require('assert'),
      Doc = require('../../lib/objects/Doc'),
      readFile = require('../../lib/utils/xl').readFile;

describe('Doc', () => {
  describe('Document creation', () => {
    const doc1Path = 'H:/работа/xlutils/test/doc_diff_top1.xlsx',
          sheet = 'лист1';
    const doc1 = new Doc(doc1Path, undefined, 'C10:C27');

    it('should correct construct document with plain rows', (done) => {            
      doc1.constructObjects('D:I').then(() => {
        assert.equal(doc1._plainRows['obj5'].row.getValueByColIndex(5), 'val5_12');
        done();
      }).catch(console.log);
    });

    it('should construct document with top objects and with children', (done) => {
      doc1.constructObjects('D:I', 'C10', 'C11').then(() => {
        assert.equal(doc1._topRows['parent1_2'].getChildren()['parent2_3'].getName(), 'parent2_3');
        assert.equal(doc1._topRows['parent1_2'].getChildren()['parent2_4'].getName(), 'parent2_4');
        done();
      }).catch(console.log);
    });

    it('should construct document with mid objects', (done) => {
      doc1.constructObjects('D:I', null, 'C11', 'C12').then(() => {
        assert.equal(doc1._midRows['parent2_1'].getChildren()['obj2'].getName(), 'obj2');
        assert.equal(doc1._midRows['parent2_2'].getChildren()['obj5'].getName(), 'obj5');
        done();
      }).catch(console.log);
    });
  });

  describe('Document diff', () => {
    const doc1Path = 'H:/работа/xlutils/test/doc_diff_top1.xlsx',
          doc2Path = 'H:/работа/xlutils/test/doc_diff_top2.xlsx',
          doc3Path = 'H:/работа/xlutils/test/doc_diff_mid1.xlsx',
          doc4Path = 'H:/работа/xlutils/test/doc_diff_mid2.xlsx',
          sheet = 'лист1';

    const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
          doc2 = new Doc(doc2Path, undefined, 'C10:C21'),
          doc3 = new Doc(doc3Path, undefined, 'C10:C25'),
          doc4 = new Doc(doc4Path, undefined, 'C10:C19');

    it('should return difference between two documents with top and mid objects', (done) => {
      doc1.constructObjects('D:I', 'C10', 'C11')
      .then(() => doc2.constructObjects('D:I', 'C10', 'C11'))
      .then(() => doc1.constructObjects('D:I', null, 'C11', 'C12'))
      .then(() => doc2.constructObjects('D:I', null, 'C11', 'C12'))
      .then(() => {
        const diff = doc1.diff(doc2);

        assert.notEqual(diff['parent1_1']['parent2_1']['obj2'], undefined);
        assert.notEqual(diff['parent1_1']['parent2_2']['obj5'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj10'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj11'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj12'], undefined);

        done();
      }).catch(console.log);
    });

    it('should return difference between two documents with mid objects', (done) => {
      doc3.constructObjects('D:I', null, 'C10', 'empty')
      .then(() => doc4.constructObjects('D:I', null, 'C10', 'empty'))
      .then(() => {
        const diff = doc3.diff(doc4);

        assert.notEqual(diff['parent2_1']['obj2'], undefined);
        assert.notEqual(diff['parent2_2']['obj5'], undefined);
        assert.notEqual(diff['parent2_4']['obj10'], undefined);
        assert.notEqual(diff['parent2_4']['obj11'], undefined);
        assert.notEqual(diff['parent2_4']['obj12'], undefined);

        done();
      }).catch(console.log);
    });

    it('should return difference between top document and mid document', (done) => {
      doc1.constructObjects('D:I', 'C10', 'C11')
      .then(() => doc1.constructObjects('D:I', null, 'C11', 'C12'))
      .then(() => doc4.constructObjects('D:I', null, 'C10', 'empty'))
      .then(() => {        
        const diff = doc1.diff(doc4);

        assert.notEqual(diff['parent1_1']['parent2_1']['obj2'], undefined);
        assert.notEqual(diff['parent1_1']['parent2_2']['obj5'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj10'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj11'], undefined);
        assert.notEqual(diff['parent1_2']['parent2_4']['obj12'], undefined);
        
        done();
      }).catch(console.log);
    });
  });
});