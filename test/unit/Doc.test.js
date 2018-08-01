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

    it('should construct document with mid objects and with children', (done) => {
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
          doc5Path = 'H:/работа/xlutils/test/doc_diff_plain1.xlsx',
          doc6Path = 'H:/работа/xlutils/test/doc_diff_plain2.xlsx',
          sheet = 'лист1';

    const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
          doc2 = new Doc(doc2Path, undefined, 'C10:C21'),
          doc3 = new Doc(doc3Path, undefined, 'C10:C25'),
          doc4 = new Doc(doc4Path, undefined, 'C10:C19'),
          doc5 = new Doc(doc5Path, undefined, 'C10:C21'),
          doc6 = new Doc(doc6Path, undefined, 'C10:C18');

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

    it('should return difference between two documents with plain rows', (done) => {            
      doc5.constructObjects('D:I')
      .then(() => doc6.constructObjects('D:I'))
      .then(() => {
        const diff = doc5.diff(doc6);
        
        assert.notEqual(diff['obj2'], undefined);
        assert.notEqual(diff['obj5'], undefined);
        assert.notEqual(diff['obj12'], undefined);
        done();
      })
      .catch(console.log);
    });
  });

  describe('Document fieldset creation', () => {
    const doc1Path = 'H:/работа/xlutils/test/doc_fieldset_1.xlsx',
          doc2Path = 'H:/работа/xlutils/test/doc_fieldset_2.xlsx',
          doc3Path = 'H:/работа/xlutils/test/doc_fieldset_3.xlsx',
          sheet = 'лист1';

    const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
          doc2 = new Doc(doc2Path, undefined, 'C10:C21'),
          doc3 = new Doc(doc3Path, undefined, 'C10:C27');

    it('should correctly create document with n-level fieldset', (done) => {
      doc1.buildFieldSet('D6:U6', 3)
      .then(() => doc2.buildFieldSet('D6:U6', 3))
      .then(() => doc3.buildFieldSet('D6:U6', 3))
      .then(() => {
        assert.equal(doc1.equalsFields(doc2), true);
        assert.equal(doc1.equalsFields(doc3), false);
        done();
      }).catch(console.log);
    });

    it('should correctly create document with 0-level fieldset', (done) => {
      doc1.buildFieldSet('D9:U9')
      .then(() => doc2.buildFieldSet('D9:U9'))
      .then(() => doc3.buildFieldSet('D9:U9'))
      .then(() => {
        assert.equal(doc1.equalsFields(doc2), true);
        assert.equal(doc1.equalsFields(doc3), false);
        done();
      }).catch(console.log);
    });
  });

  describe('Document merging', () => {
    const doc1Path = 'H:/работа/xlutils/test/doc_fieldset_1.xlsx',
          doc2Path = 'H:/работа/xlutils/test/doc_fieldset_2.xlsx',          
          sheet = 'лист1';

    const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
          doc2 = new Doc(doc2Path, undefined, 'C10:C21');

    it('should correctly merge document with plain rows', (done) => 
    {
      doc1.constructObjects('D:U')
      .then(() => doc2.constructObjects('D:U'))
      .then(() => doc1.buildFieldSet('D6:U6', 3))
      .then(() => doc2.buildFieldSet('D6:U6', 3))
      .then(() => {
        doc1.equalsFields(doc2);
        doc1.merge(doc2);
        assert.equal(doc2.getPlainRows()['obj4'].row._values[5].getValue(), 'merge_1');
        assert.equal(doc2.getPlainRows()['obj1'].row._values[21].getValue(), 'merge_4');
        done();
      }).catch(console.log);
    });

    it('should correctly merge document with top and mid objects with doc which contains '
       + 'mid objects only', (done) => 
    {
      doc1.constructObjects('D:U', 'C10', 'C11')
      .then(() => doc2.constructObjects('D:U', null, 'C11', 'C12'))
      .then(() => doc1.constructObjects('D:U', null, 'C11', 'C12'))
      .then(() => doc1.buildFieldSet('D6:U6', 3))
      .then(() => doc2.buildFieldSet('D6:U6', 3))
      .then(() => {
        doc1.equalsFields(doc2);
        doc1.merge(doc2);
        assert.equal(doc2.getMidRows()['parent2_2'].getChildren()['obj4']._values[5].getValue(), 'merge_1');
        assert.equal(doc2.getMidRows()['parent2_1'].getChildren()['obj1']._values[21].getValue(), 'merge_4');
        done();
      }).catch(console.log);
    });    
  });

  // describe('Rows append', () => {
  //   const doc1Path = 'H:/работа/xlutils/test/doc_fieldset_1.xlsx',
  //         sheet = 'лист1';

  //   let doc1 = new Doc(doc1Path, sheet, 'C10:C27');

  //   it('should correctly append row', (done) => {
  //     doc1.appendRow(17, 'C', 'C:U', 'heyho!');
  //     doc1 = new Doc(doc1Path, undefined/*sheet*/, 'C10:C27');
  //     // done();
  //     doc1.constructObjects('D:U')
  //     .then(() => {
  //       assert.equal(doc1.getPlainRows()['heyho!'].row._values[5].getValue(), 'val5_12');
  //       done();
  //     }).catch(console.log);
  //   });
  // });

  describe('Document saving', () => {
    const doc1Path = 'H:/работа/xlutils/test/doc_fieldset_1.xlsx',
          doc2Path = 'H:/работа/xlutils/test/doc_fieldset_2.xlsx',          
          sheet = 'лист1';    

    // it('should correctly save document', (done) => {
    //   const doc1 = new Doc(doc1Path, undefined, 'C10:C27');
    //   doc1.constructObjects('D:U', null, 'C11', 'C12')
    //   .then(() => doc1.save('H:/работа/xlutils/test/saved.xlsx'))
    //   .then(() => {
    //     console.log('saved');
    //     done();
    //   })
    //   .catch(console.log);
    // });

    it('should correctly save document after merging document with top and '
      + 'mid objects with doc which contains mid objects only', (done) => 
    {
      const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
            doc2 = new Doc(doc2Path, undefined, 'C10:C21');

      doc1.constructObjects('D:U', 'C10', 'C11')
      .then(() => doc2.constructObjects('D:U', null, 'C11', 'C12'))
      .then(() => doc1.constructObjects('D:U', null, 'C11', 'C12'))
      .then(() => doc1.buildFieldSet('D6:U6', 3))
      .then(() => doc2.buildFieldSet('D6:U6', 3))
      .then(() => {
        doc1.equalsFields(doc2);
        doc1.merge(doc2);
        assert.equal(doc2.getMidRows()['parent2_2'].getChildren()['obj4']._values[5].getValue(), 'merge_1');
        assert.equal(doc2.getMidRows()['parent2_1'].getChildren()['obj1']._values[21].getValue(), 'merge_4');
        console.log(doc2._ws.getCell('D14').value);
        return doc2.save('H:/работа/xlutils/test/saved_mid.xlsx');        
      })
      .then(() => {
        console.log('saved');
        done();
      })
      .catch(console.log);
    });

    /*it('should correctly merge document with plain rows', (done) => 
    {
      const doc1 = new Doc(doc1Path, undefined, 'C10:C27'),
            doc2 = new Doc(doc2Path, undefined, 'C10:C21');

      doc1.constructObjects('D:U')
      .then(() => doc2.constructObjects('D:U'))
      .then(() => doc1.buildFieldSet('D6:U6', 3))
      .then(() => doc2.buildFieldSet('D6:U6', 3))
      .then(() => {
        doc1.equalsFields(doc2);
        doc1.merge(doc2);
        assert.equal(doc2.getPlainRows()['obj4'].row._values[5].getValue(), 'merge_1');
        assert.equal(doc2.getPlainRows()['obj1'].row._values[21].getValue(), 'merge_4');
        return doc2.save('H:/работа/xlutils/test/saved_plain.xlsx');        
      })
      .then(() => {
        console.log('saved');
        done();
      })
      .catch(console.log);
    });*/
  });
});