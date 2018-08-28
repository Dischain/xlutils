'use strict';

const assert = require('assert'),
      Field = require('../../lib/items/Field'),
      Row = require('../../lib/items/Row'),
      DocObject = require('../../lib/objects/DocObject'),
      readFile = require('../../lib/utils/xl').readFile;

describe('DocObject', () => {
  const wsPromise = readFile('C:/Users/ShaytanovAI/Documents/work/xlutils/test/doc_obj_test.xlsx');

  it('should correctly construct document object', (done) => {
    wsPromise.then((res) => {
      const docObj = new DocObject('C11', 'C:I', res.ws);      
      done();
    });
  });

  it('should collect all object children by the given child color and set them ' +
     'given object as parent', (done) => {
    wsPromise.then((res) => {
      const topParents = ['parent1_1', 'parent1_2'],
            midParents = ['parent2_1', 'parent2_2', 'parent2_3', 'parent2_4'],
            docObj1 = new DocObject('C11', 'C:I', res.ws, 'C10:C27', topParents, 'empty', midParents),
            docObj2 = new DocObject('C15', 'C:I', res.ws, 'C10:C27', topParents, 'empty', midParents),
            docObj3 = new DocObject('C10', 'C:I', res.ws, 'C10:C27', null, midParents, topParents),
            docObj4 = new DocObject('C19', 'C:I', res.ws, 'C10:C27', null, midParents, topParents);
      
      assert.deepEqual(Object.keys(docObj1.getChildren()), ['obj1', 'obj2', 'obj3']);    
      assert.deepEqual(Object.keys(docObj2.getChildren()), ['obj4', 'obj5', 'obj6']);
      assert.deepEqual(Object.keys(docObj3.getChildren()), ['parent2_1', 'parent2_2']);
      assert.deepEqual(Object.keys(docObj4.getChildren()), ['parent2_3', 'parent2_4']);
      
      assert.equal(docObj1.getChildren()['obj2'].getName(), 'obj2');
      assert.equal(docObj3.getChildren()['parent2_2'].getName(), 'parent2_2');

      done();
    });
  });

  it('should return children range of a given parent', (done) => {
    wsPromise.then((res) => {
      const my1FG = ['parent1_1', 'parent1_2'],
            my2FG = ['parent2_1', 'parent2_2', 'parent2_3', 'parent2_4'];

      const docObj1 = new DocObject('C10', 'C:I', res.ws, 'C10:C27', null,  my2FG, my1FG),
            docObj2 = new DocObject('C11', 'C:I', res.ws, 'C10:C27', my1FG, null, my2FG),            
            docObj3 = new DocObject('C15', 'C:I', res.ws, 'C10:C27', my1FG, null, my2FG),
            docObj4 = new DocObject('C19', 'C:I', res.ws, 'C10:C27', null,  my2FG, my1FG),
            docObj5 = new DocObject('C20', 'C:I', res.ws, 'C10:C27', my1FG, null, my2FG);

      assert.equal(docObj1._getChildrenLimit(), 'C18');
      assert.equal(docObj2._getChildrenLimit(), 'C14');
      assert.equal(docObj3._getChildrenLimit(), 'C18');
      assert.equal(docObj4._getChildrenLimit(), 'C27');
      assert.equal(docObj5._getChildrenLimit(), 'C23');
      done();
    });
  });
});