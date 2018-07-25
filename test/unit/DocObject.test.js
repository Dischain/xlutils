'use strict';

const assert = require('assert'),
      Field = require('../../lib/items/Field'),
      Row = require('../../lib/items/Row'),
      DocObject = require('../../lib/objects/DocObject'),
      readFile = require('../../lib/utils/xl').readFile;

describe('DocObject', () => {
  const wsPromise = readFile('H:/работа/xlutils/test/doc_obj_test.xlsx');

  it('should correctly construct document object', (done) => {
    wsPromise.then((ws) => {
      const docObj = new DocObject('C11', 'C12:I12', ws);      
      done();
    });
  });

  // it('should find parent by given parent color', (done) => {
  //   wsPromise.then((ws) => {
  //     const parentFG = ws.getCell('C11').fill.fgColor,
  //           docObj1 = new DocObject('C12', 'C12:I12', ws, 'C10:C20', parentFG),
  //           docObj2 = new DocObject('C13', 'C13:I13', ws, 'C10:C20', parentFG),
  //           docObj3 = new DocObject('C13', 'C13:I13', ws, 'C10:C20', parentFG);
      
  //     assert.equal(docObj1.getParentCell().getValue(), 'parent2_1');
  //     assert.equal(docObj2.getParentCell().getValue(), 'parent2_1');
  //     assert.equal(docObj3.getParentCell().getValue(), 'parent2_1');
  //     done();
  //   });
  // });

  it('should collect all object children by the given child color and set them ' +
     'given object as parent', (done) => {
    wsPromise.then((ws) => {
      const fg1 = ws.getCell('C10').fill.fgColor,
            fg2 = ws.getCell('C11').fill.fgColor,
            docObj1 = new DocObject('C11', 'C11:I11', ws, 'C10:C27', null, 'empty', fg2),
            docObj2 = new DocObject('C15', 'C15:I15', ws, 'C10:C27', null, 'empty', fg2),
            docObj3 = new DocObject('C10', 'C10:I10', ws, 'C10:C27', null, fg2, fg1),
            docObj4 = new DocObject('C19', 'C19:I19', ws, 'C10:C27', null, fg2, fg1);
      
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
    wsPromise.then((ws) => {
      const my1FG = ws.getCell('C11').fill.fgColor,
            my2FG = ws.getCell('C10').fill.fgColor;

            const docObj1 = new DocObject('C11', 'C11:I11', ws, 'C10:C27', my2FG, null, my1FG),
                  docObj2 = new DocObject('C10', 'C10:I10', ws, 'C10:C27', null, null, my2FG),
                  docObj3 = new DocObject('C15', 'C15:I15', ws, 'C10:C27', my2FG, null, my1FG),
                  docObj4 = new DocObject('C19', 'C19:I19', ws, 'C10:C27', null, null, my2FG),
                  docObj5 = new DocObject('C20', 'C20:I20', ws, 'C10:C27', null, null, my1FG);
      assert.equal(docObj1._getChildrenLimit(), 'C14');
      assert.equal(docObj2._getChildrenLimit(), 'C18');
      assert.equal(docObj3._getChildrenLimit(), 'C18');
      assert.equal(docObj4._getChildrenLimit(), 'C27');
      assert.equal(docObj5._getChildrenLimit(), 'C23');
      done();
    });
  });
});