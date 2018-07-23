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

  it('should find parent by given parent color', (done) => {
    wsPromise.then((ws) => {
      const parentFG = ws.getCell('C11').fill.fgColor,
            docObj1 = new DocObject('C12', 'C12:I12', ws, 'C10:C20', parentFG),
            docObj2 = new DocObject('C13', 'C13:I13', ws, 'C10:C20', parentFG),
            docObj3 = new DocObject('C13', 'C13:I13', ws, 'C10:C20', parentFG);
      
      assert.equal(docObj1.getParentCell().value, 'parent2_1');
      assert.equal(docObj2.getParentCell().value, 'parent2_1');
      assert.equal(docObj3.getParentCell().value, 'parent2_1');
      done();
    });
  });

  it('should collect all object children given child color', (done) => {
    wsPromise.then((ws) => {
      const childFG = null,
            docObj1 = new DocObject('C11', 'C11:I11', ws, 'C10:C20', null, 'empty'),
            docObj2 = new DocObject('C15', 'C15:I15', ws, 'C10:C20', null, 'empty');
      
      assert.deepEqual(Object.keys(docObj1._children), ['obj1', 'obj2', 'obj3']);
      assert.deepEqual(Object.keys(docObj2._children), ['obj4', 'obj5', 'obj6', 'obj7']);
      done();
    });
  });
});