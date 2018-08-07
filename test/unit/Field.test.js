'use strict';

const assert = require('assert'),
      Field = require('../../lib/items/Field'),
      Cell = require('../../lib/cells/Cell'),
      createWS = require('../../lib/utils/xl').createWS,
      readFile = require('../../lib/utils/xl').readFile,
      StringUtils = require('../../lib/utils/StringUtils'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException'),
      EmptyFieldException = require('../../lib/exceptions/EmptyFieldException');

describe('Field', () => {
  const ws = createWS('test');
  const fieldAddress = 'B2', childAddress = 'B3';
  let field;
  
  beforeEach((done) => {
    ws.getCell(fieldAddress).value = 'field header';
    ws.getCell(childAddress).value = 'child field header';
    field = new Field(fieldAddress, null, ws);    
    done();
  });

  describe('Instance creation', () => {
    let tempField;

    beforeEach((done) => {
      ws.getCell('C2').value = 'temp field header';
      tempField = new Field('C2', null, ws);
      done();
    });

    after((done) => {
      ws.getCell('C2').value = null;
      done();
    });

    it('should throw AddressFormatException if invalid address provided', (done) => {
      assert.throws(() => new Field('23<', null, ws), AddressFormatException);
      done();
    });
  
    it('should throw EmptyFieldException if cell by given address is empty', (done) => {
      assert.throws(() => new Field('A1', null, ws), EmptyFieldException);
      done();
    });

    it('should erase all spaces, end of lines and trailing periods', (done) => {
      ws.getCell('C3').value = '.temp    field\n header..';
      tempField = new Field('C3', null, ws);

      assert.equal(tempField.getName(), 'temp field header');
      done();
    });

    it('should correctly construct path to field if no parent provided', (done) => {
      assert.equal(tempField.getPath(), 'temp field header/');
      done();
    });

    it('should correctly construct path to field if parent provided', (done) => {
      tempField = new Field('C3', field, ws);
      assert.equal(tempField.getPath(), 'field header/temp field header/');
      done();
    });
  });
    
  describe('Field comparison', () => {    
    // Two equal tables
    const set1FieldNames = 
      ['f1_1val',  'f1_2val',  'f1_3val',  'f1_4val',  'f1_5val', 
      'f1_6val',  'f1_7val',  'f1_8val',  'f1_9val',  'f1_10val',
      'f1_11val', 'f1_12val', 'f1_13val', 'f1_14val', 'f1_15val', 'f1_16val'];

    const table1AddrRange = 
      ['D6', 'D7', 'F7', 'H7', 'D8', 'E8', 'F8', 'G8', 
       'H8', 'I8', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9'];

    const table2AddrRange = 
      ['L6', 'L7', 'N7', 'P7', 'L8', 'M8', 'N8', 'O8', 
       'P8', 'Q8', 'L9', 'M9', 'N9', 'O9', 'P9', 'Q9'];

    // Table three
    const table3AddrRange = 
      ['L12', 'L13', 'N13', 'P13', 'L14', 'M14', 'N14', 'O14', 
       'P14', 'Q14', 'L15', 'M15', 'N15', 'O15', 'P15', 'Q15'];

    const set2FieldNames = 
      ['f1_1val',  'bla',  'f1_3val',  'f1_4val',  'f1_5val', 
       'f1_6val',  'f1_7val',  'f1_8val',  'f1_9val',  'f1_10val',
       'f1_11val', 'f1_12val', 'bla bla', 'f1_14val', 'f1_15val', 'f1_16val'];
    let field1, field2, field3;
    
    beforeEach((done) => {
      // Two equal tables
      set1FieldNames.forEach((name, i) => {
        ws.getCell(table1AddrRange[i]).value = name;
        ws.getCell(table2AddrRange[i]).value = name;        
      });      

      field1 = new Field(table1AddrRange[0], null, ws);
      field2 = new Field(table2AddrRange[0], null, ws);

      table1AddrRange.slice(1).forEach((adr, i) => {
        field1._addChild(adr);
      });

      table2AddrRange.slice(1).forEach((adr, i) => {
        field2._addChild(adr);
      });

      // Table three
      set2FieldNames.forEach((name, i) => {
        ws.getCell(table3AddrRange[i]).value = name;
      });

      field3 = new Field(table3AddrRange[0], null, ws);

      table3AddrRange.slice(1).forEach((adr, i) => {
        field3._addChild(adr);
      });

      done();
    });

    // it('should compare two equal fields with subfields and return true', (done) => {
    //   assert.equal(field1.compare(field2), true);
    //   done();
    // });

    // it('should compare two unequal fields with subfields and return false', (done) => {
    //   assert.equal(field1.compare(field3), false);
    //   done();
    // });
  });

  describe('Subfield construction', () => {
    const wsPromise = readFile('C:/Users/ShaytanovAI/Documents/work/xlutils/test/field_test.xlsx');
    const f1 = 'D6', f2 = 'L6', f3 = 'L12';

    it ('should correctly construct field with subfields', (done) => {
      wsPromise.then(({ ws }) => {
        const field1 = new Field(f1, null, ws, 3);
        const field2 = new Field(f2, null, ws, 3);
        const field3 = new Field(f3, null, ws, 3);
                
        done();
      });
    });

    // TO DELETE
    // it('should return the value of the field by the given row', (done) => {
    //   wsPromise.then(({ ws }) => {
    //     const field1 = new Field(f1, null, ws, 3);
    //     const field2 = field1.getChild('field1/field2/');
    //     const field3 = field1.getChild('field1/field3/');
    //     const field6 = field2.getChild('field1/field2/field6/');
    //     const field8 = field3.getChild('field1/field3/field8/');
    //     assert.equal(field1.valueAt(13), 'value');
    //     assert.equal(field2.valueAt(13), 'value');
    //     assert.equal(field6.valueAt(13), null);
    //     assert.equal(field8.valueAt(11), 'value2');
    //     assert.equal(field3.valueAt(11), null);
    //     done();
    //   });
    // });
  });

  it('should throw AddressFormatException if invalid address provided on child addition', (done) => {
    assert.throws(() => field._addChild('J-0'), AddressFormatException);
    done();
  });

  it('should correctly add child', (done) => {
    const childName = ws.getCell(childAddress).value;
    field._addChild(childAddress);
    const tempChild = new Field(childAddress, field, ws);
    const childPath = StringUtils.constructPath(field.getPath(), childName);
    assert.equal(field.getChild(childPath).getPath(), tempChild.getPath());
    assert.equal(field.getChild(childPath).getParent().getName(), tempChild.getParent().getName());
    done();
  });

  it('should return `true` if child by the given name exists and false if not', (done) => {    
    field._addChild(childAddress);
    const childName = ws.getCell(childAddress).value;
    const childPath = StringUtils.constructPath(field.getPath(), childName);
    assert.equal(field.containsChild(childPath), true);
    assert.equal(field.containsChild('invalid child name'), false);
    done();
  });

  it('should return path of the given field', (done) => {
    field._addChild(childAddress);
    assert.equal(field.getPath(), 'field header/');
    assert.equal(field.getChild('field header/child field header/').getPath(), 'field header/child field header/');
    done();
  });

  it('should return name of the given field', (done) => {    
    assert.equal(field.getName(), 'field header');
    done();
  });

  it('should return parent of a given field', (done) => {
    field._addChild(childAddress);
    assert.equal(field.getParent(), null);
    assert.deepEqual(field.getChild(field.getPath() + 'child field header/').getParent().getName(), field.getName());
    done();
  });
});