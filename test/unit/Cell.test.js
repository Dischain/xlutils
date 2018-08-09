'use strict';

const assert = require('assert'),
      Cell = require('../../lib/cells/Cell'),
      createWS = require('../../lib/utils/xl').createWS,
      CommonAddress = require('../../lib/address/CommonAddress'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException');

describe('Cell', () => {
  const ws = createWS('test');
  let cell;

  beforeEach((done) => {
    cell = new Cell('A1', ws, true);
    done();
  });

  it('should throw AddressFormatException when invalid address provided on cell creation', (done) => {
    assert.throws(() => new Cell('A-1', ws, true), AddressFormatException);
    done();
  });

  it('should correctly offset given cell', (done) => {    
    cell.offset(1, 0);
    assert.equal(cell.getAddress().toString(), 'A2');
    cell.offset(0, 1);
    assert.equal(cell.getAddress().toString(), 'B2');
    cell.offset(1, 1);
    assert.equal(cell.getAddress().toString(), 'C3');
    cell.offset(-1, -1);
    assert.equal(cell.getAddress().toString(), 'B2');
    cell.offset(0, -1);
    assert.equal(cell.getAddress().toString(), 'A2');
    cell.offset(-1, 0);
    assert.equal(cell.getAddress().toString(), 'A1');
    done();
  });

  it('should return CommonAddress instance when getting address', (done) => {
    assert.deepEqual(cell.getAddress(), new CommonAddress('A1'));    
    done();
  });

  it('should set new address or throw AddressFormatException', (done) => {
    cell.setAddress('B2');
    assert.deepEqual(cell.getAddress(), new CommonAddress('B2'));    
    assert.throws(() => cell.setAddress('B-2', true), AddressFormatException);
    done();
  });

  it('should return inner cell', (done) => {
    assert.notEqual(cell.getCell(), null);
    done();
  });

  it('should return inner cell value', (done) => {
    assert.equal(cell.getValue(), '');
    done();
  });

  it('should set inner cell value', (done) => {
    cell.setValue('test');
    assert.equal(cell.getValue(), 'test');
    done();
  });

  it('should return inner cell column', (done) => {
    assert.equal(cell.getColumn(), 'A');
    done();
  });

  it('should return inner cell column index', (done) => {
    assert.equal(cell.getColumnIndex(), 1);
    done();
  });

  it('should return inner cell row', (done) => {
    assert.equal(cell.getRow(), 1);
    done();
  });
});