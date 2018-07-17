'use strict';

const assert = require('assert'),
      CommonAddress = require('../../lib/address/CommonAddress'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException'), 
      ColumnFormatException = require('../../lib/exceptions/ColumnFormatException'),
      RowFormatException = require('../../lib/exceptions/RowFormatException');

describe('CommonAddress', () => {
  let commonAddress;

  beforeEach((done) => {
    commonAddress = new CommonAddress('AA100', true);
    done();
  });

  it('should throw AddressFormatException when invalid address provided on creation', (done) => {
    assert.throws(() => new CommonAddress('1Z', true), AddressFormatException);
    done();
  });

  it('should return row number', (done) => {
    assert.equal(commonAddress.getRow(), 100);
    done();
  });

  it('should return column value', (done) => {
    assert.equal(commonAddress.getColumn(), 'AA');
    done();
  });

  it('should return column index', (done) => {
    assert.equal(commonAddress.getColumnIndex(), '27');
    done();
  });

  it('should set row number', (done) => {
    commonAddress.setRow(2);
    assert.equal(commonAddress.getRow(), 2);
    done();
  });
  
  it('should throw RowFormatException when row is invalid', (done) => {
    assert.throws(() => commonAddress.setRow('a', true), RowFormatException);
    done();
  });

  it('should set column calue', (done) => {
    commonAddress.setColumn('Z');
    assert.equal(commonAddress.getColumn(), 'Z');
    done();
  });
  
  it('should throw ColumnFormatException when row is invalid', (done) => {
    assert.throws(() => commonAddress.setColumn('1', true), ColumnFormatException);
    done();
  });
});