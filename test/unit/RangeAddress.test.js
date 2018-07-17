'use strict';

const assert = require('assert'),
      RangeAddress = require('../../lib/address/RangeAddress'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException');

describe('RangeAddress', () => {
  let vertAddress, horAddress, squareAddress, simpleAddress;

  beforeEach((done) => {
    vertAddress = new RangeAddress('AA100:AA103', true);
    horAddress = new RangeAddress('AA100:AD100', true);
    squareAddress = new RangeAddress('AA100:AD105', true);
    simpleAddress = new RangeAddress('ZT21:DK12', true);
    done();
  });

  it('should throw AddressFormatException when invalid address provided on creation', (done) => {
    assert.throws(() => new RangeAddress('1Z:J', true), AddressFormatException);
    done();
  });

  it('should return upper bound', (done) => {
    assert.equal(simpleAddress.getUpperBound(), 'DK12');
    done();
  });

  it('should return lower bound', (done) => {
    assert.equal(simpleAddress.getLowerBound(), 'ZT21');
    done();
  });

  it('should set upper bound', (done) => {
    simpleAddress.setUpperBound('ZQ11');
    assert.equal(simpleAddress.getUpperBound(), 'ZQ11');
    done();
  });

  it('should throw AddressFormatException when upper bound is invalid', (done) => {
    assert.throws(() => simpleAddress.setUpperBound('ZQ', true), AddressFormatException);
    done();
  });

  it('should set lower bound', (done) => {
    simpleAddress.setLowerBound('ARQ22');
    assert.equal(simpleAddress.getLowerBound(), 'ARQ22');
    done();
  });

  it('should throw AddressFormatException when lower bound is invalid', (done) => {
    assert.throws(() => simpleAddress.setLowerBound('12', true), AddressFormatException);
    done();
  });

  it('should traverse vertical range', (done) => {
    let resultArr = ['AA100', 'AA101', 'AA102', 'AA103'];
    let temp = [];

    vertAddress.forEach((cellAddr) => { temp.push(cellAddr); });
    assert.deepEqual(temp, resultArr);
    done();
  });

  it('should traverse horizontal range', (done) => {
    let resultArr = ['AA100', 'AB100', 'AC100', 'AD100'];
    let temp = [];

    horAddress.forEach((cellAddr) => { temp.push(cellAddr); });
    assert.deepEqual(temp, resultArr);
    done();
  });

  it('should traverse square range', (done) => {
    let resultArr = ['AA100', 'AB100', 'AC100', 'AD100', 
                    'AA101', 'AB101', 'AC101', 'AD101', 
                    'AA102', 'AB102', 'AC102', 'AD102', 
                    'AA103', 'AB103', 'AC103', 'AD103',
                    'AA104', 'AB104', 'AC104', 'AD104',
                    'AA105', 'AB105', 'AC105', 'AD105'];
    let temp = [];

    squareAddress.forEach((cellAddr, i) => { temp.push(cellAddr); });
    assert.deepEqual(temp, resultArr);
    done();
  });
});