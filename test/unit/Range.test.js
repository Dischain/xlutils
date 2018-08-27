'use strict';

const assert = require('assert'),
      Range = require('../../lib/cells/Range'),
      Cell = require('../../lib/cells/Cell'),
      createWS = require('../../lib/utils/xl').createWS,
      RangeAddress = require('../../lib/address/RangeAddress'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException');

describe('Range', () => {
  const { ws } = createWS('test');
  let range;

  beforeEach((done) => {
    range = new Range('A1:B3', ws, true);
    done();
  });

  it('should throw AddressFormatException when invalid range address provided on range creation', (done) => {
    assert.throws(() => new Range('A1:2', ws, true), AddressFormatException);
    done();
  });

  it('should return RangeAddress instance when getting address', (done) => {
    assert.deepEqual(range.getAddress(), new RangeAddress('A1:B3'));    
    done();
  });

  it('should traverse square range', (done) => {
    let resultArr = [new Cell('A1', ws), new Cell('B1', ws), new Cell('A2', ws), 
                     new Cell('B2', ws), new Cell('A3', ws), new Cell('B3', ws)];
    let temp = [];

    range.forEach((cell, i) => { temp.push(cell); });
    assert.deepEqual(temp, resultArr);
    done();
  });

  it('should return upper bound', (done) => {
    assert.equal(range.getUpperBound(), 'B3');
    done();
  });

  it('should return lower bound', (done) => {
    assert.equal(range.getLowerBound(), 'A1');
    done();
  });

  it('should set upper bound', (done) => {
    range.setUpperBound('ZQ11');
    assert.equal(range.getUpperBound(), 'ZQ11');
    done();
  });

  it('should throw AddressFormatException when upper bound is invalid', (done) => {
    assert.throws(() => range.setUpperBound('ZQ', true), AddressFormatException);
    done();
  });

  it('should set lower bound', (done) => {
    range.setLowerBound('ARQ22');
    assert.equal(range.getLowerBound(), 'ARQ22');
    done();
  });

  it('should throw AddressFormatException when lower bound is invalid', (done) => {
    assert.throws(() => range.setLowerBound('12', true), AddressFormatException);
    done();
  });
});