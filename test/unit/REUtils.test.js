'use strict';

const assert = require('assert'),
      REUtils = require('../../lib/utils/REUtils');

describe('REUtils', () => {
  it('should test whether a given string is an Excel address', (done) => {
    assert.equal(REUtils.isAddress('A2'), true); // common address case
    assert.equal(REUtils.isAddress('z2'), true); // common address case, lowercase letter
    assert.equal(REUtils.isAddress('AZ23'), true); // common address case, big collumn index
    assert.equal(REUtils.isAddress('BT2:AK3'), true); // range address case
    
    assert.equal(REUtils.isAddress('1Z'), false); // common address case, with typo
    assert.equal(REUtils.isAddress('A,2'), false); // common address case, with typo
    assert.equal(REUtils.isAddress('A'), false); // common address case, without row
    assert.equal(REUtils.isAddress('2'), false); // common address case, with collumn
    assert.equal(REUtils.isAddress('BT2: AK3'), false); // range address case, with typo
    assert.equal(REUtils.isAddress('BT2:AK'), false); // range address case, without row
    assert.equal(REUtils.isAddress('2:AK3'), false); // range address case, without column
    assert.equal(REUtils.isAddress('BT:AK'), false); // range address case, without rowss
    done();
  });

  it('should test whether a given string is an Excel range address', (done) => {
    assert.equal(REUtils.isRangeAddress('A2'), false); // common address case
    assert.equal(REUtils.isRangeAddress('z2'), false); // common address case, lowercase letter
    assert.equal(REUtils.isRangeAddress('AZ23'), false); // common address case, big collumn index
    assert.equal(REUtils.isRangeAddress('BT2:AK3'), true); // range address case
    
    assert.equal(REUtils.isRangeAddress('1Z'), false); // common address case, with typo
    assert.equal(REUtils.isRangeAddress('A,2'), false); // common address case, with typo
    assert.equal(REUtils.isRangeAddress('A'), false); // common address case, without row
    assert.equal(REUtils.isRangeAddress('2'), false); // common address case, with collumn
    assert.equal(REUtils.isRangeAddress('BT2: AK3'), false); // range address case, with typo
    assert.equal(REUtils.isRangeAddress('BT2:AK'), false); // range address case, without row
    assert.equal(REUtils.isRangeAddress('2:AK3'), false); // range address case, without column
    assert.equal(REUtils.isRangeAddress('BT:AK'), false); // range address case, without rowss
    done();
  });

  it('should return the upper bound of range address', (done) => {
    assert.equal(REUtils.getUpperBound('BT2:AK3'), 'AK3');
    done();
  });

  it('should return the lower bound of range address', (done) => {
    assert.equal(REUtils.getLowerBound('BT2:AK3'), 'BT2');
    done();
  });

  it('should set the upper bound of range address', (done) => {
    assert.equal(REUtils.setUpperBound('BT2:AK3', 'GH88'), 'BT2:GH88');
    done();
  });

  it('should set the lower bound of range address', (done) => {
    assert.equal(REUtils.setLowerBound('BT2:AK3', 'GH88'), 'GH88:AK3');
    done();
  });

  it('should return row index of a given address', (done) => {
    assert.equal(REUtils.getRow('BT2'), 2);
    done();
  });

  it('should return collumn of a given address', (done) => {
    assert.equal(REUtils.getColumn('BT2'), 'BT');
    done();
  });

  it('should set row index of a common address', (done) => {
    assert.equal(REUtils.setRow('B2', 10), 'B10');
    done();
  });

  it('should set column of a common address', (done) => {
    assert.equal(REUtils.setColumn('B2', 'F'), 'F2');
    done();
  });

  it('should transform column numeric index to string value', (done) => {
    assert.equal(REUtils.indexToCol(26), 'Z');
    assert.equal(REUtils.indexToCol(51), 'AY');
    assert.equal(REUtils.indexToCol(52), 'AZ');
    assert.equal(REUtils.indexToCol(80), 'CB');
    assert.equal(REUtils.indexToCol(676), 'YZ');
    assert.equal(REUtils.indexToCol(702), 'ZZ');
    assert.equal(REUtils.indexToCol(705), 'AAC');
    done();
  });

  it('should transform column string value to numeric index', (done) => {
    assert.equal(REUtils.colToIndex('Z'), 26);
    assert.equal(REUtils.colToIndex('ay'), 51);
    assert.equal(REUtils.colToIndex('AZ'), 52);
    assert.equal(REUtils.colToIndex('CB'), 80);
    assert.equal(REUtils.colToIndex('yZ'), 676);
    assert.equal(REUtils.colToIndex('Zz'), 702);
    assert.equal(REUtils.colToIndex('AaC'), 705);
    done();
  });
});