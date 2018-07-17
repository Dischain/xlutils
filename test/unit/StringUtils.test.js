'use strict';
const assert = require('assert'),
      StringUtils = require('../../lib/utils/StringUtils');

describe('SringUtils', () => {
  it('should erase spaces from given string', (done) => {
    assert.equal(StringUtils.eraseSPs('mm   mmm da   ooo o'), 'mm mmm da ooo o');
    assert.equal(StringUtils.eraseSPs('mm mmm da ooo o'), 'mm mmm da ooo o');
    done();
  });

  it('should erase end of lines from given string', (done) => {
    assert.equal(StringUtils.eraseEOLs('mm\nmmm\nda\nooo o'), 'mmmmmdaooo o');
    done();
  });

  it('should trim trailing periods from given string', (done) => {
    assert.equal(StringUtils.eraseTrailingPeriods('...dsds..'), 'dsds');
    assert.equal(StringUtils.eraseTrailingPeriods('.dsds.'), 'dsds');
    assert.equal(StringUtils.eraseTrailingPeriods('13dsds.'), '13dsds');
    assert.equal(StringUtils.eraseTrailingPeriods('.'), '');
    assert.equal(StringUtils.eraseTrailingPeriods('...1.3dsds'), '1.3dsds');
    done();
  });
});