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

  it('should construct path by the given parent path and child name', (done) => {
    assert.equal(StringUtils.constructPath('parent/path/', 'child'), 'parent/path/child/');
    assert.equal(StringUtils.constructPath(null, 'child'), 'child/');
    assert.equal(StringUtils.constructPath(undefined, 'child'), 'child/');
    done();
  });

  it('should return true if the given string contains one of specfied array item', (done) => {
    assert.equal(StringUtils.containsArrItem('Министерство чегото там', ['Министерство', 'комитет', 'служба', 'управление']), true);
    assert.equal(StringUtils.containsArrItem('комитет чегото там', ['министерство', 'комитет', 'служба', 'управление']), true);
    assert.equal(StringUtils.containsArrItem('управление чегото там', ['министерство', 'комитет', 'служба', 'управление']), true);
    assert.equal(StringUtils.containsArrItem('дирекция чегото там', ['министерство', 'комитет', 'служба', 'управление']), false);
    assert.equal(StringUtils.containsArrItem('дирекция чегото там', ['дирекция', 'комитет', 'служба', 'управление']), true);
    done();
  });
});