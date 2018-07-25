'use strict';

const assert = require('assert'),
      Field = require('../../lib/items/Field'),
      Row = require('../../lib/items/Row'),
      Cell = require('../../lib/cells/Cell'),
      createWS = require('../../lib/utils/xl').createWS,
      readFile = require('../../lib/utils/xl').readFile,
      StringUtils = require('../../lib/utils/StringUtils'),
      AddressFormatException = require('../../lib/exceptions/AddressFormatException');

describe('Row', () => {
  const wsPromise = readFile('H:/работа/xlutils/test/row_test.xlsx');

  it('should correctly construct row', (done) => {
    wsPromise.then((ws) => {
      const row = new Row('C10', 'D10:I10', ws);
      assert.equal(row.getName(), 'row1');
      assert.equal(row.getAddress(), 'C10');
      assert.equal(row.getRow(), 10);
      assert.equal(row.getValuesRange().getAddress().toString(), 'D10:I10');
      done();
    });
  });

  it('should get value by column index', (done) => {
    let row;
    wsPromise.then((ws) => {
      row = new Row('C10', 'D10:I10', ws);
      return Promise.resolve();
    }).then(() => {
      assert.equal(row.getValueByColIndex(6), 'value3');
      done();
    });
  });

  it('should set value by column index', (done) => {
    let row;
    wsPromise.then((ws) => {
      row = new Row('C10', 'D10:I10', ws);
      return Promise.resolve();
    }).then(() => {
      row.setValueByColIndex(6, 'test');
      assert.equal(row.getValueByColIndex(6), 'test');
      done();
    });
  });
});