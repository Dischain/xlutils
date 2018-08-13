'use strict';

const CommonAddress = require('../address/CommonAddress'),
      RangeAddress = require('../address/RangeAddress'),
      Cell = require('../cells/Cell'),
      Range = require('../cells/Range'),
      EmptyFieldException = require('../exceptions/EmptyFieldException'),
      StringUtils = require('../utils/StringUtils'),
      REUtils = require('../utils/REUtils');

// May throw AddressFormatException
class Row {
  constructor(nameAddress, range, ws) {
    this._cell = new Cell(nameAddress, ws, true);
    this._name = StringUtils.eraseAll(this._cell.getValue());
    this._valuesRange = new Range(range, ws, true);
    this._values = {};
    this._valuesRange.forEach((cell, i) => {
      this._values[cell.getColumnIndex()] = cell;
    });
    this._ws = ws;
  }

  getCell() { return this._cell; }

  getName() { return this._name; }

  setName(name) { this._name = name; }

  getValuesRange() { return this._valuesRange; }

  getValueByColIndex(colIndex) {
    return this._values[colIndex].getValue();
  }

  setValueByColIndex(colIndex, value) {
    if (this._values[colIndex] == undefined) {
      console.log(this._valuesRange.getAddress().toString());
      console.log('-->' + REUtils.indexToCol(colIndex) + this.getRow());
    }
    this._values[colIndex].setValue(value);
  }

  getAddress() { return this._cell.getAddress().toString(); }

  getRow() { return this._cell.getRow(); }
}

module.exports = Row;