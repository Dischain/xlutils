'use strict';

const RangeAddress = require('../address/RangeAddress'),
      Cell = require('./Cell'),
      REUtils = require('../utils/REUtils');

class Range {
  constructor(address, workSheet, validate) {
    this._ws = workSheet;
    this._address = new RangeAddress(address, validate);
    this._cells = [];
    this._address.forEach((addr, i) => {
      this._cells.push(new Cell(addr, workSheet));
    });    
  }

  getAddress() { return this._address; }

  getUpperBound() {
    return this._address.getUpperBound();
  }

  getLowerBound() {
    return this._address.getLowerBound();
  }

  setUpperBound(uBound, validate) {
    return this._address.setUpperBound(uBound, validate);
  }

  setLowerBound(uBound, validate) {
    return this._address.setLowerBound(uBound, validate);
  }

  // forEach(cb) {
  //   const addr = this._address,
  //         ws = this._ws;
  //   addr.forEach((addr, i) => {
  //     cb(new Cell(addr, ws), i);
  //   });
  // }

  forEach(cb) {
    const addr = this._address,
          ws = this._ws;
    this._cells.forEach((cell, i) => {
      cb(cell, i);
    });
  }
}

module.exports = Range;