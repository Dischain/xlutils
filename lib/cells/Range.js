'use strict';

const RangeAddress = require('../address/RangeAddress'),
      Cell = require('./Cell'),
      REUtils = require('../utils/REUtils');

const _storage = new WeakMap();
      
function storage(obj) {
  if (!_storage.has(obj)) {
    _storage.set(obj, Object.create(null));
  }
  
  return _storage.get(obj);
}

class Range {
  constructor(address, workSheet, validate) {
    storage(this).address = new RangeAddress(address, validate);
    storage(this).ws = workSheet;
  }

  getAddress() { return storage(this).address; }

  getUpperBound() {
    return storage(this).address.getUpperBound();
  }

  getLowerBound() {
    return storage(this).address.getLowerBound();
  }

  setUpperBound(uBound, validate) {
    return storage(this).address.setUpperBound(uBound, validate);
  }

  setLowerBound(uBound, validate) {
    return storage(this).address.setLowerBound(uBound, validate);
  }

  forEach(cb) {
    const addr = storage(this).address,
          ws = storage(this).ws;
    addr.forEach((addr, i) => {
      cb(new Cell(addr, ws), i);
    });
  }
}

module.exports = Range;