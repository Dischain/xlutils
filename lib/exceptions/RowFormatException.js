'use strict';

class RowFormatException extends Error {
  constructor(row) {
    super();    
    this.message = 'Invalid row format: "' + row + '"';
  }
}

module.exports = RowFormatException;