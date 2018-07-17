'use strict';

class ColumnFormatException extends Error {
  constructor(col) {
    super();    
    this.message = 'Invalid column format: "' + col + '"';
  }
}

module.exports = ColumnFormatException;