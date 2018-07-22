'use strict';

class EmptyFieldException extends Error {
  constructor(address) {
    super();    
    this.message = 'Field can not be empty: "' + address + '"';
  }
}

module.exports = EmptyFieldException;