'use strict';

class AddressFormatException extends Error {
  constructor(address) {
    super();    
    this.message = 'Invalid address format: "' + address + '"';
  }
}

module.exports = AddressFormatException;