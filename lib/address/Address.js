'use strict';

const REUtils = require('../utils/REUtils'),
			AddressFormatException = require('../exceptions/AddressFormatException');

class Address {
	constructor(address, validate) {
		if (validate) {
			if (REUtils.isAddress(address)) {				
				this._address = address;
			} else {
				throw new AddressFormatException(address);
			}
		} else {
			this._address = address;
		}
	}

	toString() { return this._address; }
}

module.exports = Address;