'use strict';

const REUtils = require('../utils/REUtils'),
			AddressFormatException = require('../exceptions/AddressFormatException');

class Address {
	constructor(address, validate) {
		if (validate) {
			if (REUtils.isAddress(address)) {				
				this.address = address;
			} else {
				throw new AddressFormatException(address);
			}
		} else {
			this.address = address;
		}
	}

	toString() { return this.address; }
}

module.exports = Address;