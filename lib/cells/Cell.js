'use strict';

const CommonAddress = require('../address/CommonAddress'),
      REUtils = require('../utils/REUtils');

const _storage = new WeakMap();

function storage(obj) {
	if (!_storage.has(obj)) {
		_storage.set(obj, Object.create(null));
	}

	return _storage.get(obj);
}

class Cell {
	constructor(address, workSheet, validate) {
		storage(this).address = new CommonAddress(address, validate);
		storage(this).cell = workSheet.getCell(address);
		storage(this).ws = workSheet;
	}

	offset (row, col) { 
		const curColIndex = REUtils.colToIndex(this.getColumn());
		const curRow = this.getRow();
		
		const newRow = curRow + row;
		const newColIndex = curColIndex + col;

		const newCol = REUtils.indexToCol(newColIndex);

		storage(this).address.setRow(newRow);
		storage(this).address.setColumn(newCol);

		storage(this).cell = storage(this).ws.getCell(storage(this).address.toString());
	}

	getAddress() { return storage(this).address; }

	/**
	 * 
	 * @param {String} address 
	 */
	setAddress(address, validate) {
		storage(this).address = new CommonAddress(address, validate);
		storage(this).cell = storage(this).ws.getCell(storage(this).address);
	}

	getCell() { return storage(this).cell; }

	getValue() { return storage(this).cell.value; }

	setValue(value) { storage(this).cell.value = value; }

	getColumn() { return storage(this).address.getColumn(); }

	getColumnIndex() { return storage(this).address.getColumnIndex(); }

	getRow() { return storage(this).address.getRow(); }
}

module.exports = Cell;