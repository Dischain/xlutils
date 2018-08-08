'use strict';

const CommonAddress = require('../address/CommonAddress'),
      REUtils = require('../utils/REUtils');

const _storage = new WeakMap();
      
class Cell {
	constructor(address, workSheet, validate) {
		if (workSheet.getCell(address) == null || workSheet.getCell(address) == undefined) {
			workSheet.getCell(address) = {};
			workSheet.getCell(address).value = '';
			console.log('empty cell created by address ' + address);
		}
		this._address = new CommonAddress(address, validate);		
		this._ws = workSheet;
		_storage.set(this, workSheet.getCell(address).value);
	}

	offset (row, col) { 
		const curColIndex = REUtils.colToIndex(this.getColumn());
		const curRow = this.getRow();
		
		const newRow = curRow + row;
		const newColIndex = curColIndex + col;

		const newCol = REUtils.indexToCol(newColIndex);

		this._address.setRow(newRow);
		this._address.setColumn(newCol);
		
		_storage.set(this, this._ws.getCell(this._address.toString()).value);
		//this._value = this._cell.value;
	}

	getAddress() { return this._address; }

	setAddress(address, validate) {
		this._address = new CommonAddress(address, validate);
		_storage.set(this, this._ws.getCell(this._address.toString()).value);
	}

	getCell() { return this._ws.getCell(this._address.toString()); }

	getValue() { return _storage.get(this); }
	// getValue() { 
	// 	//console.log('getting ' + this._ws.getCell(this._address.toString()).value);
	// 	return this._ws.getCell(this._address.toString()).value; 
	// }

	setValue(value) {
		this._ws.getCell(this._address.toString()).value = value;		
		//this._value = this._cell.value;
		_storage.set(this, this._ws.getCell(this._address).value);
	}

	getColumn() { return this._address.getColumn(); }

	getColumnIndex() { return this._address.getColumnIndex(); }

	getRow() { return this._address.getRow(); }
}

module.exports = Cell;