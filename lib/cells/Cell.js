'use strict';

const CommonAddress = require('../address/CommonAddress'),
      REUtils = require('../utils/REUtils');

class Cell {
	constructor(address, workSheet) {
		this._address = new CommonAddress(address);
		this._cell = workSheet.getCell(address);
	}

	offset (row, col) { 
		let newRow = this._cell.getRow() + row;
		
		let curColIndex = REUtils.colToIndex(this._cell.getCollumn());
		let newColIndex = curColIndex + col;
		let newCol = REUtils.indexToCol(newColIndex);

		this._address.setRow(newRow);
		this._address.setCollumn(newCol);

		this._cell = workSheet.getCell(this._address);
	}

	getCollumn() { return address.getCollumn(); }
	getRow() { return address.getRow(); }
}