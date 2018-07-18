'use strict';

const Address = require('./Address'),
      REUtils = require('../utils/REUtils'),
      ColumnFormatException = require('../exceptions/ColumnFormatException'),
      RowFormatException = require('../exceptions/RowFormatException');

class CommonAddress extends Address {
  constructor(address, validate) {
    super(address, validate);
  }

  getRow() { 
		const strRow = REUtils.getRow(this._address); 
		return Number(strRow);
  }

	getColumn() { 
    return REUtils.getColumn(this._address); 
  }

  getColumnIndex() {
    const col = this.getColumn();
    return REUtils.colToIndex(col);
  }

	setRow(row, validate) { 
		if (validate) {
			if (REUtils.isRow(row)) {
				this._address = REUtils.setRow(this._address, row); 
			} else {
				throw new RowFormatException(row);
			}
		} else {
			this._address = REUtils.setRow(this._address, row); 
		}
	}

	setColumn(col, validate) { 
		if (validate) {
			if (REUtils.isCol(col)) {
				this._address = REUtils.setColumn(this._address, col); 
			} else {
				throw new ColumnFormatException(col);
			}
		} else {
			this._address = REUtils.setColumn(this._address, col);  
		}
	}	
}

module.exports = CommonAddress;