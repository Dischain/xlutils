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
    return REUtils.getRow(this.address); 
  }

	getColumn() { 
    return REUtils.getColumn(this.address); 
  }

  getColumnIndex() {
    const col = this.getColumn();
    return REUtils.colToIndex(col);
  }

	setRow(row, validate) { 
		if (validate) {
			if (REUtils.isRow(row)) {
				this.address = REUtils.setRow(this.address, row); 
			} else {
				throw new RowFormatException(row);
			}
		} else {
			this.address = REUtils.setRow(this.address, row); 
		}
	}

	setColumn(col, validate) { 
		if (validate) {
			if (REUtils.isCol(col)) {
				this.address = REUtils.setColumn(this.address, col); 
			} else {
				throw new ColumnFormatException(col);
			}
		} else {
			this.address = REUtils.setColumn(this.address, col);  
		}
	}
}

module.exports = CommonAddress;