'use strict';

const Address = require('./Address'),
      REUtils = require('../utils/REUtils'),
      AddressFormatException = require('../exceptions/AddressFormatException');

class RangeAddress extends Address{ 
  constructor(address, validate) {
    super(address, validate);
    if (validate) {
			if (REUtils.isRangeAddress(address)) {				
				this._address = address;
			} else {
				throw new AddressFormatException(address);
			}
		} else {
			this._address = address;
		}
  }

  getUpperBound() {
    return REUtils.getUpperBound(this._address);
  }

  getLowerBound() {
    return REUtils.getLowerBound(this._address);
  }
  
  setUpperBound(uBound, validate) {
    if (validate) {
      if (REUtils.isAddress(uBound)) {
        this._address = REUtils.setUpperBound(this._address, uBound);
      } else {
        throw new AddressFormatException(uBound);
      }
    }	else {
      this._address = REUtils.setUpperBound(this._address, uBound);
    }
  }
  
  setLowerBound(lBound, validate) {
    if (validate) {
      if (REUtils.isAddress(lBound)) {
        this._address = REUtils.setLowerBound(this._address, lBound);
      } else {
        throw new AddressFormatException(lBound);
      }
    } else {
      this._address = REUtils.setLowerBound(this._address, lBound);
    }
  }
  
    // cb(cellAddress)
  forEach(cb) {    
    const uBound = this.getUpperBound(),
          lBound = this.getLowerBound();
  
    const uBoundColIndex = REUtils.colToIndex(REUtils.getColumn(uBound)),
          lBoundColIndex = REUtils.colToIndex(REUtils.getColumn(lBound)),
          uBoundRow = REUtils.getRow(uBound),
          lBoundRow = REUtils.getRow(lBound);
  
    const colRange = Math.abs(uBoundColIndex - lBoundColIndex),
          rowRange = Math.abs(uBoundRow - lBoundRow);
  
    let mostLeftCol, mostRightCol, upperRow, lowerRow;
  
    if (uBoundColIndex > lBoundColIndex) { 
      mostLeftCol = lBoundColIndex; 
      mostRightCol = uBoundColIndex; 
    } else {
      mostLeftCol = uBoundColIndex; 
      mostRightCol = lBoundColIndex; 			
    }
  
    if (uBoundRow > lBoundRow) {
      upperRow = uBoundRow;
      lowerRow = lBoundRow;
    } else {
      upperRow = lBoundRow;
      lowerRow = uBoundRow;
    }
    
    // vertical range
    if (colRange == 0 && rowRange > 0) {
      let counter = 0, col = REUtils.getColumn(uBound);

      for (let i = lowerRow; i <= upperRow; i ++) {
        cb(col + i, counter);
        counter += 1;
      }
    }

    // horizontal range
    if (rowRange == 0 && colRange > 0) {
      let counter = 0, row = REUtils.getRow(uBound);

      for (let i = mostLeftCol; i <= mostRightCol; i ++) {
        cb(REUtils.indexToCol(i) + row, counter);
        counter += 1;
      }
    }

    // 2-dimensional range
    if (colRange > 0 && rowRange > 0) {      
      let counter = 0;

      for (let i = lowerRow; i <= upperRow; i ++) {
        for (let j = mostLeftCol; j <= mostRightCol; j ++) {
          cb(REUtils.indexToCol(j) + i, counter);
          counter += 1;
        }
      }
    }
  }
}

module.exports = RangeAddress;