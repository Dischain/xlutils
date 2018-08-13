'use strict';

const AddressFormatException = require('../exceptions/AddressFormatException');

const _addressRegex = /^[a-z]+\d+(:[a-z]+\d+)?$/i;
const _addressUpperBound = /[a-z]+\d+$/i;
const _addressLowerBound = /^[a-z]+\d+/i;
const _rangeAddressRegex = /^[a-z]+\d+:[a-z]+\d+$/i;

const _commonAddressRow = /\d+$/;
const _commonAddressColumn = /^[a-z]+/i;

const _rowRE = /[0-9]+/;
const _colRE = /[a-z]+/i;    

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

class REUtils {
  static isAddress(addr) { return _addressRegex.test(addr); }

  static isRangeAddress(addr) { return _rangeAddressRegex.test(addr); }

  static getUpperBound(addr) { 
    return _addressUpperBound.exec(addr)[0];
  }
	static getLowerBound(addr) {
    return _addressLowerBound.exec(addr)[0];
  }

	static setUpperBound(addr, uBound) { 
    return addr.replace(_addressUpperBound, uBound); 
  }

	static setLowerBound(addr, lBound) { 
    return addr.replace(_addressLowerBound, lBound); 
  }

	static getRow(addr) {
    return Number.parseInt(_commonAddressRow.exec(addr)[0]);
  }
	static getColumn(addr) {
    return  _commonAddressColumn.exec(addr)[0];
  }

	static setRow(addr, row) { 
    return addr.replace(_commonAddressRow, row); 
  }

  static offsetRow(addr, offset) {
    const curRow = this.getRow(addr);
    return this.setRow(addr, curRow + offset);
  }

  static getMinRow(range) {
    const rangeLBound = REUtils.getLowerBound(range),
          rangeUBound = REUtils.getUpperBound(range),
          lBoundRow = this.getRow(rangeLBound),
          uBoundRow = this.getRow(rangeUBound);

    if (lBoundRow > uBoundRow) {
      return uBoundRow;
    } else {
      return lBoundRow;
    }
  }

  static getMaxRow(range) {
    const rangeLBound = REUtils.getLowerBound(range),
          rangeUBound = REUtils.getUpperBound(range),
          lBoundRow = this.getRow(rangeLBound),
          uBoundRow = this.getRow(rangeUBound);

    if (lBoundRow < uBoundRow) {
      return uBoundRow;
    } else {
      return lBoundRow;
    }
  }

  static getMaxCol(range) {
    const rangeLBound = REUtils.getLowerBound(range),
          rangeUBound = REUtils.getUpperBound(range),
          lBoundCol = this.getColumn(rangeLBound),
          uBoundCol = this.getColumn(rangeUBound);
    
    const lBoundColIndex = this.colToIndex(lBoundCol),
          uBoundColIndex = this.colToIndex(uBoundCol);
    
    if (lBoundColIndex < uBoundColIndex) {
      return uBoundColIndex;
    } else {
      return lBoundColIndex;
    }
  }

  static getMinCol(range) {
    const rangeLBound = REUtils.getLowerBound(range),
          rangeUBound = REUtils.getUpperBound(range),
          lBoundCol = this.getColumn(rangeLBound),
          uBoundCol = this.getColumn(rangeUBound);

    const lBoundColIndex = this.colToIndex(lBoundCol),
          uBoundColIndex = this.colToIndex(uBoundCol);

    if (lBoundColIndex < uBoundColIndex) {
      return lBoundColIndex;
    } else {
      return uBoundColIndex;
    }
  }

	static setColumn(addr, col) { 
    return addr.replace(_commonAddressColumn, col); 
  }

  static offsetColumn(addr, offset) {
    const curCol = this.getColumn(addr),
          curColIndex = this.colToIndex(curCol),
          newCol = this.indexToCol(curColIndex + offset);
    
    return this.setColumn(addr, newCol);
  }

	static isRow(row) { return _rowRE.test(row); }
	static isCol(col) { return _colRE.test(col); }

  static fillColRangeWithRows(colRange, row) {
    const [ frst, last ] = colRange.split(':');
    return frst + row + ':' + last + row;
  }

	static colToIndex(col) {
    let colName = col.toUpperCase();
    let i, j, result = 0;

    for (i = 0, j = colName.length - 1; i < colName.length; i+= 1, j -= 1) {
      result += Math.pow(ALPHABET.length, j) * (ALPHABET.indexOf(colName[i]) + 1);
    }

    return result;
	}

	static indexToCol(i) {
		let colNum = i, colName = '';

		while(colNum > 0) {
			let rem = colNum % 26;

			if (rem == 0) {
				colName += 'Z';
				colNum = (colNum / 26) - 1;
			} else {
				colName += String.fromCharCode(97 + (rem - 1));
				colNum = Math.floor(colNum / 26);
			}
		}

		return colName.split('').reverse().join('').toUpperCase();
  }
  
  static rangeContainsAddr(range, addr) {
    const frst = this.getUpperBound(range),
          scnd = this.getLowerBound(range),
          minRow = this.getMinRow(range),
          maxRow = this.getMaxRow(range),
          maxColIndex = this.getMaxCol(range),
          minColIndex = this.getMinCol(range);
    
    const addrRow = this.getRow(addr),
          addrColIndex = this.colToIndex(this.getColumn(addr));

    return (minRow <= addrRow && minColIndex <=addrColIndex
        && maxRow >= addrRow && maxColIndex >= maxColIndex);    
  }
}

function _validateOrThrow(fun, addr) {
  const res = fun(addr);
  if (REUtils.isAddress(addr) && res != null) { return res[0]; }  
  else { throw new AddressFormatException(addr); }
}

module.exports = REUtils;