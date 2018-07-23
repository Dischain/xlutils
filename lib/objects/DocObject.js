'use strict';

const Row = require('../items/Row'),
      REUtils = require('../utils/REUtils');

class DocObject {
  constructor(nameAddress, valuesRange, ws, objectsRange, parentFG, childFG) {
    this.row = new Row(nameAddress, valuesRange, ws);
    this._ws = ws;
    this._valuesRange = valuesRange;
    this._objectsRange = objectsRange;
    
    if ((parentFG != undefined || parentFG != null) && objectsRange != undefined) {
      this._setParent(parentFG, objectsRange);
    }

    if ((childFG != undefined || childFG != null)  && objectsRange != undefined) {
      this._children = {};
      this._collectChildren(childFG, objectsRange);
    }
  }

  getParentCell() { return this._parent; }

  _collectChildren(childFG, objectsRange) {
    const maxRow = REUtils.getMaxRow(objectsRange),
          col = this.row.getCell().getColumn();

    for (let i = this.row.getRow() + 1; i <= maxRow; i ++) {
      const cell = this._ws.getCell(col + i);
      if (cell.fill != undefined) break;

      if ((cell.fill == undefined && childFG == 'empty') /*|| 
          (cell.fill.fgColor.argb == childFG.argb ||
          cell.fill.fgColor.tint == childFG.tint)*/) {
        if (this._children[cell.value] != undefined) {
          console.log('already exists: ' + col + i + ' ' + cell.value);
          continue;
        }

        if (cell.value != undefined && cell.value != null) {
          // const childObj = 
          //   new DocObject(cell.address, this._valuesRange, this._ws, this._objectsRange, )
          this._children[cell.value] = cell;
        }
      }
    }
  }

  _setParent(parentFG, objectsRange) {
    const minRow = REUtils.getMinRow(objectsRange),
          col = this.row.getCell().getColumn();
    
    for (let i = this.row.getRow() - 1; i >= minRow; i --) {
      const cell = this._ws.getCell(col + i);
      if (cell.fill != undefined && 
         (cell.fill.fgColor.argb == parentFG.argb &&
          cell.fill.fgColor.tint == parentFG.tint)) {
        
        this._parent = cell;
      }
    }
  }
}

module.exports = DocObject;