'use strict';

const Row = require('../items/Row'),
      Cell = require('../cells/Cell'),
      Range = require('../cells/Range'),
      REUtils = require('../utils/REUtils'),
      StringUtils = require('../utils/StringUtils');

// If used to construct sub documents for the document, which attached to some
// parent, api caller should specify the `parentFG` to properly stop search of
// children limit. In other case specifying only the self color- `myGF` should be 
// enough.
class DocObject {
  constructor(nameAddress, valuesRange, ws, objectsRange, parentFG, childFG, myFG) {
    this.row = new Row(nameAddress, REUtils.fillColRangeWithRows(valuesRange, REUtils.getRow(nameAddress)), ws);
    this._ws = ws;
    this._valuesRange = valuesRange;
    this._objectsRange = objectsRange;
    this._myFG = myFG;
    this._parentFG = parentFG;

    // if ((parentFG != undefined || parentFG != null) && objectsRange != undefined) {
    //   this._setParent(parentFG, objectsRange);
    // }
    
    if ((childFG != undefined || childFG != null)  && objectsRange != undefined) {
      this._children = {};
      this._collectChildren(childFG, objectsRange);      
    }    
  }

  getParentCell() { return this._parent; }

  getChildren() { return this._children; }

  _collectChildren(childFG, objectsRange) {
    const childrenLimit = this._getChildrenLimit(),
          childrenRange = REUtils.offsetRow(this.row.getAddress(), 1)
                        + ':' + childrenLimit;
    
    if (childrenLimit == undefined) return;

    const range = new Range(childrenRange, this._ws);
    
    range.forEach((cell) => {
      if (childFG == 'empty' && cell.getCell().fill != undefined) return;
      if (childFG == 'empty' && cell.getCell().fill == undefined) {        
        this._children[StringUtils.eraseAll(cell.getValue())] = 
          new Row(cell.getAddress().toString(), REUtils.fillColRangeWithRows(this._valuesRange, cell.getRow()), this._ws);
          return;
      }
      
      if (cell.getCell().fill != undefined && childFG != 'empty') {
        if (cell.getCell().fill.fgColor.tint == childFG.tint || 
            cell.getCell().fill.fgColor.argb == childFG.argb) 
        {
          this._children[StringUtils.eraseAll(cell.getValue())] = 
            new Row(cell.getAddress().toString(), REUtils.fillColRangeWithRows(this._valuesRange, cell.getRow()), this._ws);
        }
      }
    });
  }

  _getChildrenLimit() {
    const maxRow = REUtils.getMaxRow(this._objectsRange),
          col = this.row.getCell().getColumn();

    for (let i = this.row.getRow() + 1; i <= maxRow; i ++) {
      const cell = this._ws.getCell(col + i);
      if (i == maxRow) {        
        return cell.address;
      }

      if (cell.fill != undefined) {
        if (this._parentFG != null) {
          if (cell.fill.fgColor.tint == this._parentFG.tint ||
              cell.fill.fgColor.argb == this._parentFG.argb) {
            return REUtils.offsetRow(cell.address, -1);
          }
        }
        if (cell.fill.fgColor.tint == this._myFG.tint || 
            cell.fill.fgColor.argb == this._myFG.argb) {
          return REUtils.offsetRow(cell.address, -1);
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
         (cell.fill.fgColor.argb == parentFG.argb ||
          cell.fill.fgColor.tint == parentFG.tint)) {
        
        this._parent = new Cell(cell.address, this._ws);
      }
    }
  }

  _setParentSimple(parent) {
    this._parent = new Cell(cell.address, this._ws);
  }
}

module.exports = DocObject;