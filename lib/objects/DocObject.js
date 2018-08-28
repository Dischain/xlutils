'use strict';

const Row = require('../items/Row'),
      Cell = require('../cells/Cell'),
      Range = require('../cells/Range'),
      REUtils = require('../utils/REUtils'),
      StringUtils = require('../utils/StringUtils');

// If used to construct sub documents for the document, which attached to some
// parent, api caller should specify the `parentSigns` to properly stop search of
// children limit. In other case specifying only the self color- `myGF` should be 
// enough.
class DocObject {
  constructor(nameAddress, valuesRange, ws, objectsRange, parentSigns, childSigns, mySigns) {
    this.row = new Row(nameAddress, REUtils.fillColRangeWithRows(valuesRange, REUtils.getRow(nameAddress)), ws);
    this._ws = ws;
    this._valuesRange = valuesRange;
    this._objectsRange = objectsRange;
    
    if (mySigns != null) {
      this._mySigns = mySigns;
    } else {
      this._mySigns = [];
    }

    if (parentSigns != null) {
      this._parentSigns = parentSigns;
    } else {
      this._parentSigns = [];
    }
    
    this._isPlain = false;
    if ((parentSigns == undefined || parentSigns == null) && 
        (childSigns == undefined || childSigns == null) && 
        (mySigns == undefined || mySigns == null))
    {
      this._isPlain = true;
    }
    // if ((parentSigns != undefined || parentSigns != null) && objectsRange != undefined) {
    //   this._setParent(parentSigns, objectsRange);
    // }
    
    if ((childSigns != undefined || childSigns != null)  && objectsRange != undefined) {
      this._children = {};
      this._collectChildren(childSigns, objectsRange);
    }
  }

  _setPath(parentPath) { 
    this._path = StringUtils.constructPath(parentPath, this.row.getName()); 
  }

  getPath() { return this._path; }

  getChildren() { return this._children; }

  getChildrenRange() { return this._childrenRange; }

  _collectChildren(childSigns, objectsRange) {
    const childrenLimit = this._getChildrenLimit(),
          childrenRange = REUtils.offsetRow(this.row.getAddress(), 1)
                        + ':' + childrenLimit;

    if (childrenLimit == undefined) return;

    const range = new Range(childrenRange, this._ws);
    this._childrenRange = range;

    range.forEach((cell) => {
      let childName = StringUtils.eraseAll(cell.getValue());
      
      if (childSigns == 'empty' && /*cell.getCell().fill != undefined*/
          (StringUtils.containsArrItem(cell.getValue(), this._parentSigns) || 
          StringUtils.containsArrItem(cell.getValue(), this._mySigns))) return;

      if (childSigns == 'empty' && 
        (!StringUtils.containsArrItem(cell.getValue(), this._parentSigns) || 
        !StringUtils.containsArrItem(cell.getValue(), this._mySigns))) 
      {
        this._children[childName] = 
          new Row(cell.getAddress().toString(), REUtils.fillColRangeWithRows(this._valuesRange, cell.getRow()), this._ws);
          return;
      }

      if (StringUtils.containsArrItem(cell.getValue(), childSigns))
      /*if (cell.getCell().fill != undefined && childSigns != 'empty') {
        if (cell.getCell().fill.fgColor.tint == childSigns.tint || 
            cell.getCell().fill.fgColor.argb == childSigns.argb) */
      { 
        this._children[childName] = 
          new Row(cell.getAddress().toString(), REUtils.fillColRangeWithRows(this._valuesRange, cell.getRow()), this._ws);
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

      let containsParent = StringUtils.containsArrItem(cell.value, this._parentSigns),
          containsMy = StringUtils.containsArrItem(cell.value, this._mySigns);

      if (containsParent || containsMy) {
        return REUtils.offsetRow(cell.address, -1);
      }
    }
  }
}

module.exports = DocObject;