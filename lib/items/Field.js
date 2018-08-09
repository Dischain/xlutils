'use strict';

const CommonAddress = require('../address/CommonAddress'),
      Cell = require('../cells/Cell'),
      Range = require('../cells/Range'),
      EmptyFieldException = require('../exceptions/EmptyFieldException'),
      StringUtils = require('../utils/StringUtils'),
      REUtils = require('../utils/REUtils');

// May throw AddressFormatException, EmptyFieldException
class Field {
  constructor(address, parent, ws, height) {
    this._cell = new Cell(address, ws, true);

    if (this._cell.getValue() == '') {
      throw new EmptyFieldException(address);
    }

    this._name = StringUtils.eraseAll(this._cell.getValue());
    this._parent = parent;
    this._ws = ws;
    if (parent == null || parent == undefined) {
      this._path = StringUtils.constructPath(null, this._name);
    } else {
      this._path = StringUtils.constructPath(parent.getPath(), this._name);
    }    
    this._children = {};
    this._lowerFields = {};
    if (height != undefined && height != 0) {      
      this._constructWithSubfields(height);
      this._height = height;
    } else {
      this._height = 0;
      this._lowerFields[this._path] = this;
    }
  }

  // TO DELETE
  // valueAt(row) {
  //   const cellAddr = REUtils.setRow(this.getCell().getAddress().toString(), row);  
  //   return this._ws.getCell(cellAddr).value;
  // }

  equals(anotherField) {
    if (this._height != anotherField.getHeight()) return;
    for (let k in this._lowerFields) {
      if (anotherField.getLowerFields()[k] == undefined)
        return false;
    }
    return true;
  }

  diff(anotherField) {
    const diff = {};
        
    for (let k in this._lowerFields) {
      if (anotherField.getLowerFields()[k] == undefined)
        diff[k] = anotherField.getLowerFields()[k];
    }
    
    return diff;
  }

  containsChild(path) {
    if (this._children[path] != undefined) { return true; }
    else return false;
  }

  getChild(path) {
    return this._children[path];
  }

  getChildren() { return this._children; }

  getPath() { return this._path; }

  getName() { return this._name; }
  
  getParent() { return this._parent; }

  getCell() { return this._cell; }

  getColumnIndex() { return this._cell.getAddress().getColumnIndex(); }

  getLowerFields() { return this._lowerFields; }

  getHeight() { return this._height; }

  /**
   * Compare this field with another field by recursive comparing
   * those subfield names (children). Returns true, if all subfields
   * matches.
   * 
   * @param {Field} anotherField
   * @returns {Boolean}
   */
  // Note: expiremental feature
  compare(anotherField) {
    if (this._name == anotherField.getName()) {
      return deepEqual(this, anotherField);
    } else {
      return false;
    }

    function deepEqual(a, b) {      
      if (Object.keys(a.getChildren()).length == 0 && 
        Object.keys(b.getChildren()).length == 0) {        
        return true;
      } else if (Object.keys(a.getChildren()) != 0 && Object.keys(b.getChildren()) != 0) {
        for (let childName in a.getChildren()) {
          const childA = a.getChild(childName);
  
          if (b.containsChild(childA.getName())) {
            return deepEqual(childA, b.getChild(childA.getName()));                          
          } else {
            return false;
          }
        }  
      } else {
        return false;
      }
    }
  }
  
  _constructWithSubfields(height) {
    if (height == 0 || height == null || height == undefined) return;

    const self = this, rightLimit = findRightLimit(this._cell);
    
    buildSubFields(self, height, self._cell.getAddress().toString(), rightLimit);

    function findRightLimit(cell) {
      function run(initial, cell) {
        while (cell.getValue() == initial) { cell.offset(0, 1); }
        cell.offset(0, (-1));      
        return cell.getAddress().toString();
      }

      const initialVal = cell.getValue(),
            initialAddr = cell.getAddress().toString(),
            initialCell = cell.getCell();

      const limit = run(initialVal, cell);
      cell.setAddress(initialAddr);

      return limit;
    }
    
    function buildSubFields(field, height, leftLimit, rightLimit) {
      if (height == 0) {
        self._lowerFields[field.getPath()] = field;
        return;
      }
      const descRange = new Range(REUtils.offsetRow(leftLimit, 1) + ':' + 
                                  REUtils.offsetRow(rightLimit, 1), self._ws);
      let prevAddr = null;

      descRange.forEach((cell, i) => {        
        const rl = findRightLimit(cell), ll = cell.getAddress().toString();        

        if (rl != prevAddr) {
          prevAddr = rl; 
          field._addChild(cell.getAddress().toString());
          const newChildPath = StringUtils.constructPath(field.getPath(), cell.getValue());
          const newChild = field.getChild(newChildPath);
          buildSubFields(newChild, height - 1, ll, rl);
        }
      });
    }
  }

  _addChild(childAddress) {
    const childField = new Field(childAddress, this, this._ws);
    this._children[childField.getPath()] = childField;
  }
}

module.exports = Field;
