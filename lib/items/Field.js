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

    if (this._cell.getValue() == null) {
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
    }
  }

  equals(anotherField) {
    for (let k in this._lowerFields) {
      if (anotherField.getLowerFields()[k] == undefined)
        return false;
    }
    return true;
  }

  containsChild(name) {
    if (this._children[name] != undefined) { return true; }
    else return false;
  }

  getChild(name) {
    return this._children[name];
  }

  getChildren() { return this._children; }

  getPath() { return this._path; }

  getName() { return this._name; }
  
  getParent() { return this._parent; }

  getCell() { return this._cell; }

  getLowerFields() { return this._lowerFields; }

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
          const newChild = field.getChild(cell.getValue());
          buildSubFields(newChild, height - 1, ll, rl);
        }
      });
    }
  }

  _addChild(childAddress) {
    const childField = new Field(childAddress, this, this._ws);
    this._children[childField.getName()] = childField;
    //console.log('added child: ' + childAddress + ' to ' + this._cell.getAddress());
  }
}

module.exports = Field;
