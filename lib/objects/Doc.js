'use strict';

const DocObject = require('./DocObject'),
      Range = require('../cells/Range'),
      REUtils = require('../utils/REUtils'),
      xl = require('../utils/xl');

/**
  * 
  * @param {*} path 
  * @param {*} sheet 
  * @param {*} objectsRange rows diapasone, where column contains
  * name of a row (`DocObject`)
  */
class Doc {
  constructor(path, sheet, objectsRange) {
    this._path = path;
    this._sheet = sheet;
    this._objectsRange = objectsRange;
    this._isPlainDoc = false;
    this._containsMidRows = false;
    this._isTopDoc = false;
  }

  /**
   * 
   * @param {*} valuesRange is a columns diapasone (as `A:Z`), which contains
   * all values of current `DocObject` (row)
   * @param {*} topSign 
   * @param {*} midSign 
   * @param {*} lowSign 
   */
  constructObjects(valuesRange, topSign, midSign, lowSign) {
    this._valuesRange = valuesRange;

    // Construct document with plain rows
    if ((topSign == null || topSign == undefined) && 
        (midSign == null || midSign == undefined)) 
    {
      this._isPlainDoc = true;

      return xl.readFile(this._path, this._sheet).then((ws) => {
        this._ws = ws;
        this._plainRows = {};
          
        const objsRange = new Range(this._objectsRange, ws, true);

        objsRange.forEach((cell) => {          
          const obj = new DocObject(cell.getAddress().toString(),
                                    REUtils.fillColRangeWithRows(valuesRange, cell.getRow()),
                                    ws, this._objectsRange);
          this._plainRows[obj.row.getName()] = obj;
        });
      });
    }

    // Construct document with top objects and with children
    if ((topSign != null && topSign != undefined) &&
        (midSign != null && midSign != undefined))
    {
      this._isTopDoc = true;

      return xl.readFile(this._path, this._sheet).then((ws) => {
        this._ws = ws;
        this._topRows = {};
        const topFG = ws.getCell(topSign).fill.fgColor,
              midFG = ws.getCell(midSign).fill.fgColor;

        const objsRange = new Range(this._objectsRange, ws, true);
        
        objsRange.forEach((cell) => {
          if (cell.getCell().fill != undefined) {
            if (cell.getCell().fill.fgColor.tint == topFG.tint || 
                cell.getCell().fill.fgColor.argb == topFG.argb)
            {
              const obj = 
                new DocObject(cell.getAddress().toString(),
                              REUtils.fillColRangeWithRows(valuesRange, cell.getRow()),
                              ws, this._objectsRange, null, midFG, topFG);

              this._topRows[obj.row.getName()] = obj;
            }
          }
        });
      });
    }
    
    // Construct document with mid objects and with children
    if ((lowSign != null || lowSign != undefined) && 
        (midSign != null || midSign != undefined))
    {
      this._containsMidRows = true;
      
      return xl.readFile(this._path, this._sheet).then((ws) => {
        this._ws = ws;
        this._midRows = {};
        const lowFG = 'empty'/*ws.getCell(lowSign).fill.fgColor*/,
              midFG = ws.getCell(midSign).fill.fgColor;

        const objsRange = new Range(this._objectsRange, ws, true);

        objsRange.forEach((cell) => {
          if (cell.getCell().fill != undefined) {
            if (cell.getCell().fill.fgColor.tint == midFG.tint || 
                cell.getCell().fill.fgColor.argb == midFG.argb)
            {
              const obj = 
                new DocObject(cell.getAddress().toString(),
                              REUtils.fillColRangeWithRows(valuesRange, cell.getRow()),
                              ws, this._objectsRange, null, lowFG, midFG);
      
              this._midRows[obj.row.getName()] = obj;
            }
          }
        });
      });
    }
  }

  diff(anotherDoc) {
    if (this._isTopDoc && anotherDoc.isTopDoc()) {
      return this._diffTop(anotherDoc);
    } else if (this._isTopDoc && !anotherDoc.isTopDoc() && this._containsMidRows && anotherDoc.containsMidRows()) {
      return this._diffTopWithMid(anotherDoc);
    } else if (this._containsMidRows && anotherDoc.containsMidRows()) {
      return this._diffMid(anotherDoc);
    } 
  }

  _diffTop(anotherDoc) {
    const diff = {};
    
    const thisTR = this._topRows,
          thisMR = this._midRows,
          anotherTR = anotherDoc.getTopRows(),
          anotherMR = anotherDoc.getMidRows();

    // for each top row
    for (let thisTRitemName in thisTR) {
      const thisTRitemChildren = thisTR[thisTRitemName].getChildren(),
            anotherTRitem = anotherTR[thisTRitemName],
            anotherTRitemChildren = anotherTRitem.getChildren();
              
      // if such a top row not exists - copy it and it childrent
      if (anotherTRitem == undefined) {
        diff[thisTRitemName] = {};
        Object.keys(thisTRitemChildren).forEach((midName) => {
          diff[thisTRitemName]= thisTRitemChildren[midName];

          // and subsequently copy all children of each child
          if (this._containsMidRows) {
            diff[thisTRitemName][midName] = thisMR[midName].getChildren();
          }
        });
      // else if it exists
      } else {
        // for each child row of top row
        for (let thisTRchildName in thisTRitemChildren) {
          // if such a child not exists
          if (anotherTRitemChildren[thisTRchildName] == undefined) {
            diff[thisTRitemName][thisTRchildName] = {};
              
            if (this._containsMidRows) {
              if (diff[thisTRitemName] == undefined) { diff[thisTRitemName] = {}; }
              diff[thisTRitemName][thisTRchildName] = thisMR[thisTRchildName].getChildren();
            }
          // else if it exists
          } else {
            const thisLR = thisMR[thisTRchildName].getChildren(),
                  anotherLR = anotherMR[thisTRchildName].getChildren();
              
            if (diff[thisTRitemName] == undefined) { diff[thisTRitemName] = {}; }

            for (let thisLRitem in thisLR) {
              if (anotherLR[thisLRitem] == undefined) {
                if (diff[thisTRitemName][thisTRchildName] == undefined) diff[thisTRitemName][thisTRchildName] = {};
                diff[thisTRitemName][thisTRchildName][thisLRitem] = thisLR[thisLRitem];
              }
            }
          }
        }
      }
    }    

    return diff;
  }

  _diffMid(anotherDoc) {
    const diff = {};
    
    const thisMR = this._midRows,
          anotherMR = anotherDoc.getMidRows();

    for (let thisMRname in thisMR) {
      // if such a child not exists
      if (anotherMR[thisMRname] == undefined) {
        diff[thisMRname] = thisMR[thisMRname].getChildren();        
      // else if it exists
      } else {
        const thisLR = thisMR[thisMRname].getChildren(),
              anotherLR = anotherMR[thisMRname].getChildren();
        
        for (let thisLRitem in thisLR) {
          if (anotherLR[thisLRitem] == undefined) {
            if (diff[thisMRname] == undefined) { diff[thisMRname] = {}; }
            // if (diff[thisMRname] == undefined) { diff[thisMRname] = {}; }
            diff[thisMRname][thisLRitem] = thisLR[thisLRitem];
          }
        }
      }
    }          

    return diff;
  }

  // This document is top and another document is mid
  _diffTopWithMid(anotherMidDoc) {
    const diff = {};

    const thisTR = this._topRows,
          thisMR = this._midRows,
          anotherMR = anotherMidDoc.getMidRows();

    for (let thisTRName in thisTR) {
      diff[thisTRName] = {};
      const thisTRitemChildren = thisTR[thisTRName].getChildren();
      for (let thisTRitemMidName in thisTRitemChildren) {
        if (anotherMR[thisTRitemMidName] == undefined) {
          diff[thisTRName][thisTRitemMidName] = thisMR[thisTRitemMidName].getChildren();
        } else {
          const thisLR = thisMR[thisTRitemMidName].getChildren(),
                anotherLR = anotherMR[thisTRitemMidName].getChildren();

          for (let thisLRname in thisLR) {
            if (anotherLR[thisLRname] == undefined) {
              if (diff[thisTRName][thisTRitemMidName] == undefined) { diff[thisTRName][thisTRitemMidName] = {}; }
              diff[thisTRName][thisTRitemMidName][thisLRname] = thisLR[thisLRname];
            }
          }
        }
      }
    }

    return diff;
  }
  
  getTopRows() { return this._topRows; }

  getMidRows() { return this._midRows; }

  getPlainRows() { return this._plainRows; }

  isPlainDoc() { return this._isPlainDoc; }
  
  isTopDoc() { return this._isTopDoc; }

  containsMidRows() { return this._containsMidRows; }
}

module.exports = Doc;