'use strict';

const DocObject = require('./DocObject'),
      Range = require('../cells/Range'),
      Field = require('../items/Field'),
      REUtils = require('../utils/REUtils'),
      xl = require('../utils/xl'),
      spawnSync = require('child_process').spawnSync;

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
    this._fields = {};
    this._cells = {};
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
      
      return xl.readFile(this._path, this._sheet).then((res) => {
        this._ws = res.ws;
        this._wb = res.wb;
        this._plainRows = {};
        
        this._accumulateCells(this._ws, valuesRange);
        
        const objsRange = new Range(this._objectsRange, this._ws, true);

        objsRange.forEach((cell) => { 
          if (cell.getCell().fill == undefined) {
            const obj = new DocObject(cell.getAddress().toString(),
                                      valuesRange,
                                      this._ws, this._objectsRange);
            
            this._plainRows[obj.row.getName()] = obj;
          }
        });
      });
    }

    // Construct document with top objects and with children
    if ((topSign != null && topSign != undefined) &&
        (midSign != null && midSign != undefined))
    {
      console.log('top rows creation');
      this._isTopDoc = true;

      return xl.readFile(this._path, this._sheet).then((res) => {
        this._ws = res.ws;
        this._wb = res.wb;
        
        this._topRows = {};
        const topFG = this._ws.getCell(topSign).fill.fgColor,
              midFG = this._ws.getCell(midSign).fill.fgColor;
              
        this._accumulateCells(this._ws, valuesRange);

        const objsRange = new Range(this._objectsRange, this._ws, true);

        objsRange.forEach((cell) => {
          if (cell.getCell().fill != undefined) {
            if (cell.getCell().fill.fgColor.tint == topFG.tint || 
                cell.getCell().fill.fgColor.argb == topFG.argb)
            {
              if (cell.getCell().fill.fgColor.indexed == 9 || 
                (cell.getCell().fill.fgColor.theme == 0 && cell.getCell().fill.fgColor.tint == undefined)) return;

                const obj = 
                  new DocObject(cell.getAddress().toString(),
                              valuesRange,
                              this._ws, this._objectsRange, null, midFG, topFG);

                this._topRows[obj.row.getName()] = obj;
                // this._topRows[obj.getPath()] = obj;
            }
          }
        });
      });
    }
    
    // Construct document with mid objects and with children
    if ((lowSign != null || lowSign != undefined) && 
        (midSign != null || midSign != undefined))
    {
      console.log('mid objects creation');

      this._containsMidRows = true;
      
      return xl.readFile(this._path, this._sheet).then((res) => {
        this._ws = res.ws;
        this._wb = res.wb;
        this._midRows = {};
        const lowFG = 'empty',
              midFG = this._ws.getCell(midSign).fill.fgColor;

        this._accumulateCells(this._ws, valuesRange);

        const objsRange = new Range(this._objectsRange, this._ws, true);

        objsRange.forEach((cell) => {
          if (cell.getCell().fill != undefined) {
            if (cell.getCell().fill.fgColor.tint == midFG.tint || 
                cell.getCell().fill.fgColor.argb == midFG.argb)
            {
              if (cell.getCell().fill.fgColor.indexed == 9 || 
                 (cell.getCell().fill.fgColor.theme == 0 && cell.getCell().fill.fgColor.tint == undefined)) return;

              // const parentName = this._getParentByChildrenRange(cell.getAddress().toString());
              const obj = 
                new DocObject(cell.getAddress().toString(),
                              valuesRange,
                              this._ws, this._objectsRange, null, lowFG, midFG);
              
              let midRowName = obj.row.getName();
              // obj.setPath(parentName, midRowName);

              this._midRows[midRowName] = obj;
              // this._midRows[obj.getPath()] = obj;
            }
          }
        });
      });
    }
  }

  refresh() {
    this._isPlainDoc = false;
    this._isTopDoc = false;
    this._containsMidRows = false;
    this._fields = {};
    this._cells = {};
    this._ws = undefined;
    this._wb = undefined;
    this._topRows = this._midRows = {};    
  }

  _getParentByChildrenRange(childAddr) {
    for (let parentName in this._topRows) {
      const childrenRange = this._topRows[parentName].getChildrenRange();
      if (childrenRange.containsAddr(childAddr)) {
        return parentName;
      }
    }
    return null;
  }

  diff(anotherDoc) {
    if (this._isTopDoc && anotherDoc.isTopDoc()) {
      return this._diffTop(anotherDoc);
    } else if (this._isTopDoc && !anotherDoc.isTopDoc() && this._containsMidRows && anotherDoc.containsMidRows()) {
      return this._diffTopWithMid(anotherDoc);
    } else if (this._containsMidRows && anotherDoc.containsMidRows()) {
      return this._diffMid(anotherDoc);
    } else if (this._isPlainDoc && anotherDoc.isPlainDoc()) {
      return this._diffPlain(anotherDoc);
    }
  }

  _diffTop(anotherDoc) {
    console.log('diffTop');
    const diff = {};
    
    const thisTR = this._topRows,
          thisMR = this._midRows,
          anotherTR = anotherDoc.getTopRows(),
          anotherMR = anotherDoc.getMidRows();

    // for each top row
    for (let thisTRitemName in thisTR) {
      const thisTRitemChildren = thisTR[thisTRitemName].getChildren(),
            anotherTRitem = anotherTR[thisTRitemName];
            
      // const anotherTRitemChildren = anotherTRitem.getChildren();
              
      // if such a top row not exists - copy it and it's childrent
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
        if (diff[thisTRitemName] == undefined) { diff[thisTRitemName] = {}; }
        // for each child row of top row
        for (let thisTRchildName in thisTRitemChildren) {
          // if such a child not exists
          const anotherTRitemChildren = anotherTRitem.getChildren();
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
            if (diff[thisMRname] == undefined) { 
              diff[thisMRname] = {};
              diff[thisMRname]['childrenRange'] = anotherMR[thisMRname].getChildrenRange();
            }
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

  _diffPlain(anotherPlainDoc) {
    const diff = {};

    const thisPL = this._plainRows,
          anotherPL = anotherPlainDoc.getPlainRows();

    for (let rowName in thisPL) {
      if (anotherPL[rowName] == undefined) {
        diff[rowName] = thisPL[rowName];
      }
    }

    return diff;
  }

  merge(anotherDoc) {
    // both contains top rows, not implemented yet
    if (this._isTopDoc && anotherDoc.isTopDoc()) {
      this._mergeTopWithTop(anotherDoc);
    // this contains top and mid, and another contains mid ondly
    } else if (this._isTopDoc && !anotherDoc.isTopDoc() && 
               this._containsMidRows && anotherDoc.containsMidRows()) 
    { 
      this._mergeTopWithMid(anotherDoc);
    // this contains mid rows, and another contains top and mid
    } else if (!this._isTopDoc && anotherDoc.isTopDoc() && 
              this._containsMidRows && anotherDoc.containsMidRows()) 
    {
      this._mergeMidWithTop(anotherDoc);
    // both contains mid rows
    } else if (this._containsMidRows && anotherDoc.containsMidRows()) {      
      this._mergeTopWithMid(anotherDoc);
    // both contains plain rows only
    } else if (this._isPlainDoc && anotherDoc.isPlainDoc()) {
      this._mergePlain(anotherDoc);
    }
  }
  
  _mergeTopWithTop(anotherTopDoc) {
    console.log('mergeTopWithTop');
    const thisTRs = this._topRows,
          anotherTRs = anotherTopDoc.getTopRows();
  
    for (let thisTopRowName in thisTRs) {
      if (anotherTRs[thisTopRowName] == undefined) continue;

      const thisMRNames = thisTRs[thisTopRowName].getChildren(),
            thisMRs = this._midRows,
            anotherMRs = anotherTopDoc.getMidRows();
      
      for (let thisMidRowName in thisMRNames) {
        if (anotherMRs[thisMidRowName] == undefined) continue;
                
        const thisLowRows = thisMRs[thisMidRowName].getChildren(),
              anotherLowRows = anotherMRs[thisMidRowName].getChildren();

        for (let thisLowRowName in thisLowRows) {
          const thisLowRow = thisLowRows[thisLowRowName],
                anotherLowRow = anotherLowRows[thisLowRowName];
  
          if (anotherLowRow == undefined) continue;

          this._mergeRow(thisLowRow, this._fields, anotherLowRow, anotherTopDoc.getFields());
        }
      }
    }
  }

  _mergeTopWithMid(anotherMidDoc) {
    console.log('merge top with mid');
    const thisMRs = this._midRows,
          anotherMRs = anotherMidDoc.getMidRows();

    for (let thisMidRowName in thisMRs) {
      if (anotherMRs[thisMidRowName] == undefined) continue;

      const thisLowRows = thisMRs[thisMidRowName].getChildren(),
            anotherLowRows = anotherMRs[thisMidRowName].getChildren();
      
      for (let thisLowRowName in thisLowRows) {
        const thisLowRow = thisLowRows[thisLowRowName],
              anotherLowRow = anotherLowRows[thisLowRowName];
        
        if (anotherLowRow == undefined) continue;

        this._mergeRow(thisLowRow, this._fields, anotherLowRow, anotherMidDoc.getFields());        
      }
    }    
  }

  _mergeMidWithTop(anotherDoc) {
    console.log('merge mid with top');
    const anotherTRs = anotherDoc.getTopRows(),
          anotherMRs = anotherDoc.getMidRows(),
          thisMRs = this._midRows;
    
    for (let trName in anotherTRs) {
      const anotherMRNames = anotherTRs[trName].getChildren();

      for (let anotherMRName in anotherMRNames) {
        if (thisMRs[anotherMRName] == undefined) continue; 

        const thisLowRows = thisMRs[anotherMRName].getChildren(),
              anotherLowRows = anotherMRs[anotherMRName].getChildren();
        
        for (let thisLowRowName in thisLowRows) {
          const thisLowRow = thisLowRows[thisLowRowName],
                anotherLowRow = anotherLowRows[thisLowRowName];

          if (anotherLowRow == undefined) continue; 
          this._mergeRow(thisLowRow, this._fields, anotherLowRow, anotherDoc.getFields()); 
        }
      }
    }
  }

  _mergePlain(anotherDoc) {
    console.log('merging plain');
    for (let anotherName in anotherDoc.getPlainRows()) {
    }
    for (let thisRowName in this._plainRows) {
      const thisRow = this._plainRows[thisRowName].row,
            anotherRow = anotherDoc.getPlainRows()[thisRowName];

      if (anotherRow == undefined) continue;
      this._mergeRow(thisRow, this._fields, anotherRow.row, anotherDoc.getFields());
    }
  }

  _mergeRow(fromRow, fromFields, toRow, toFields) {
    for (let topFromFieldName in fromFields) {
      const fromLowFields = fromFields[topFromFieldName].getLowerFields(),
            toLowFields = toFields[topFromFieldName].getLowerFields();

      for (let lowFromFieldName in fromLowFields) {
        const fromColIdx = fromLowFields[lowFromFieldName].getColumnIndex(),
              toColIdx = toLowFields[lowFromFieldName].getColumnIndex(), 
              cellAddr = REUtils.indexToCol(fromColIdx) + fromRow.getRow(),       
              cell = this.getCell(cellAddr);
        
        let value;

        // Заменить значение в рещультирующей ячейке на пустое, даже если в
        // исходной значение отсутсвует
        if (cell == undefined || cell == null) {
          value = '';
        } else {
          value = cell.getValue();
          if (value.sharedFormula != undefined) {
            value = value.result;
          }
        }
        // console.log(value);
        toRow.setValueByColIndex(toColIdx, value);
      }
    }
  }

  _accumulateCells(ws, valuesRange) {
    const [left, right] = valuesRange.split(':'),
          [topCell, lowCell] = this._objectsRange.split(':'),
          range = left + REUtils.getRow(topCell) + ':' + right + REUtils.getRow(lowCell);
      
    const r = new Range(range, ws);
    r.forEach((cell) => {
      this._cells[cell.getAddress().toString()] = cell;
    });
  }

  equalsFields(another) {
    for (let thisFieldName in this._fields) {
      const thisField = this._fields[thisFieldName],
            anotherField = another.getFields()[thisFieldName];

      if (anotherField == undefined || !thisField.equals(anotherField)) {
        return false;
      }
    }

    for (let anotherFieldName in another.getFields()) {
      const thisField = this._fields[anotherFieldName],
            anotherField = another.getFields()[anotherFieldName];

      if (thisField == undefined || !anotherField.equals(thisField)) {
        return false;
      }
    }

    return true;
  }

  buildFieldSet(fieldRange, height) {
    console.log('building field set');
    return xl.readFile(this._path, this._sheet).then((res) => {
      const topFields = [];
      let prevVal;

      new Range(fieldRange, res.ws).forEach((cell) => {
        if (cell.getValue() != prevVal) {
          prevVal = cell.getValue();
          topFields.push(cell.getAddress().toString());
        }
      });

      topFields.forEach((addr) => {
        const f = new Field(addr, null, res.ws,height);
        this._fields[f.getName()] = f;
      });
    });
  }

  // Unstable
  appendRow(rowNum, col, valuesRange, strToAppend) {
    const cp = spawnSync('cscript', 
              ['C:/Users/ShaytanovAI/Documents/work/xlutils/lib/vbs/append_row.vbs', this._path,
              this._sheet, rowNum, col, valuesRange,
              strToAppend]);
    
    console.log('err');
    console.log(cp.stderr.toString('utf8'));
    console.log('out');
    console.log(cp.stdout.toString());
  }

  save(path) {
    return xl.save(path, this._wb);
  }

  getCell(addr) { return this._cells[addr]; }

  getTopRows() { return this._topRows; }

  getMidRows() { return this._midRows; }

  getPlainRows() { return this._plainRows; }

  getFields() { return this._fields; }

  isPlainDoc() { return this._isPlainDoc; }
  
  isTopDoc() { return this._isTopDoc; }

  containsMidRows() { return this._containsMidRows; }
}

module.exports = Doc;