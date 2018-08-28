'use strict';

class StringUtils {
	static eraseSPs(str) {
		return str.replace(/\s+/g, ' ');
  }
  
	static eraseEOLs(str) {
		return str.replace(/\r?\n|\r/g, '');
  }
  
	static eraseTrailingPeriods(str) {            
    let acc = '';
    let strlen = str.length;

    for (let i = 0; ; i ++) {      
      if (str[i] != '.') { acc += str.slice(i, strlen); break; }
      else continue;
    }
    strlen = acc.length;

    for (let i = 0; ; i ++) {      
      if (acc[strlen - i - 1] != '.') {         
        acc = acc.slice(0, strlen - i);
        break; 
      }
      else continue;
    }

    return acc;
  }
  
	static eraseAll(str) {
		return this.eraseEOLs(this.eraseSPs(this.eraseTrailingPeriods(str)));
  }
  
  static constructPath(parentPath, childName) {
    if (parentPath == undefined || parentPath == null) {
      return childName + '/';
    }
    return parentPath + childName + '/';
  }

  static containsArrItem(str, arr) {
    let contains = false;
    for (let item of arr) {

      if (str.indexOf(item) != -1) {
        contains = true;
        break;
      }

      else continue;
    }

    return contains;
  }
}

module.exports = StringUtils;