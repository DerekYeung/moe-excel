'use strict';
const workbook = require('./workbook');
const WORKBOOK = Symbol('moeExcel#workbook');

class Base {

  get workbook() {
    if (this[WORKBOOK]) {
      return this[WORKBOOK];
    }
    this[WORKBOOK] = this.createWorkbook(this.config);
    return this[WORKBOOK];
  }

  createWorkbook(config = {}) {
    return new workbook(config, this);
  }

}

module.exports = Base;
