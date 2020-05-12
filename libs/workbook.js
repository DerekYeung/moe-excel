'use strict';
const worksheet = require('./worksheet');
const loader = require('./loader');
const INSTANCE = Symbol('moeExcel#workbookInstance');

class Workbook {
  constructor(config = {}, master = {}) {
    this.config = config;
    this.master = master;
    this.worksheets = [];
    this.loader = new loader(this);
  }

  get instance() {
    if (this[INSTANCE]) {
      return this[INSTANCE];
    }
    this[INSTANCE] = this.master.engine.createWorkbook(this.config);
    return this[INSTANCE];
  }

  get sheet() {
    return this.first;
  }

  get first() {
    return this.getFirstWorksheet();
  }

  getFirstWorksheet() {
    if (this.worksheets[0]) {
      return this.worksheets[0];
    }
    return this.createWorksheet();
  }

  getWorksheet(id = '') {
    if (!isNaN(id)) {
      return this.worksheets[id - 1];
    }
    return null;
  }

  addWorksheet(worksheet = null) {
    this.worksheets.push(worksheet);
    return this;
  }

  createWorksheet(config = {}) {
    const sheet = new worksheet(config, this);
    this.addWorksheet(sheet);
    return sheet;
  }

  overwrite(path = '') {
    return this.master.engine.overwrite(path);
  }

  export() {
    return this.master.engine.export();
  }

}

module.exports = Workbook;
