'use strict';
const Base = require('./base');
const master = require('./master');
const workbook = require('./workbook');
const WORKBOOK = Symbol('moeExcel#workbook');

class Writer extends Base {
  constructor(config = {}) {
    super(config);
    this.config = Object.assign({}, config);
    this.worker = master.getWorker(this);
    this.engine = this.worker.getEngine();
  }

  get workbook() {
    if (this[WORKBOOK]) {
      return this[WORKBOOK];
    }
    this[WORKBOOK] = this.createWorkbook(this.config);
    return this[WORKBOOK];
  }

  get sheet() {
    return this.createWorksheet(this.config);
  }

  get stream() {
    return this.createStreamWriter();
  }

  createStreamWriter() {
    this.config.stream = true;
    return this;
  }

  createWorkbook(config = {}) {
    return new workbook(config, this);
  }

  createWorksheet(config = {}) {
    return this.workbook.createWorksheet(config);
  }

  write(config = {}, worksheet = null, workbook = null) {
    return this.workbook.write(config, worksheet, workbook);
  }

  export() {
    return this.workbook.export();
  }
}

module.exports = Writer;
