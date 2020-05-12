'use strict';
const Base = require('./base');
const master = require('./master');

class Reader extends Base {

  constructor(filepath = '') {
    super(filepath);
    this.filepath = filepath;
    this.worker = master.getWorker(this);
    this.engine = this.worker.getEngine();
  }

  readFile(filepath = '') {
    return this.engine.readFile(filepath);
  }

  overwrite(path = this.filepath) {
    return this.workbook.overwrite(path);
  }

  read() {
    return this.engine.read();
  }

}

module.exports = Reader;
