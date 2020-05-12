'use strict';

const exceljsEngine = require('./engines/exceljs');

const Engines = {};

class Master {
  constructor(engine = 'exceljs', order = {}) {
    this.engine = engine;
    this.work = order;
  }

  getWorker(work = {}) {
    const instance = new Master(this.engine, work);
    return instance;
  }

  registerEngine(name = '', engine = {}) {
    Engines[name] = engine;
  }

  getEngine(name = '') {
    name = name || this.engine;
    return new Engines[name](this.work);
  }

}

const master = new Master();

master.registerEngine('exceljs', exceljsEngine);

module.exports = master;
