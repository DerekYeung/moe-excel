'use strict';
const reader = require('./libs/reader');
const writer = require('./libs/writer');
const master = require('./libs/master');

const Excel = {
  reader,
  writer,
  master,
};

module.exports = Excel;
