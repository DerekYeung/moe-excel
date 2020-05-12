'use strict';

const clone = require('clone');
const fs = require('fs-extra');
const qs = require('query-string');
const fly = require('flyio');

function strlen(str) {
  if (str == null) return 0;
  if (typeof str !== 'string') {
    str += '';
  }
  // eslint-disable-next-line no-control-regex
  return str.replace(/[^\x00-\xff]/g, '01').length;
}

module.exports = {
  strlen,
  clone,
  fs,
  qs,
  fly,
};
