'use strict';
const util = require('./util');
class Loader {
  constructor(workbook = {}) {
    this.workbook = workbook;
    this.jobs = [];
  }

  write(config, worksheet = null) {
    this.jobs.push({
      config,
      worksheet,
    });
  }

  load() {
    return Promise.all(this.jobs.map(node => {
      return this.loadToSheet(node.config, node.worksheet);
    }));
  }

  loadToSheet(config = {}, worksheet = null) {
    return new Promise(resolve => {
      if (config.remote && config.remote.url) {
        const limit = config.length || 20000;
        return resolve(this.request(config, worksheet, limit, 0));
      }
      if (config.datas) {
        return resolve(this.push(config, worksheet, config.datas));
      }
    });
  }

  request(config, worksheet, limit, offset) {
    const remote = config.remote || {};
    const url = remote.url || '';
    const method = (remote.method || 'GET');
    const query = remote.query || {};
    query.offset = offset;
    query.limit = limit;
    return util.fly.request(url, query, {
      method,
    }).then(response => {
      const json = response.data || {};
      return remote.onLoad(json, response);
    }).then(result => {
      const datas = result.datas || [];
      const total = result.total || 0;
      const handle = (offset + limit);
      const end = (handle >= total || total < limit);
      this.push(config, worksheet, datas);
      if (!end) {
        offset += limit;
        return this.request(config, worksheet, limit, offset);
      }
      return null;
    });
  }

  push(config, worksheet, datas) {
    const workbook = this.workbook.instance;
    const headers = config.headers || [];
    datas.forEach(node => {
      const data = [];
      const emptys = [];
      const names = [];
      headers.forEach((target, index) => {
        const id = target.key;
        names.push(id);
        let value = node[id];
        if (target.formatter) {
          if (typeof (target.formatter) === 'function') {
            value = target.formatter(value, node);
          }
        }
        const isEmpty = (!value || value == '' || value == null || value.length <= 0);
        if (isEmpty && target.empty) {
          value = target.empty;
          emptys.push(index);
        }
        const length = util.strlen(value);
        data[index] = value == null ? '' : value;
        const haeder_max = target.maxLength || 0;
        if (length > haeder_max) {
          target.maxLength = length;
          worksheet.getColumn(target.columnIndex).width = length + 2.62;
        }
      });
      const row = worksheet.addRow(data);
      row.eachCell((cell, colNumber) => {
        const name = names[colNumber - 1];
        const address = cell.address;
        if (!node._map) {
          node._map = {};
        }
        node._map[name] = address;
      });
      if (emptys.length > 0) {
        emptys.forEach(node => {
          row.getCell(node + 1).style.font = {
            color: {
              argb: 'FFE4E2E1',
            },
          };
        });
      }
      row.alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      if (workbook.isStream) {
        row.commit();
      }
    });
    datas = null;
    return true;

  }

}

module.exports = Loader;
