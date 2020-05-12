'use strict';

class Worksheet {
  constructor(config = {}, workbook = {}) {
    const defaultConfig = {
      id: '',
      sheet_name: '',
      title: '',
      groups: {},
      headers: [],
      show: {
        title: true,
        time: true,
      },
      frozen: {
        x: true,
        y: true,
      },
      datas: [],
    };
    this.maps = config.maps || [];
    this.data = config.data || [];
    delete config.maps;
    delete config.data;
    this.config = Object.assign({}, defaultConfig, (config || {}));
    this.workbook = workbook;
    this.instance = null;
  }

  set(key = '', value = '') {
    if (key instanceof Object) {
      this.config = Object.assign({}, this.config, key);
      return this;
    }
    this.config[key] = value;
    return this;
  }

  name(name) {
    return this.set('sheet_name', name);
  }

  title(title = '') {
    return this.set('title', title);
  }

  headers(headers = []) {
    return this.set('headers', headers);
  }

  groups(groups = []) {
    return this.set('groups', groups);
  }

  datas(datas) {
    return this.set('datas', datas);
  }

  show(key) {
    this.config.show[key] = true;
    return this;
  }

  hide(key) {
    this.config.show[key] = false;
    return this;
  }

  frozen(key, frozen) {
    this.config.show[key] = (frozen || frozen == undefined);
    return this;
  }

  write(config = {}) {
    config = Object.assign({}, this.config, config);
    const workbook = this.workbook;
    if (!this.instance) {
      this.instance = this.workbook.master.engine.createWorksheet(config, workbook);
    }
    workbook.loader.write(config, this.instance);
  }

}

module.exports = Worksheet;
