'use strict';

const defaultEngine = require('./default');
const exceljs = require('exceljs');
const util = require('../util');
const fs = util.fs;

class Engine extends defaultEngine {

  createWorkbook(config = {}) {
    let workbook;
    if (this.work.config.stream) {
      const target = this.work.config.streamTarget || 'test';
      fs.ensureDirSync('./cache/' + target);
      const options = {
        filename: './cache/' + target + '/' + config.filename + '.xlsx',
        useStyles: true,
        useSharedStrings: true,
      };
      workbook = new exceljs.stream.xlsx.WorkbookWriter(options);
    } else {
      workbook = new exceljs.Workbook();
    }

    workbook.creator = config.creator || '';
    workbook.lastModifiedBy = config.lastModifiedBy;
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.worksheetData = {};
    workbook.isStream = !!this.work.config.stream;
    return workbook;
  }

  write(config = {}, worksheet = null, workbook = null) {
    workbook = workbook || this.work.workbook;
    if (!worksheet) {
      worksheet = this.createWorksheet(config, workbook);
    }
    workbook.loader.write(config, worksheet);
  }

  export() {
    const loader = this.work.workbook.loader;
    return loader.load().then(() => {
      const workbook = this.work.workbook.instance;
      if (workbook.isStream) {
        return workbook.commit();
      }
      return workbook.xlsx.writeFile('./write.xlsx');
    });
  }

  readFile(filepath = '') {
    return new Promise(resolve => {
      const stream = fs.createReadStream(filepath);
      const workbook = new exceljs.Workbook();
      const input = workbook.xlsx.createInputStream();
      stream.pipe(input);
      input.on('done', () => {
        resolve(workbook);
      }).on('error', () => {
        resolve(workbook);
      });
    });
  }

  createBasicSheet(config, workbook, views, real) {
    views = views || [];
    const worksheets = workbook.worksheets || [];
    const sheet_name = config.sheet_name || `Sheet${worksheets.length + 1}`;
    const headers = config.headers || [];
    const groups = Object.keys(config.groups) == 0 ? null : config.groups;
    const frozen = config.frozen || {};
    const mainTime = config.time || {
      show: false,
    };
    mainTime.time = mainTime.time || new Date().toLocaleString();
    let line = 0;
    const worksheet = workbook.addWorksheet(sheet_name, {
      views,
    });

    if (config.show.title) {
      line++;
      const titleRow = worksheet.addRow([ config.title ]);
      titleRow.height = 30;
      titleRow.alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      titleRow.font = {
        name: 'Microsoft Yahei',
        family: 4,
        size: 16,
        bold: false,
      };
    }
    if (mainTime.show) {
      line++;
      worksheet.addRow([ '数据统计时间：' + mainTime.time ]);
    }
    const tops = groups;
    if (tops) {
      line++;
      const topRow = [];
      for (const k in tops) {
        topRow.push(tops[k]);
      }
      worksheet.addRow(topRow);
    }
    const row = [];
    let columnTarget = 0;
    headers.forEach(node => {
      ++columnTarget;
      node.columnIndex = columnTarget;
      row.push(node.name);
    });
    const unNumber = /\d+/g;
    const columnRow = worksheet.addRow(row);
    columnRow.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    let xSplit = 0;
    line++;

    config.columnIndex = line;

    columnRow.eachCell(function(cell, colNumber) {
      const head = headers.find(function(data) {
        return data.columnIndex == colNumber;
      });
      const column = cell.address.replace(unNumber, '');
      if (head) {
        head.column = column;
        head.address = cell.address;
        const length = util.strlen(head.name);
        head.maxLength = length;
        worksheet.getColumn(head.columnIndex).width = length + 2.62;
        if (head.frozen) {
          xSplit = colNumber;
        }
      }
    });

    xSplit = frozen.x ? xSplit : 0;
    const ySplit = frozen.y ? line : 0;
    const finalViews = [{
      state: 'frozen',
      xSplit,
      ySplit,
    }];
    const dataTop = line - 1;
    for (const k in tops) {
      const groups = headers.filter(data => {
        return data.group == k;
      });
      if (groups && groups.length > 0) {
        const begin = groups[0].address.replace(unNumber, '');
        const end = groups[groups.length - 1].address.replace(unNumber, '');
        worksheet.mergeCells(begin + dataTop + ':' + end + dataTop);
        const cell = worksheet.getCell(begin + dataTop);
        cell.value = tops[k];
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
        };
      }
    }
    const total_begin = headers[0].address.replace(unNumber, '');
    const total_end = headers[headers.length - 1].address.replace(unNumber, '');
    let titleTop = 0;
    if (config.show.title) {
      titleTop = 1;
      const title_merge = total_begin + titleTop + ':' + total_end + titleTop;
      worksheet.mergeCells(title_merge);
    }
    if (mainTime.show) {
      titleTop++;
      const date_merge = total_begin + titleTop + ':' + total_end + titleTop;
      worksheet.mergeCells(date_merge);
    }
    // worksheet.getCell('A3').value = {
    //     formula: 'HYPERLINK("#1.1!c1","1")',
    //     result: '1.1'
    // };
    return real ? worksheet : finalViews;
  }

  createWorksheet(config = {}, workbook = null) {
    workbook = workbook || this.createWorkbook(config);
    config = config || {};

    const mainTime = config.time || {
      show: false,
    };
    mainTime.time = mainTime.time || new Date().toLocaleString();

    const headbook = new exceljs.Workbook();
    const views = this.createBasicSheet(config, headbook, null, false);
    const worksheet = this.createBasicSheet(config, workbook.instance, views, true);
    return worksheet;
  }

  read() {
    return this.readFile(this.work.filepath).then(workbook => {
      workbook.eachSheet(worksheet => {
        const maps = [];
        const data = [];
        const row = worksheet.getRow(1);
        row.values.forEach(value => {
          maps.push(value);
        });
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1) {
            const result = {};
            const values = row.values || [];
            maps.forEach((k, i) => {
              let value = values[i + 1] || '';
              let text = '';
              if (value instanceof Object) {
                if (value.text) {
                  value = value.text;
                }
                if (value.richText instanceof Array) {
                  value.richText.forEach(rich => {
                    text += rich.text;
                  });
                  value = text;
                }
              }
              value = value || '';
              try {
                value = value.toString();
                value = value.replace(/^\s+|\s+$/gm, '');
              } catch (e) {
                value = (value || '');
              }
              if (k) {
                result[k] = value;
              }
            });
            data.push(result);
          }
        });
        return this.work.workbook.createWorksheet({
          maps,
          data,
        });
      });
      return this.work.workbook;
    }).catch(e => {
      return e;
    });
  }

  overwrite(path = '') {
    return new Promise((resolve, reject) => {
      this.work.workbook.instance.xlsx.writeBuffer().then(buffer => {
        const writeStream = fs.createWriteStream(path);
        writeStream.on('finish', () => {
          resolve();
        });
        writeStream.on('error', e => {
          reject(e);
        });
        writeStream.write(buffer);
        writeStream.end();
      });
    });
  }

}

module.exports = Engine;
