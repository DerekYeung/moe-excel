'use strict';

const excel = require('../index');

const writer = excel.writer;
const xlsx = new writer();
const sheet = xlsx.sheet;
sheet.name('导出数据').title('演示数据').headers([
  {
    group: 'school',
    key: 'school_name',
    name: '学校',
    link: {
      sheet: '测试',
    },
    frozen: true,
    formatter(value) {
      return `[${value}]`;
    },
  },
  {
    group: 'school',
    key: 'name',
    name: '姓名',
  },
  {
    group: 'classes',
    key: 'classes',
    name: '班级',
  },
])
  .groups({
    school: '学校信息',
    classes: '班级信息',
  })
  .write({
    datas: [{
      school_name: '测试',
    }],
    // remote: {
    //   url: 'http://127.0.0.1:7001/api/test/xlsx',
    //   method: 'GET',
    //   query: {},
    //   length: 200,
    //   onLoad(json) {
    //     return {
    //       datas: json.data.students,
    //       total: json.data.total,
    //     };
    //   },
    // },
  });

xlsx.export();
