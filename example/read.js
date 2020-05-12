'use strict';

const excel = require('../index');

const reader = excel.reader;

const read = new reader('./read.example.xlsx');
read.read().then(workbook => {
  const sheet = workbook.first;
  console.log(sheet.maps);
  console.log(sheet.data);
});
