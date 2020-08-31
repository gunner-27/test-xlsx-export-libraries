const Excel = require("exceljs");

const data = require("../data/mock-data");

// Create workbook & add worksheet
const workbook = new Excel.Workbook();
const sheet1 = workbook.addWorksheet(data.sheet1.name);
const sheet2 = workbook.addWorksheet(data.sheet2.name);

sheet1.columns = data.sheet1.headers;

let columnsNames = [];

data.sheet1.headers.forEach((h) => {
  columnsNames.push({ header: h, key: h });
});
// add column headers
sheet1.columns = columnsNames;

columnsNames = [];

data.sheet2.headers.forEach((h) => {
  columnsNames.push({ header: h, key: h });
});
// add column headers
sheet2.columns = columnsNames;

let rows = [];
data.sheet1.data.forEach((obj) => {
  let prepArr = [];
  prepArr.push(
    obj.FIO,
    obj.date,
    obj.num,
    obj.text,
    obj.status,
    obj.colored_num
  );
  rows.push(prepArr);
});

sheet1.addRows(rows);

rows = [];
data.sheet2.data.forEach((obj) => {
  let prepArr = [];
  prepArr.push(obj.FIO, obj.status, obj.date, obj.text, obj.sex);
  rows.push(prepArr);
});

sheet2.addRows(rows);

// save workbook to disk
const create = workbook.xlsx
  .writeFile(`./samples/ExcelJs-example-${new Date().toISOString()}.xlsx`)
  .then(() => {
    console.log("created");
  })
  .catch((err) => {
    console.log("err", err);
  });

module.exports = create;
