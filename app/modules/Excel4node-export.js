// Require library
var xl = require("excel4node");

const data = require("../data/mock-data");

// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
var sheet1 = wb.addWorksheet(data.sheet1.name);
var sheet2 = wb.addWorksheet(data.sheet2.name);

const sheets = [sheet1, sheet2];

const setHeaders = (sheets, headers) => {
  sheets.forEach((s, i) => {
    headers[i].forEach((h, i) => {
      s.column(i + 1).setWidth(h.width);
      s.cell(1, i + 1).string(h.text);
    });
  });
};

setHeaders(sheets, [data.sheet1.headers, data.sheet2.headers]);

const parseData = (sheets, data) => {
  sheets.forEach((sheet, i) => {
    data[i].forEach((obj, i) => {
      Object.values(obj).forEach((val, j) => {
        switch (typeof val) {
          case "number":
            sheet
              .cell(i + 2, j + 1, i + 2, Object.values(obj).length)
              .number(val);
            break;
          default:
            sheet
              .cell(i + 2, j + 1, i + 2, Object.values(obj).length)
              .string(val.toString());
        }
      });
    });
  });
};

parseData(sheets, [data.sheet1.data, data.sheet2.data]);

findValue(sheets, [
  { columnNum: 5, value: "Архивирован" },
  { columnNum: 2, value: "Мертв" },
]);

const create = wb.write(
  `./samples/Excel4node-example-${new Date().toISOString()}.xlsx`
);

module.exports = create;
