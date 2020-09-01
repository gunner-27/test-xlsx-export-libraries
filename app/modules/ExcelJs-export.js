const Excel = require("exceljs");

const data = require("../data/mock-data");

// Create workbook & add worksheets
const workbook = new Excel.Workbook();
const sheet1 = workbook.addWorksheet(data.sheet1.name);
const sheet2 = workbook.addWorksheet(data.sheet2.name);
const sheets = [sheet1, sheet2];

// Add column headers
const prepareHeaders = function (data) {
  let columnsNames = [];
  data.forEach((h) => {
    columnsNames.push({
      header: h.text,
      key: h.key,
      width: h.width,
    });
  });
  return columnsNames;
};

sheet1.columns = prepareHeaders(data.sheet1.headers);
sheet2.columns = prepareHeaders(data.sheet2.headers);

// Add rows
const prepareDate = function (data, keys) {
  let rows = [];
  data.forEach((obj) => {
    let prepArr = [];
    keys.forEach((key) => {
      prepArr.push(obj[key]);
    });
    rows.push(prepArr);
  });
  return rows;
};

sheet1.addRows(
  prepareDate(data.sheet1.data, [
    "FIO",
    "date",
    "num",
    "text",
    "status",
    "colored_num",
  ])
);

sheet2.addRows(
  prepareDate(data.sheet2.data, ["FIO", "status", "date", "text", "sex"])
);

// Styling
sheet1.getColumn("colored_num").font = {
  color: { argb: "ffe55ffe" },
};

const findAndStyleRows = (sheets, valuesInfos) => {
  sheets.forEach((sheet, i) => {
    sheet.eachRow((row) => {
      if (row.getCell(valuesInfos[i].column) == valuesInfos[i].value) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" },
          };
        });
      }
    });
  });
};

findAndStyleRows(sheets, [
  { column: "status", value: "Архивирован" },
  { column: "status", value: "Мертв" },
]);

sheet2.getRow(1).eachCell((cell, i) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: `${data.sheet2.styles.headers.headersColors[i - 1]}` },
  };
});

const stylingHeaders = function (sheets) {
  sheets.forEach((sheet, i) => {
    sheet.getRow(1).font = data[`sheet${i + 1}`].styles.headers.font;
    sheet.getRow(1).alignment = data[`sheet${i + 1}`].styles.headers.alignment;
  });
};

stylingHeaders(sheets);

// Add auto filter
sheet2.autoFilter = "A1:E1";

// Save workbook to disk
const create = workbook.xlsx
  .writeFile(`./samples/ExcelJs-example-${new Date().toISOString()}.xlsx`)
  .catch((err) => {
    console.log("err", err);
  });

module.exports = create;
