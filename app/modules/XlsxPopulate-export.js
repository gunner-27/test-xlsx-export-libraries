const XlsxPopulate = require("xlsx-populate");

const data = require("../data/mock-data");

// Load a new blank workbook
const create = XlsxPopulate.fromBlankAsync().then((workbook) => {
  const sheet1 = workbook.sheet(0).name(data.sheet1.name);
  const sheet2 = workbook.addSheet(data.sheet2.name, 1);

  const sheets = [sheet1, sheet2];

  const setHeaders = (sheets, headers) => {
    sheets.forEach((s, i) => {
      headers[i].forEach((h, i) => {
        s.column(i + 1).width(h.width);
        s.row(1)
          .cell(i + 1)
          .value(h.text);
      });
    });
  };

  setHeaders(sheets, [data.sheet1.headers, data.sheet2.headers]);

  const parseData = (sheets, data) => {
    sheets.forEach((sheet, i) => {
      let values = [];
      data[i].forEach((obj) => {
        let row = Object.values(obj);
        values.push(row);
      });
      sheet.cell("A2").value(values);
    });
  };

  parseData(sheets, [data.sheet1.data, data.sheet2.data]);

  const findValue = (sheets, sheetsData) => {
    sheets.forEach((sheet, i) => {
      sheet._rows.forEach((row) => {
        if (row.cell(sheetsData[i].columnNum).value() == sheetsData[i].value) {
          row._cells.forEach((cell) => {
            cell.style("fill", "ff0000");
          });
        }
      });
    });
  };

  findValue(sheets, [
    { columnNum: 5, value: "Архивирован" },
    { columnNum: 2, value: "Мертв" },
  ]);

  sheet1.column(6).style({ fontColor: "ffe55ffe" });

  sheet2.row(1)._cells.forEach((cell) => {
    cell.style("fill", data.sheet2.styles.headers.headersColors.shift());
  });

  const stylingHeaders = (sheets, headersStyles) => {
    sheets.forEach((sheet, i) => {
      prepareStyles = {
        bold: headersStyles[i].font.bold,
        fontSize: headersStyles[i].font.size,
        fontFamily: headersStyles[i].font.name,
        fontColor: "000000",
        horizontalAlignment: headersStyles[i].alignment.horizontal,
        verticalAlignment: headersStyles[i].alignment.vertical,
      };
      sheet.row(1).style(prepareStyles);
    });
  };

  stylingHeaders(sheets, [
    data.sheet1.styles.headers,
    data.sheet2.styles.headers,
  ]);

  sheet2.range("A1:E1").autoFilter();

  // Write to file.
  return workbook.toFileAsync(
    `./samples/XlsxPopulate-example-${new Date().toISOString()}.xlsx`
  );
});

module.exports = create;
