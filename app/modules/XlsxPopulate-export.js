const XlsxPopulate = require("xlsx-populate");

// Load a new blank workbook
const create = XlsxPopulate.fromBlankAsync().then((workbook) => {
  // Modify the workbook.
  workbook.sheet("Sheet1").cell("A1").value("This is neat!");

  // Write to file.
  return workbook.toFileAsync(
    `./samples/XlsxPopulate-example-${new Date().toISOString()}.xlsx`
  );
});

module.exports = create;
