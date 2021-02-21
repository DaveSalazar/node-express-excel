const express = require("express");
const exceljs = require("exceljs");
const app = express();
const port = 3000;
const data = require("./sample_data.json");

// sample data
app.get("/", (request, response) => {
  const workbook = exportData(data);
  response.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  response.setHeader(
    "Content-Disposition",
    "attachment; filename=" + "data.xlsx"
  );

  return workbook.xlsx.write(response).then(function () {
    response.status(200).end();
  });
});

app.listen(port, () => {
  console.log(`Listening at http://localhost:${port}`);
});

const exportData = (data) => {
  let workbook = new exceljs.Workbook();
  let worksheet = workbook.addWorksheet("Worksheet");

  let columns = data.reduce((acc, obj) => acc = Object.getOwnPropertyNames(obj), [])  

  worksheet.columns = columns.map((el) => {
    return { header: el, key: el, width: 20 };
  });

  worksheet.addRows(data);

  return workbook;
};
