const express = require("express");
const path = require("path");
const ExcelJs = require("exceljs");

// const { convertJsonToExcel } = require("./excelExport");
const users = [
  { name: "Raj", email: "r@g.com", age: 23 },
  { name: "Rahul", email: "rh@g.com", age: 32 },
];

const app = express();
const port = 3000;

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.get("/export", async (req, res) => {
  try {
    const workbook = new ExcelJs.Workbook();
    const worksheet = workbook.addWorksheet("My Users");
    worksheet.columns = [
      { header: "S_no", key: "s_no", width: 10 },
      { header: "Name", key: "name", width: 10 },
      { header: "Email", key: "email", width: 10 },
      { header: "Age", key: "age", width: 10 },
    ];

    // adding S_no manually by incrementing the count value and assigning to to S_no column
    let count = 1;
    users.forEach((user) => {
      user.s_no = count;
      worksheet.addRow(user);
      count += 1;
    });

    // making first row columns bold
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
    });

    // const data = await workbook.xlsx.writeFile("users.xlsx");
    // const data = await workbook.xlsx.writeFile("users.xlsx", res);

    // const data = await workbook.xlsx.write(res);
    await workbook.xlsx.write(res);
    // const data = await workbook.xlsx.write(res);
    // await workbook.xlsx.writeFile("users.xlsx", res);
    // res.send("done");
    // res.send(data);
  } catch (e) {
    res.status(500).send(e);
  }
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
