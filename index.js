const express = require("express");
const path = require("path");

const { exportExcel } = require("./excelExport");

const app = express();
const port = 3000;

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.get("/export", async (req, res) => {
  try {
    const originalData = [
      {
        clientName: "Raj",
        customerEmail: "r@g.com",
        age: 23,
        Admin: {
          name: "Aayush",
          age: 23,
        },
        Student: {
          Bachelor: {
            name: "Ram",
          },
        },
      },
      {
        clientName: "Rahul",
        customerEmail: "rh@g.com",
        age: 32,
        Admin: {
          name: "Alex",
          age: 45,
        },
        Student: {
          Bachelor: {
            name: "Sita",
          },
        },
      },
    ];

    let data = [];

    originalData.forEach((info, index) => {
      let { Admin, Student, ...rest } = info;
      let dataObj = {
        ...rest,
        adminName: Admin.name,
        adminAge: Admin.age,
        studentName: Student.Bachelor.name,
      };
      data.push(dataObj);
    });

    const columns = [
      { header: "S_no", key: "s_no", width: 20 },
      { header: "Client Name", key: "clientName", width: 20 },
      { header: "Customer Email", key: "customerEmail", width: 20 },
      { header: "Age", key: "age", width: 20 },
      { header: "Admin Name", key: "adminName", width: 20 },
      { header: "Admin Age", key: "adminAge", width: 20 },
      { header: "Student Name", key: "studentName", width: 20 },
    ];

    exportExcel(data, columns, res);
  } catch (e) {
    res.status(500).send(e);
  }
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
