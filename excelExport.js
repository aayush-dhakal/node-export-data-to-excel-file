const ExcelJs = require("exceljs");

const exportExcel = async (data, columns, res) => {
  const workbook = new ExcelJs.Workbook();
  const worksheet = workbook.addWorksheet("My Users");
  // worksheet.columns = [
  //   { header: "S_no", key: "s_no", width: 10 },
  //   { header: "Name", key: "name", width: 10 },
  //   { header: "Email", key: "email", width: 10 },
  //   { header: "Age", key: "age", width: 10 },
  // ];

  worksheet.columns = columns;

  // adding S_no manually by incrementing the count value and assigning to to S_no column
  let count = 1;
  data.forEach((d) => {
    d.s_no = count;
    worksheet.addRow(d);
    count += 1;
  });

  // making first row columns bold
  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
  });

  // this downloads the excel file locally
  await workbook.xlsx.writeFile("users.xlsx");

  // this sends the excel data in response. Basically writes the excel data to response stream and so we got the data in response
  await workbook.xlsx.write(res);
};

module.exports = { exportExcel };
