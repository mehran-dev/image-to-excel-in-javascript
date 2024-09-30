const ExcelJS = require("exceljs");

// Create a new workbook and a worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("My Sheet");

// Add column headers
worksheet.columns = [
  { header: "ID", key: "id", width: 10 },
  { header: "Name", key: "name", width: 30 },
  { header: "Age", key: "age", width: 10 },
  { header: "Location", key: "location", width: 20 },
];

// Add some rows
worksheet.addRow({ id: 1, name: "John Doe", age: 28, location: "New York" });
worksheet.addRow({
  id: 2,
  name: "Jane Smith",
  age: 32,
  location: "San Francisco",
});

// Apply styles (optional)
worksheet.getRow(1).font = { bold: true }; // Bold headers
worksheet.getColumn(2).alignment = { horizontal: "left" }; // Align Name column to the left

// Save the file
workbook.xlsx
  .writeFile("my-excel-file.xlsx")
  .then(() => {
    console.log("Excel file created successfully!");
  })
  .catch((error) => {
    console.error("Error creating Excel file:", error);
  });
