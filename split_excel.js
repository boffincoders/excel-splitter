const XLSX = require("xlsx");
const path = require("path");

function splitExcelFile(filePath, linesPerFile = 1000) {
  // Read the Excel file
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the sheet to JSON format (array of arrays)
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // Extract the header row and the data rows
  const header = jsonData[0];
  const data = jsonData.slice(1); // All rows except the header

  // Calculate the number of files needed
  const totalFiles = Math.ceil(data.length / linesPerFile);

  // Loop through and create new Excel files
  for (let i = 0; i < totalFiles; i++) {
    // Slice the data for the current file
    const start = i * linesPerFile;
    const end = start + linesPerFile;
    const currentData = data.slice(start, end);

    // Combine the header with the current data
    const newSheetData = [header, ...currentData];

    // Create a new worksheet and workbook
    const newWorksheet = XLSX.utils.aoa_to_sheet(newSheetData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

    // Write the new workbook to a file
    const newFileName = path.join(__dirname, `output_file_${i + 1}.xlsx`);
    XLSX.writeFile(newWorkbook, newFileName);
    console.log(`Created file: ${newFileName}`);
  }
}

const inputFilePath = path.join(__dirname, "input_file.xlsx");
splitExcelFile(inputFilePath, 10000);
