const XLSX = require('xlsx');
const fs = require('fs');

// Get a list of all the Excel files in the current directory
const excelFiles = fs.readdirSync('.')
  .filter(filename => filename.endsWith('.xlsx'));

// Create a new workbook to hold all the sheets
const combinedWorkbook = XLSX.utils.book_new();

// Loop through each Excel file
excelFiles.forEach(file => {
  // Read the workbook
  const workbook = XLSX.readFile(file);

  // Loop through each sheet in the workbook
  workbook.SheetNames.forEach((sheetName) => {
    if (sheetName === "to_copy") {
        const sheet = workbook.Sheets[sheetName];

        // Copy the sheet to the combined workbook
        XLSX.utils.book_append_sheet(combinedWorkbook, sheet, `${file} :: ${sheetName}`);
    }
  });
});

// Write the combined workbook to a new file
XLSX.writeFile(combinedWorkbook, 'combined.xlsx');
