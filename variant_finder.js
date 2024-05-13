ExcelJS = require('exceljs');
const fs = require('fs');

// Function to convert column number to letter
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Load the Excel files
console.log('Loading Excel files...');
const mainWorkbook = new ExcelJS.Workbook();
const variantWorkbook = new ExcelJS.Workbook();
mainWorkbook.xlsx.readFile('main.xlsx').then(mainFile => {
  variantWorkbook.xlsx.readFile('variant.xlsx').then(variantFile => {
    console.log('Excel files loaded successfully.');

    // Get the worksheet names
    console.log('Getting worksheet names...');
    const mainWorksheet = mainFile.worksheets[0];
    const variantWorksheet = variantFile.worksheets[0];
    console.log('Worksheet names retrieved.');


    // Create a new workbook for the results
    console.log('Creating new workbook for the results...');
    const resultWorkbook = new ExcelJS.Workbook();
    const mainWorksheetCopy = resultWorkbook.addWorksheet('Main Worksheet');
    const variantWorksheetCopy = resultWorkbook.addWorksheet('Variant Worksheet');
    const commentsWorksheet = resultWorkbook.addWorksheet('Comments');
    console.log('New workbook created.');

    // Copy main and variant worksheets
    mainWorksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        mainWorksheetCopy.getCell(rowNumber, colNumber).value = cell.value;
      });
    });

    variantWorksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        variantWorksheetCopy.getCell(rowNumber, colNumber).value = cell.value;
      });
    });
    
    
    
    mainWorksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 0) { 
        let changedCells = [];
        for (let colNumber = 7; colNumber <= mainWorksheet.columnCount; colNumber++) {
          if (mainWorksheet.getCell(rowNumber, colNumber).value !== variantWorksheet.getCell(rowNumber, colNumber).value) {
            changedCells.push(columnToLetter(colNumber) + rowNumber);
          }
        }
        if (changedCells.length > 0) {
          commentsWorksheet.addRow([`Row ${rowNumber}`, `${changedCells.join(', ')} has changed`]);
        } else {
          const row = commentsWorksheet.addRow([`Row ${rowNumber}`, `No change in row ${rowNumber}`]);
          row.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF00' } // Yellow color
            };
          });
      }
    }
    });
    // Iterate through the rows to find changes and highlight them
    console.log('Iterating through the rows to find changes...');
    mainWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      // Get the company name
      const companyName = row.getCell(1).value;
      let comments = [];

      // Indicate the row number in a separate column
      mainWorksheetCopy.getCell(rowNumber, 1).value = `Row ${rowNumber}`;

      // Iterate through the columns from D onwards
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber < 0) return; // Skip columns A-C

        // Get the values from the main and variant worksheets
        const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
        const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;

        // Check if values are different and highlight them
        const cellAddress = colNumber === 0 ? 'A' : columnToLetter(colNumber);
        const variantCell = variantWorksheetCopy.getCell(`${cellAddress}${rowNumber}`);
        variantCell.value = variantValue;

        if (mainValue !== variantValue) {
          // Log the change
          // Highlight the change in red
          variantCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF0000' } // Red color
          };

          // Add the change to the comments worksheet
          comments.push(`${cellAddress}${rowNumber}`);
        } else {
          // Highlight matching cells in green
          variantCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00FF00' } // Green color
          };
        }
      });

    });
    console.log('Rows iteration completed.');

    // Save the results in a new workbook
    console.log('Saving the results in a new workbook...');
    resultWorkbook.xlsx.writeFile('result.xlsx').then(() => {
      console.log('Results saved successfully in a new workbook.');
    });
  });
});