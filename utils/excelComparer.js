const ExcelJS = require('exceljs');
const { promises: fs } = require('fs');
const path = require('path'); 

async function compareExcelFiles(mainFilePath, variantFilePath) {
  try {
    // Load the Excel files
    const mainWorkbook = new ExcelJS.Workbook();
    const variantWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.readFile(mainFilePath);
    await variantWorkbook.xlsx.readFile(variantFilePath);

    // Get the worksheets (assuming only one sheet in each workbook)
    const mainWorksheet = mainWorkbook.worksheets[0];
    const variantWorksheet = variantWorkbook.worksheets[0];

    // Create a new workbook for the results
    const resultWorkbook = new ExcelJS.Workbook();
    const mainWorksheetCopy = resultWorkbook.addWorksheet('Main Worksheet');
    const variantWorksheetCopy = resultWorkbook.addWorksheet('Variant Worksheet');
    const commentsWorksheet = resultWorkbook.addWorksheet('Comments');

    // Copy main and variant worksheets
    copyWorksheet(mainWorksheet, mainWorksheetCopy);
    copyWorksheet(variantWorksheet, variantWorksheetCopy);

    // Iterate through the rows to find changes and highlight them
    await mainWorksheet.eachRow({ includeEmpty: true }, async (row, rowNumber) => {
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

    await mainWorksheet.eachRow({ includeEmpty: true }, async (row, rowNumber) => {
      const companyName = row.getCell(1).value;
      let comments = [];

      // Indicate the row number in a separate column
      mainWorksheetCopy.getCell(rowNumber, 1).value = `Row ${rowNumber}`;

      await row.eachCell({ includeEmpty: true }, async (cell, colNumber) => {
        if (colNumber < 0) return; 

        const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
        const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;

        const cellAddress = colNumber === 0 ? 'A' : columnToLetter(colNumber);
        const variantCell = variantWorksheetCopy.getCell(`${cellAddress}${rowNumber}`);
        variantCell.value = variantValue;

        if (mainValue !== variantValue) {
          // Highlight the change in red
          variantCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF0000' } // Red color
          };
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

    // Save the results in a new workbook
    const resultFilePath = path.join(__dirname, '../../public', 'result.xlsx'); // Save in the public folder
    await resultWorkbook.xlsx.writeFile(resultFilePath);
    
    return resultFilePath;

  } catch (error) {
    console.error('Error comparing Excel files:', error);
    throw error; // Re-throw to let the API route handle it
  }
}

function copyWorksheet(sourceWorksheet, targetWorksheet) {
  sourceWorksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      targetWorksheet.getCell(rowNumber, colNumber).value = cell.value;
    });
  });
}

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

module.exports = compareExcelFiles;