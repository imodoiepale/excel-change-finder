import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { IncomingForm } from 'formidable';

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

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  console.log("Request received in /api/excel handler"); // DEBUG

  try {
    const form = new IncomingForm();
    form.parse(req, async (err, fields, files) => {
      if (err) {
        console.error('Error parsing form:', err); // DEBUG
        return res.status(500).json({ error: 'Error parsing form' });
      }

      console.log("Files received:", files); // DEBUG - Check if files exist

      const mainFile = files.mainFile;
      const variantFile = files.variantFile;

      if (!mainFile || !variantFile) {
        console.error("Missing mainFile or variantFile"); // DEBUG
        return res.status(400).json({ error: 'Missing file uploads' });
    }

      // Create temporary file paths
    const mainFilePath = path.join(process.cwd(), 'tmp', 'main.xlsx');
    const variantFilePath = path.join(process.cwd(), 'tmp', 'variant.xlsx');
    console.log('Main File Path:', mainFile.filepath);
    console.log('Variant File Path:', variantFile.filepath);

    try {
        console.log("Saving files to:", mainFilePath, variantFilePath); // DEBUG
        // Save the uploaded files to the temporary file paths
        try {
            await fs.promises.writeFile(mainFilePath, mainFile.file);
            await fs.promises.writeFile(variantFilePath, variantFile.file);
        } catch (err) {
            console.error('Error writing files:', err);
            return res.status(500).json({ error: 'Error writing temporary files' });
        }

        console.log('Loading Excel files...'); // DEBUG
        // Load the Excel files
        const mainWorkbook = new ExcelJS.Workbook();
        const variantWorkbook = new ExcelJS.Workbook();

        await mainWorkbook.xlsx.readFile(mainFilePath);
        await variantWorkbook.xlsx.readFile(variantFilePath);

        console.log('Excel files loaded successfully.'); // DEBUG

        // Get the worksheet names
        console.log('Getting worksheet names...'); // DEBUG
        const mainWorksheet = mainWorkbook.worksheets[0];
        const variantWorksheet = variantWorkbook.worksheets[0];
        console.log('Worksheet names retrieved:', mainWorksheet.name, variantWorksheet.name); // DEBUG

        // Create a new workbook for the results
        console.log('Creating new workbook for the results...'); // DEBUG
        const resultWorkbook = new ExcelJS.Workbook();
        const mainWorksheetCopy = resultWorkbook.addWorksheet('Main Worksheet');
        const variantWorksheetCopy = resultWorkbook.addWorksheet('Variant Worksheet');
        const commentsWorksheet = resultWorkbook.addWorksheet('Comments');
        console.log('New workbook created.'); // DEBUG

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

        // Find changes and add to comments worksheet
        mainWorksheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1) { // Start from the second row (assuming the first row is headers)
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
        console.log('Iterating through the rows to find changes...'); // DEBUG
        mainWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          // Get the company name (assuming it's in the first column)
          const companyName = row.getCell(1).value;
          let comments = [];

          // Indicate the row number in a separate column
          mainWorksheetCopy.getCell(rowNumber, 1).value = `Row ${rowNumber}`;

          // Iterate through the columns from D onwards
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber < 4) return; // Skip columns A-C

            // Get the values from the main and variant worksheets
            const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
            const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;

            // Check if values are different and highlight them
            const cellAddress = columnToLetter(colNumber); 
            const variantCell = variantWorksheetCopy.getCell(`${cellAddress}${rowNumber}`);
            variantCell.value = variantValue;

            if (mainValue !== variantValue) {
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
        console.log('Rows iteration completed.'); // DEBUG

        // Save the results in a new workbook
        console.log('Saving the results in a new workbook...'); // DEBUG
        const resultBuffer = await resultWorkbook.xlsx.writeBuffer();
        const resultFileName = 'result.xlsx';

        // Set the appropriate headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${resultFileName}"`);

        // Send the result file as the response
        res.status(200).send(resultBuffer);

      } catch (error) {
        console.error('Error processing Excel files:', error); // DEBUG
        res.status(500).json({ error: 'Error processing Excel files' });
      } finally {
        // Remove the temporary files
        try {
          console.log("Deleting temporary files:", mainFilePath, variantFilePath); // DEBUG
          try {
            if (fs.existsSync(mainFilePath)) {
              await fs.promises.unlink(mainFilePath);
            }
            if (fs.existsSync(variantFilePath)) {
              await fs.promises.unlink(variantFilePath);
            }
          } catch (err) {
            console.error('Error deleting temporary files:', err);
          }
        } catch (err) {
          console.error('Error deleting temporary files:', err); // DEBUG
        }
      }
    });
  } catch (outerError) {
    console.error('Unexpected error:', outerError); // DEBUG - Catch any errors outside the form parsing
    res.status(500).json({ error: 'An unexpected error occurred' });
  }
}