// @ts-nocheck
// @ts-ignore
// @ts-nocheck
// @ts-ignore
import { IncomingForm } from 'formidable';
import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { sendProgressUpdate } from '../progress';

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
    externalResolver: true,
  },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const form = new IncomingForm();
  const uploadDir = path.join(process.cwd(), 'tmp');

  if (!fsSync.existsSync(uploadDir)) {
    fsSync.mkdirSync(uploadDir, { recursive: true });
  }

  form.uploadDir = uploadDir;
  form.keepExtensions = true;

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error('Error parsing form:', err);
      sendProgressUpdate(0, 'Error parsing form');
      return res.status(500).json({ error: 'Error parsing form' });
    }

    const mainFile = files.mainFile?.[0] || files.mainFile;
    const variantFile = files.variantFile?.[0] || files.variantFile;

    if (!mainFile || !variantFile) {
      sendProgressUpdate(0, 'Missing file uploads');
      return res.status(400).json({ error: 'Missing file uploads' });
    }

    const mainFilePath = mainFile.filepath;
    const variantFilePath = variantFile.filepath;

    if (!mainFilePath || !variantFilePath) {
      sendProgressUpdate(0, 'Error obtaining file paths');
      return res.status(500).json({ error: 'Error obtaining file paths' });
    }

    try {
      const mainWorkbook = new ExcelJS.Workbook();
      const variantWorkbook = new ExcelJS.Workbook();

      await mainWorkbook.xlsx.readFile(mainFilePath);
      sendProgressUpdate(10, 'Main file read successfully');

      await variantWorkbook.xlsx.readFile(variantFilePath);
      sendProgressUpdate(20, 'Variant file read successfully');

      const mainWorksheet = mainWorkbook.worksheets[0];
      const variantWorksheet = variantWorkbook.worksheets[0];

      const resultWorkbook = new ExcelJS.Workbook();
      const mainWorksheetCopy = resultWorkbook.addWorksheet('Main Worksheet');
      const variantWorksheetCopy = resultWorkbook.addWorksheet('Variant Worksheet');
      const commentsWorksheet = resultWorkbook.addWorksheet('Comments');

      // Copy main and variant worksheets
      mainWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          mainWorksheetCopy.getCell(rowNumber, colNumber).value = cell.value;
        });
      });
      sendProgressUpdate(30, 'Main worksheet copied');

      variantWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          variantWorksheetCopy.getCell(rowNumber, colNumber).value = cell.value;
        });
      });
      sendProgressUpdate(40, 'Variant worksheet copied');

      mainWorksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 0) {
          let changedCells = [];
          for (let colNumber = 1; colNumber <= mainWorksheet.columnCount; colNumber++) {
            const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
            const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;
            if (mainValue !== variantValue) {
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
                fgColor: { argb: 'FFFFFF00' },
              };
            });
          }
        }
      });
      sendProgressUpdate(60, 'Compared main and variant worksheets');

      mainWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber < 1) return;

          const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
          const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;
          const cellAddress = columnToLetter(colNumber) + rowNumber;
          const variantCell = variantWorksheetCopy.getCell(rowNumber, colNumber);
          variantCell.value = variantValue;

          if (mainValue !== variantValue) {
            variantCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFF0000' },
            };
          } else {
            variantCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF00FF00' },
            };
          }
        });
      });
      
      ['Main Worksheet', 'Variant Worksheet', 'Comments'].forEach(sheetName => {
        const worksheet = resultWorkbook.getWorksheet(sheetName);
        worksheet.columns.forEach((column, columnIndex) => {
          let maxLength = 0;
          column.eachCell((cell, rowIndex) => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) maxLength = cellLength;
          });
          worksheet.getColumn(columnIndex + 1).width = maxLength + 2;
        });
      });
      
      sendProgressUpdate(90, 'Prepared result workbook');

      const resultBuffer = await resultWorkbook.xlsx.writeBuffer();
      const resultFilePath = path.join(process.cwd(), 'public', 'Comparison Results.xlsx');
      await fs.writeFile(resultFilePath, resultBuffer);
      sendProgressUpdate(100, 'Comparison completed');

      try {
        await fs.unlink(mainFilePath);
        await fs.unlink(variantFilePath);
      } catch (err) {
        console.error('Error deleting temporary files:', err);
      }

      const downloadLink = `${req.headers.origin}/Comparison Results.xlsx`;
      return res.status(200).json({ downloadLink });
    } catch (error) {
      console.error('Error processing Excel files:', error);
      sendProgressUpdate(0, 'Error processing Excel files');
      return res.status(500).json({ error: 'Error processing Excel files' });
    }
  });
}
