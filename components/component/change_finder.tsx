// @ts-nocheck
// @ts-ignore
import { useEffect, useState } from 'react';
import Link from "next/link";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { Loader2 } from "lucide-react";
import toast, { Toaster } from 'react-hot-toast';
import * as ExcelJS from 'exceljs'; 

export function Change_Finder() {
  const [mainFile, setMainFile] = useState(null);
  const [variantFile, setVariantFile] = useState(null);
  const [progress, setProgress] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [isLoading2, setIsLoading2] = useState(false);
  const [consoleLog, setConsoleLog] = useState('');
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [downloadSuccess , setDownloadSuccess] = useState(null);

  const sendProgressUpdate = (progress, message) => {
    setProgress(progress);
    setConsoleLog(message);
  };

  const compareExcelFiles = async () => {
    setIsLoading(true);
    sendProgressUpdate(5, 'Processing files...');
  
    if (!mainFile || !variantFile) {
      toast.error('Please upload both main and variant files!')
      sendProgressUpdate(0, 'Please upload both main and variant files.');
      setIsLoading(false);
      return;
    }
    try {
    
      const mainWorkbook = new ExcelJS.Workbook();
      const variantWorkbook = new ExcelJS.Workbook();
  
      const mainFileBuffer = await mainFile.arrayBuffer();
      await mainWorkbook.xlsx.load(mainFileBuffer);
      sendProgressUpdate(20, 'Main file read successfully');
  
      const variantFileBuffer = await variantFile.arrayBuffer();
      await variantWorkbook.xlsx.load(variantFileBuffer);
      sendProgressUpdate(40, 'Variant file read successfully');

      const mainWorksheet = mainWorkbook.worksheets[0];
      const variantWorksheet = variantWorkbook.worksheets[0];

      const resultWorkbook = new ExcelJS.Workbook();
      const mainWorksheetCopy = resultWorkbook.addWorksheet('Main Worksheet');
      const variantWorksheetCopy = resultWorkbook.addWorksheet('Variant Worksheet');
      const commentsWorksheet = resultWorkbook.addWorksheet('Comments');

      // Copy main and variant worksheets
      mainWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const targetCell = mainWorksheetCopy.getCell(rowNumber, colNumber);
          targetCell.value = cell.value;
      
          // Copy styles
          targetCell.style = cell.style;
        });
      });
      
      sendProgressUpdate(50, 'Main worksheet copied');

      variantWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          variantWorksheetCopy.getCell(rowNumber, colNumber).value = cell.value;
        });
      });
      sendProgressUpdate(60, 'Variant worksheet copied');

      mainWorksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
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
      sendProgressUpdate(70, 'Compared main and variant worksheets');

      mainWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber < 1) return;

          const mainValue = mainWorksheet.getCell(rowNumber, colNumber).value;
          const variantValue = variantWorksheet.getCell(rowNumber, colNumber).value;
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
      sendProgressUpdate(90, 'Highlighted cells in variant worksheet');

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
      sendProgressUpdate(95, 'Adjusted column widths');

      const outputBuffer = await resultWorkbook.xlsx.writeBuffer();
      const outputBlob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const downloadUrl = URL.createObjectURL(outputBlob);

      sendProgressUpdate(100, 'Files processed successfully!');

      setDownloadUrl(downloadUrl);
      toast.success('Files Compared Successfully!')

    } catch (error) {
      console.error('Error processing files:', error);
      sendProgressUpdate(0, 'An error occurred while processing the files.');
      toast.error('FILE COMPARISON ERROR!')
    } finally {
      setIsLoading(false);
    }
  };

  const downloadExcelFile = async (downloadUrl) => {
    setIsLoading2(true); // Set isLoading to true when download starts
    try {
      // Your code to trigger file download
      // For example:
      const response = await fetch(downloadUrl);
      const blob = await response.blob();
      const url = window.URL.createObjectURL(new Blob([blob]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'COMPARED_RESULTS.xlsx');
      document.body.appendChild(link);
      link.click();
      link.parentNode.removeChild(link);
      setTimeout(() => {
        setIsLoading2(false);
      }, 3000);
      setDownloadSuccess(true);
    } catch (error) {
      console.error('Error downloading file:', error);
      setIsLoading2(false); // Set isLoading to false if download fails
    }
  };

  const columnToLetter = (columnIndex) => {
    let columnName = '';
    let dividend = columnIndex;

    while (dividend > 0) {
      const modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }

    return columnName;
  };

  return (
    <div className="flex min-h-screen flex-col">
      <Toaster position="top-right" reverseOrder={false} toastOptions={{ duration: 3000 }} />
      <header className="bg-gray-900 py-4 px-6 text-white">
        <div className="container mx-auto flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <FileSpreadsheetIcon className="h-6 w-6" />
            <h1 className="text-2xl font-bold">Excel Comparator</h1>
          </div>
        </div>
      </header>

      <main className="flex-1 bg-gray-100 py-12 px-6 dark:bg-gray-900">
        <div className="container mx-auto max-w-3xl space-y-8">
          <div className="space-y-4">
            <h2 className="text-3xl font-bold text-gray-900 dark:text-white">
              Compare Excel Files
            </h2>
            <p className="text-gray-600 dark:text-gray-400">
              Upload your main and variant Excel files to compare the changes.
            </p>
          </div>

          <div className="space-y-6">
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div>
                <Label htmlFor="main-file">Main File</Label>
                <Input
                  accept=".xlsx"
                  className="mt-1 block w-full"
                  id="main-file"
                  name="mainFile"
                  type="file"
                  onChange={(e) => setMainFile(e.target.files[0])}
                />
              </div>
              <div>
                <Label htmlFor="variant-file">Variant File</Label>
                <Input
                  accept=".xlsx"
                  className="mt-1 block w-full"
                  id="variant-file"
                  name="variantFile"
                  type="file"
                  onChange={(e) => setVariantFile(e.target.files[0])}
                />
              </div>
            </div>
            <Button
              className="w-full" onClick={compareExcelFiles}
              disabled={isLoading}
            >
              {isLoading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Comparing...
                </>
              ) : (
                "Compare Files"
              )}
            </Button>

            <div className="text-gray-900 dark:text-white">
              Comparison Progress:
            </div>
            <div className="space-y-2 text-center">
              <div className="text-gray-600 dark:text-gray-400">
                {consoleLog}
              </div>
              <Progress
                className="h-2 bg-gray-300 dark:bg-gray-800"
                value={progress}
              />
              <div className="text-gray-600 dark:text-gray-400">
                {progress}% Complete
              </div>
            </div>
            {progress === 100 && (
              <div className=" justify-center text-center flex space-x-6 mt-6">
                <div>
                {isLoading2 ? (
                    <Button className="w-full" disabled>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" /> Downloading...
                    </Button>
                  ) : downloadSuccess ? (
                    <Button className="w-full bg-green-700 hover:bg-green-600" onClick={() => setIsLoading2(false)}>Download Successful</Button>
                  ) : (
                    <Button className="w-full" onClick={() => downloadExcelFile(downloadUrl)}>Download Compared File</Button>
                  )}
                </div>
                <div>
                  <Button
                    className="w-full"
                    onClick={() => window.location.reload()}
                  >
                    Refresh Page
                  </Button>
                </div>
              </div>
            )}
          </div>
        </div>
      </main>

      <footer className="bg-gray-900 py-4 px-6 text-white">
        <div className="container mx-auto flex items-center justify-center">
          <p>
            © 2024 Excel Comparator. All rights reserved by{" "}
            <Link
              className="hover:underline"
              href="http://booksmartconsult.com/"
            >
              BCL
            </Link>
          </p>
        </div>
      </footer>
    </div>
  );
}

function FileSpreadsheetIcon(props) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      width="24"
      height="24"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" />
      <path d="M14 2v4a2 2 0 0 0 2 2h4" />
      <path d="M8 13h2" />
      <path d="M14 13h2" />
      <path d="M8 17h2" />
      <path d="M14 17h2" />
    </svg>
  );
}
