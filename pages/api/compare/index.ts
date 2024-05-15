// @ts-nocheck
//@ts-ignore

// pages/api/compare/index.ts
import { NextApiRequest, NextApiResponse } from 'next';
import { promises as fs } from 'fs';
import { createRouter } from 'next-connect';
import { tmpdir } from 'os'; // Use os.tmpdir() for a temporary directory

import compareExcelFiles from '../../../utils/excelComparer';

const api = createRouter<NextApiRequest, NextApiResponse>();



api.post(async (req, res) => {
  try {
    const mainFile = req.files.mainFile;
    const variantFile = req.files.variantFile;

    // Create temporary file paths in the temporary directory
    const tempDir = tmpdir();
    const mainTempPath = path.join(tempDir, mainFile.originalname);
    const variantTempPath = path.join(tempDir, variantFile.originalname);

    // Move uploaded files to temporary locations
    await fs.writeFile(mainTempPath, mainFile.buffer);
    await fs.writeFile(variantTempPath, variantFile.buffer);

    // Invoke the comparison logic
    const resultFilePath = await compareExcelFiles(mainTempPath, variantTempPath);

    // Return the result file path
    res.status(200).json({ resultFilePath });
  } catch (error) {
    console.error('Error comparing Excel files:', error);
    res.status(500).json({ error: 'Internal server error' });
  } finally {
    // Remove temporary files
    await fs.unlink(mainTempPath).catch(() => {});
    await fs.unlink(variantTempPath).catch(() => {});
  }
});

// Create the API handler with error handling
export default api.handler({
  onError: (err, req, res) => {
    console.error(err.stack);
    res.status(500).end(err.message);
  },
});

// Middleware for file uploads 
api.use((req, res, next) => {
  if (!req.files) {
    return res.status(400).json({ error: 'No files uploaded' });
  }
  next();
});