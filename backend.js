const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 5001;

// Middleware
app.use(bodyParser.json());
app.use(cors({
  origin: ['https://vwg-frontend.vercel.app', 'http://localhost:3000' , 'https://vwg-frontend-alpha.vercel.app/']
}));

// Helper function to get the Excel file path based on year and file number
const getExcelFilePath = (year, fileNumber) => {
  return path.join(__dirname, `./${year}_${fileNumber}.xlsx`);
};

// API to fetch years (fetching distinct years from available files)
app.get('/api/getYears', (req, res) => {
  try {
    const years = [];
    const files = fs.readdirSync(__dirname).filter(file => file.match(/^\d{4}_\d+\.xlsx$/)); // Match files like 2024_1.xlsx
    
    files.forEach(file => {
      const year = file.split('_')[0];
      if (!years.includes(year)) {
        years.push(year);
      }
    });
    
    res.json({ years });
  } catch (error) {
    console.error('Error fetching years:', error);
    res.status(500).json({ error: 'Error fetching years.' });
  }
});

// API to fetch files for a specific year (e.g., 2024_1, 2024_2, etc.)
app.get('/api/getFiles/:year', (req, res) => {
  const { year } = req.params;
  
  try {
    const files = fs.readdirSync(__dirname).filter(file => file.startsWith(`${year}_`) && file.endsWith('.xlsx'));

    if (files.length === 0) {
      return res.status(404).json({ error: `No files found for year ${year}.` });
    }

    res.json({ files });
  } catch (error) {
    console.error('Error fetching files:', error);
    res.status(500).json({ error: 'Error fetching files.' });
  }
});

// API to fetch data for a specific file (e.g., 2024_1.xlsx)
app.get('/api/getPBUData/:file', (req, res) => {
  const { file } = req.params;
  
  try {
    const EXCEL_FILE_PATH = path.join(__dirname, file);

    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      return res.status(404).json({ error: `Excel file ${file} not found.` });
    }

    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]; // Default to first sheet
    
    // Get the range of cells
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    
    // Ensure empty cells are included in the JSON output
    const jsonData = [];
    
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      const rowData = {};
      
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellAddress];
        
        rowData[`Column${colNum + 1}`] = cell ? cell.v : ''; // Retain empty cells as empty strings
      }
      
      jsonData.push(rowData);
    }

    res.json({ data: jsonData });
  } catch (error) {
    console.error('Error fetching PBU data:', error);
    res.status(500).json({ error: 'Error fetching PBU data.' });
  }
});

// API to get the spreadsheet data (added for creating a new PBU template)
app.get('/api/getSpreadsheetData', (req, res) => {
  try {
    // Path to the Excel file (use a fixed file or dynamic based on your needs)
    const EXCEL_FILE_PATH = path.join(__dirname, './Book1.xlsx'); // Update this path if needed

    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      return res.status(404).json({ error: 'Excel file not found.' });
    }

    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]; // Fetching the first sheet
    
    // Get the range of cells
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    
    // Ensure empty cells are included in the JSON output
    const jsonData = [];
    
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      const rowData = {};
      
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellAddress];
        
        rowData[`Column${colNum + 1}`] = cell ? cell.v : ''; // Retain empty cells as empty strings
      }
      
      jsonData.push(rowData);
    }

    res.json({ data: jsonData });
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ error: 'Error reading Excel file.' });
  }
});

// API to save data to a specific file (e.g., 2024_1.xlsx)
// app.post('/api/saveFiles/:year/:file', (req, res) => {
//   const { year, file } = req.params;
  
//   try {
//     const jsonData = req.body.data;

//     if (!jsonData || jsonData.length === 0) {
//       return res.status(400).json({ error: 'Invalid or empty data provided.' });
//     }

//     const EXCEL_FILE_PATH = path.join(__dirname, `${year}_${file}.xlsx`);

//     if (!fs.existsSync(EXCEL_FILE_PATH)) {
//       return res.status(404).json({ error: `Excel file ${year}_${file}.xlsx not found.` });
//     }

//     // Assuming jsonData contains all rows including headers
//     const headers = Object.keys(jsonData[0]);
    
//     // Convert the data into a sheet format
//     const dataArray = jsonData.map((row) => headers.map((header) => row[header] || ''));

//     const worksheet = xlsx.utils.aoa_to_sheet(dataArray);
//     const workbook = xlsx.utils.book_new();

//     xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//     xlsx.writeFile(workbook, EXCEL_FILE_PATH);

//     res.json({ success: true, message: `Data saved to ${year}_${file}.xlsx successfully.` });
//   } catch (error) {
//     console.error('Error saving Excel file:', error);
//     res.status(500).json({ error: 'Error saving Excel file.' });
//   }
// });
// API to save data to a specific file (create a new file if necessary)
app.post('/api/saveFiles/:year/:file', (req, res) => {
  const { year, file } = req.params;
  try {
    const jsonData = req.body.data;

    if (!jsonData || jsonData.length === 0) {
      return res.status(400).json({ error: 'Invalid or empty data provided.' });
    }

    let EXCEL_FILE_PATH;

    // If it's a new file, create a new one
    if (file === 'newFile') {
      const newFileName = `2025_new_1.xlsx`;
      EXCEL_FILE_PATH = path.join(__dirname, newFileName);
    } else {
      EXCEL_FILE_PATH = path.join(__dirname, `${file}`);
    }

    // Create or overwrite the Excel file
    const headers = Object.keys(jsonData[0]);
    const dataArray = jsonData.map((row) => headers.map((header) => row[header] || ''));

    const worksheet = xlsx.utils.aoa_to_sheet(dataArray);
    const workbook = xlsx.utils.book_new();

    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, EXCEL_FILE_PATH);

    res.json({ success: true, message: `Data saved to ${EXCEL_FILE_PATH}` });
  } catch (error) {
    console.error('Error saving Excel file:', error);
    res.status(500).json({ error: 'Error saving Excel file.' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});

module.exports = app;
