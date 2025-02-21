const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 5000;

// Path to your Excel file
const EXCEL_FILE_PATH = path.join(__dirname, './Book1.xlsx');

// Middleware
app.use(bodyParser.json());
app.use(cors({
  origin: 'http://localhost:3000', // Allow requests only from your React app
}));

// API to fetch spreadsheet data
app.get('/api/getSpreadsheetData', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      return res.status(404).json({ error: 'Excel file not found.' });
    }

    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
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

// API to save spreadsheet data
app.post('/api/saveSpreadsheetData', (req, res) => {
  try {
    const jsonData = req.body.data;

    if (!jsonData || jsonData.length === 0) {
      return res.status(400).json({ error: 'Invalid or empty data provided.' });
    }

    // Assuming jsonData contains all rows including headers
    const headers = Object.keys(jsonData[0]);
    // console.log(headers);
    
    // Convert the data into a sheet format
    const dataArray = jsonData.map((row) => headers.map((header) => row[header] || ''));
    // console.log(dataArray[0]);

    const worksheet = xlsx.utils.aoa_to_sheet([...dataArray]);

    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Write the updated data to the Excel file
    xlsx.writeFile(workbook, EXCEL_FILE_PATH);

    res.json({ success: true, message: 'Data saved successfully.' });
  } catch (error) {
    console.error('Error saving Excel file:', error);
    res.status(500).json({ error: 'Error saving Excel file.' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
