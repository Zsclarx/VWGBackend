const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const postgres = require("postgres");
require("dotenv").config();

const app = express();
app.use(express.json());
app.use(cors());
const PORT = 5001;

// PostgreSQL Connection
const connectionString = process.env.DATABASE_URL;
const sql = postgres(connectionString);

// Middleware
app.use(bodyParser.json());
app.use(
  cors({
    origin: ["https://vwg-frontend.vercel.app", "http://localhost:3000"],
  })
);

// JWT Middleware
const authenticateToken = (req, res, next) => {
  const token = req.header("Authorization");
  if (!token) return res.status(401).json({ message: "Unauthorized" });

  jwt.verify(token, process.env.JWT_SECRET, (err, user) => {
    if (err) return res.status(403).json({ message: "Invalid token" });
    req.user = user;
    next();
  });
};

// Fetch spreadsheet data from local file (unchanged)
app.get("/api/getSpreadsheetData", (req, res) => {
  try {
    const EXCEL_FILE_PATH = path.join(__dirname, "./Book1.xlsx");
    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      return res.status(404).json({ error: "Excel file not found." });
    }

    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = xlsx.utils.decode_range(worksheet["!ref"]);
    const jsonData = [];

    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      const rowData = {};
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellAddress];
        rowData[`Column${colNum + 1}`] = cell ? cell.v : "";
      }
      jsonData.push(rowData);
    }

    res.json({ data: jsonData });
  } catch (error) {
    console.error("Error reading Excel file:", error);
    res.status(500).json({ error: "Error reading Excel file." });
  }
});

// Register API
app.post("/register", async (req, res) => {
  const { username, password } = req.body;
  const hashedPassword = await bcrypt.hash(password, 10);

  try {
    await sql`INSERT INTO users (username, password_hash) VALUES (${username}, ${hashedPassword})`;
    res.status(201).json({ message: "User registered successfully" });
  } catch (error) {
    console.error("Error registering user:", error);
    res.status(400).json({ message: "User already exists" });
  }
});

// Login API
app.post("/login", async (req, res) => {
  const { username, password } = req.body;

  try {
    const results = await sql`SELECT * FROM users WHERE username = ${username}`;
    if (results.length === 0) {
      return res.status(401).json({ message: "Invalid credentials" });
    }

    const user = results[0];
    const validPassword = await bcrypt.compare(password, user.password_hash);

    if (!validPassword) {
      return res.status(401).json({ message: "Invalid credentials" });
    }

    const token = jwt.sign(
      { id: user.id, username: user.username },
      process.env.JWT_SECRET,
      { expiresIn: "1h" }
    );

    res.json({ token });
  } catch (error) {
    console.error("Error during login:", error);
    res.status(500).json({ message: "Internal server error" });
  }
});

// Fetch user-specific Excel sheets
app.get("/excel-sheets", authenticateToken, async (req, res) => {
  const userId = req.user.id;
  try {
    const results = await sql`SELECT * FROM excel_sheets WHERE user_id = ${userId}`;
    res.json(results);
  } catch (error) {
    console.error("Error fetching excel sheets:", error);
    res.status(500).json({ message: "Database error" });
  }
});

// Fetch distinct years
app.get("/api/getYears", authenticateToken, async (req, res) => {
  const userId = req.user.id;
  try {
    const results = await sql`SELECT DISTINCT EXTRACT(YEAR FROM created_at) AS year FROM excel_sheets WHERE user_id = ${userId} ORDER BY year DESC`;
    const years = results.map((row) => row.year);
    res.json({ years });
  } catch (error) {
    console.error("Error fetching years:", error);
    res.status(500).json({ error: "Error fetching years from database." });
  }
});

// Fetch files for a specific year
app.get("/api/getFiles/:year", authenticateToken, async (req, res) => {
  const { year } = req.params;
  const userId = req.user.id;
  try {
    const results = await sql`
      SELECT excel_id, created_at 
      FROM excel_sheets 
      WHERE EXTRACT(YEAR FROM created_at) = ${year} AND user_id = ${userId}
      ORDER BY created_at DESC
    `;
    if (results.length === 0) {
      return res.status(404).json({ error: `No records found for year ${year}.` });
    }
    const files = results.map((row) => ({
      excel_id: row.excel_id,
      created_at: row.created_at,
    }));
    res.json({ files });
  } catch (error) {
    console.error("Error fetching files:", error);
    res.status(500).json({ error: "Error fetching files from database." });
  }
});

app.get("/api/getPBUData/:excelId", authenticateToken, async (req, res) => {
  const { excelId } = req.params;
  const userId = req.user.id;
  try {
    // Fetch highlightRows from excel_sheets
    const sheetResult = await sql`
      SELECT highlight_rows 
      FROM excel_sheets 
      WHERE excel_id = ${excelId} AND user_id = ${userId}
    `;
    if (sheetResult.length === 0) {
      return res.status(404).json({
        error: `No data found for Excel ID ${excelId} or unauthorized access.`,
      });
    }
    const highlightRows = sheetResult[0].highlight_rows || [];

    // Fetch cell data from excel_data
    const dataResults = await sql`
      SELECT field_key, field_value 
      FROM excel_data 
      WHERE excel_id = ${excelId}
    `;
    const jsonData = dataResults.map((row) => ({
      field_key: row.field_key,
      field_value: row.field_value,
    }));

    // Return both data and highlightRows
    res.json({ data: jsonData, highlightRows });
  } catch (error) {
    console.error("Error fetching PBU data:", error);
    res.status(500).json({ error: "Error fetching Excel data from database." });
  }
});

// Save data to database
app.post("/api/saveFiles", authenticateToken, async (req, res) => {
  const userId = req.user.id;
  const { data, highlightRows } = req.body; // Expecting { data: [...], highlightRows: [...] }
  if (!data || data.length === 0) {
    return res.status(400).json({ error: "Invalid or empty data provided." });
  }
  try {
    // Insert into excel_sheets with highlightRows
    const sheetResult = await sql`
      INSERT INTO excel_sheets (user_id, highlight_rows) 
      VALUES (${userId}, ${highlightRows ? JSON.stringify(highlightRows) : '[]'}) 
      RETURNING excel_id
    `;
    const excelId = sheetResult[0].excel_id;

    // Insert cell data into excel_data
    const dataToInsert = data.map((row) => ({
      excel_id: excelId,
      field_key: row.field_key,
      field_value: row.field_value,
    }));
    await sql`INSERT INTO excel_data ${sql(dataToInsert)}`;

    res.json({ success: true, message: "Data saved successfully", excelId });
  } catch (error) {
    console.error("Error saving data:", error);
    res.status(500).json({ error: "Error saving data to the database." });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});

module.exports = app;