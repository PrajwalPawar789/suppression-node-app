const express = require('express');
const multer = require('multer');
const fs = require('fs');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const app = express();

// Helper function to normalize strings (only trim whitespace)
const normalizeString = (str) => {
  if (str === undefined || str === null) {
    console.log('Attempting to normalize an undefined or null value.');
    return '';
  }
  return str.trim();
};

// Multer configuration for file storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  }
});
const upload = multer({ storage: storage });

// PostgreSQL connection settings
const pool = new Pool({
  user: 'postgres',
  host: 'localhost',
  database: 'supppression-db',
  password: 'root',
  port: 5432
});

// Function to check the database for a match based on left_3 and left_4
async function checkDatabase(left3, left4, clientCode, dateFilter) {
  const client = await pool.connect();
  try {
    // console.log(`Checking for left_3: ${left3}, left_4: ${left4}, clientCode: ${clientCode}`);
    const query = `
      SELECT date_, EXISTS (
        SELECT 1
        FROM campaigns
        WHERE 
          left_3 = $1 AND
          left_4 = $2 AND
          client = $3
      ) AS match_found
      FROM campaigns
      WHERE 
          left_3 = $1 AND
          left_4 = $2 AND
          client = $3
      LIMIT 1;`;
    const result = await client.query(query, [left3, left4, clientCode]);
    const row = result.rows[0];
    if (row) {
      const dateFromDb = new Date(row.date_);
      const currentDate = new Date();
      const monthsAgoDate = new Date(currentDate.setMonth(currentDate.getMonth() - dateFilter));

      return {
        exists: row.match_found,
        dateStatus: dateFromDb < monthsAgoDate ? 'Suppression Cleared' : 'Still Suppressed'
      };
    }
    return { exists: false, dateStatus: 'Fresh Lead GTG' }; // No record matched
  } finally {
    client.release();
  }
}

// Read the Excel file, calculate left_3 and left_4, check the database, and add status
async function processFile(filePath, clientCode, dateFilter) { // Include dateFilter as a parameter
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  // Columns as per Excel headers
  let companyIndex, firstNameIndex, lastNameIndex, emailIndex, phoneIndex;
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    switch(cell.value) {
      case 'Company Name':
        companyIndex = colNumber;
        break;
      case 'First Name':
        firstNameIndex = colNumber;
        break;
      case 'Last Name':
        lastNameIndex = colNumber;
        break;
      case 'Email ID':
        emailIndex = colNumber;
        break;
      case 'Phone Number':
        phoneIndex = colNumber;
        break;
    }
  });

  const statusColumn = worksheet.getColumn(worksheet.columnCount + 1);
  statusColumn.header = 'Match Status';

  const clientCodeStatusColumn = worksheet.getColumn(worksheet.columnCount + 2);
  clientCodeStatusColumn.header = 'Client Code Status';

  const dateStatusColumn = worksheet.getColumn(worksheet.columnCount + 3);
  dateStatusColumn.header = 'Date Status';

  for (let i = 2; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i);
    const firstName = normalizeString(row.getCell(firstNameIndex).value);
    const lastName = normalizeString(row.getCell(lastNameIndex).value);
    const companyName = normalizeString(row.getCell(companyIndex).value);

    const left3 = `${firstName.substring(0, 3)}${lastName.substring(0, 3)}${companyName.substring(0, 3)}`;
    const left4 = `${firstName.substring(0, 4)}${lastName.substring(0, 4)}${companyName.substring(0, 4)}`;

    const dbResult = await checkDatabase(left3, left4, clientCode, dateFilter);
    row.getCell(statusColumn.number).value = dbResult.exists ? 'Match' : 'Unmatch';
    row.getCell(clientCodeStatusColumn.number).value = dbResult.exists ? 'Match' : 'Unmatch';
    row.getCell(dateStatusColumn.number).value = dbResult.dateStatus;

    row.commit();
  }

  const newFilePath = 'Updated-' + Date.now() + '.xlsx';
  await workbook.xlsx.writeFile(newFilePath);
  return newFilePath;
}

app.set('view engine', 'ejs');

app.get('/', (req, res) => {
  res.render('upload');
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }
  const clientCode = req.body.clientCode; // Capture the client code from the form
  const dateFilter = parseInt(req.body.dateFilter); // Capture the date filter from the form
  try {
    const newFilePath = await processFile(req.file.path, clientCode, dateFilter);
    res.download(newFilePath, (err) => {
      if (err) throw err;
      fs.unlinkSync(newFilePath);
      fs.unlinkSync(req.file.path);
    });
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).send('Error processing the file.');
  }
});

const port = 3000;
app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});