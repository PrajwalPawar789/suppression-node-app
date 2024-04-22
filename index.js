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
  database: 'suppression-db',
  password: 'root',
  port: 5432
});

// Function to check the database for a match based on left_3 and left_4
async function checkDatabase(left3, left4) {
  const client = await pool.connect();
  try {
    console.log(`Checking for left_3: ${left3}, left_4: ${left4}`);
    const result = await client.query(
      `SELECT EXISTS (
        SELECT 1
        FROM campaigns
        WHERE 
          left_3 = $1 AND
          left_4 = $2
      )`, [left3, left4]
    );
    return result.rows[0].exists;
  } finally {
    client.release();
  }
}

// Read the Excel file, calculate left_3 and left_4, check the database, and add status
async function processFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  const headerRow = worksheet.getRow(1);

  let companyIndex, firstNameIndex, lastNameIndex, emailIndex, phoneIndex;
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    if (cell.value === 'Company Name') companyIndex = colNumber;
    else if (cell.value === 'First Name') firstNameIndex = colNumber;
    else if (cell.value === 'Last Name') lastNameIndex = colNumber;
    else if (cell.value === 'Email ID') emailIndex = colNumber;
    else if (cell.value === 'Phone Number') phoneIndex = colNumber;
  });

  console.log({companyIndex, firstNameIndex, lastNameIndex, emailIndex, phoneIndex});

  if (!companyIndex || !firstNameIndex || !lastNameIndex || !emailIndex || !phoneIndex) {
    throw new Error('One or more required headers not found in the Excel sheet.');
  }

  const statusColumn = worksheet.getColumn(worksheet.columnCount + 1);
  statusColumn.header = 'Status';

  for (let i = 2; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i);
    const firstName = row.getCell(firstNameIndex).value;
    const lastName = row.getCell(lastNameIndex).value;
    const companyName = row.getCell(companyIndex).value;

    // Skip processing for rows where critical cells are empty
    if (!firstName && !lastName && !companyName) {
      console.log(`Skipping empty row ${i}.`);
      continue;
    }

    const normalizedFirstName = normalizeString(firstName);
    const normalizedLastName = normalizeString(lastName);
    const normalizedCompanyName = normalizeString(companyName);

    const left3 = `${normalizedFirstName.substring(0, 3)}${normalizedLastName.substring(0, 3)}${normalizedCompanyName.substring(0, 3)}`;
    const left4 = `${normalizedFirstName.substring(0, 4)}${normalizedLastName.substring(0, 4)}${normalizedCompanyName.substring(0, 4)}`;

    console.log(`Row ${i} - Checking for left_3: ${left3}, left_4: ${left4}`);

    const status = await checkDatabase(left3, left4) ? 'Match' : 'Unmatch';
    row.getCell(worksheet.columnCount).value = status;
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
  try {
    const newFilePath = await processFile(req.file.path);
    res.download(newFilePath, (err) => {
      if (err) throw err;
      fs.unlinkSync(newFilePath); // Optionally delete the file after sending
      fs.unlinkSync(req.file.path); // Delete the uploaded file as well
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
