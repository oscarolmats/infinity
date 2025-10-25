const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();

// Konfigurera multer för Vercel
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  
  if (req.method === 'OPTIONS') {
    res.sendStatus(200);
  } else {
    next();
  }
});

app.use(express.json());

// Root endpoint - servera index.html
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
        <title>Tabellen</title>
        <meta charset="utf-8">
    </head>
    <body>
        <h1>Tabellen App</h1>
        <p>Server is running on Vercel!</p>
        <p><a href="/api/health">Test API Health</a></p>
    </body>
    </html>
  `);
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: 'Server is running on Vercel',
    timestamp: new Date().toISOString()
  });
});

// Upload endpoint
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Ingen fil uppladdad');
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return res.status(400).send('Ingen worksheet hittades');
    }

    let html = '<table><thead><tr>';
    const headers = [];
    
    // Läs header-raden
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      const value = cell.value || '';
      headers.push(value);
      html += `<th>${value}</th>`;
    });
    html += '</tr></thead><tbody>';

    // Läs alla datarader
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      html += '<tr>';
      row.eachCell({ includeEmpty: true }, (cell) => {
        const value = cell.value || '';
        html += `<td>${value}</td>`;
      });
      html += '</tr>';
    }
    html += '</tbody></table>';

    res.send(html);
  } catch (error) {
    console.error('Fel vid läsning av Excel:', error);
    res.status(500).send('Kunde inte läsa Excel-filen: ' + error.message);
  }
});

// Export för Vercel
module.exports = app;