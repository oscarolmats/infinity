const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();

// Konfigurera multer för Vercel (använd memory storage istället för disk)
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// CORS för frontend
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

// Serve static files
app.use(express.static(path.join(__dirname, '../')));

// Hantera Excel-uppladdningar
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Ingen fil uppladdad');
    }

    const workbook = new ExcelJS.Workbook();
    
    // Läs från buffer istället för fil
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
    res.status(500).send('Kunde inte läsa Excel-filen');
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'Server is running' });
});

// Root endpoint för att servera index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../index.html'));
});

// Export app för Vercel (VIKTIGT: använd inte app.listen())
module.exports = app;
