const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

// CORS för lokal utveckling
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

// Servera statiska filer från root
app.use(express.static('.'));

// Hantera Excel-uppladdningar
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Ingen fil uppladdad');
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);
    
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

    // Ta bort den tillfälliga filen
    const fs = require('fs');
    fs.unlinkSync(req.file.path);

    res.send(html);
  } catch (error) {
    console.error('Fel vid läsning av Excel:', error);
    res.status(500).send('Kunde inte läsa Excel-filen');
  }
});

// Proxy endpoint for co2data.fi API (to avoid CORS)
app.get('/api/co2data', async (req, res) => {
  try {
    const response = await fetch('https://co2data.fi/api/co2data_construction.json');
    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error('co2data.fi API Error:', error);
    res.status(500).json({ error: 'Failed to fetch co2data.fi data', message: error.message });
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'Server is running' });
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server körs på http://localhost:${PORT}`);
  console.log('Öppna index.html i din webbläsare');
});







