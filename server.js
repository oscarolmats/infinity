const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const app = express();
const upload = multer({ dest: 'uploads/' });
app.use(express.json());

// Simple config via env
const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-in-prod';
const USERS_FILE = path.join(__dirname, 'users.json');

function readUsers(){
  try {
    if(!fs.existsSync(USERS_FILE)) return [];
    const raw = fs.readFileSync(USERS_FILE, 'utf8');
    return JSON.parse(raw || '[]');
  } catch { return []; }
}
function writeUsers(users){
  fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
}

function authMiddleware(req, res, next){
  const auth = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : null;
  if(!token) return res.status(401).json({ error: 'Unauthorized' });
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.user = payload;
    next();
  } catch (e){
    return res.status(401).json({ error: 'Invalid token' });
  }
}

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

// Auth endpoints
app.post('/api/register', (req, res) => {
  const { email, password } = req.body || {};
  if(!email || !password) return res.status(400).json({ error: 'Email and password required' });
  const users = readUsers();
  if(users.find(u => u.email.toLowerCase() === String(email).toLowerCase())){
    return res.status(409).json({ error: 'User exists' });
  }
  const hash = bcrypt.hashSync(password, 10);
  const user = { id: Date.now().toString(), email, passwordHash: hash, createdAt: new Date().toISOString() };
  users.push(user);
  writeUsers(users);
  const token = jwt.sign({ sub: user.id, email: user.email }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, email: user.email } });
});

app.post('/api/login', (req, res) => {
  const { email, password } = req.body || {};
  if(!email || !password) return res.status(400).json({ error: 'Email and password required' });
  const users = readUsers();
  const user = users.find(u => u.email.toLowerCase() === String(email).toLowerCase());
  if(!user) return res.status(401).json({ error: 'Invalid credentials' });
  const ok = bcrypt.compareSync(password, user.passwordHash);
  if(!ok) return res.status(401).json({ error: 'Invalid credentials' });
  const token = jwt.sign({ sub: user.id, email: user.email }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, email: user.email } });
});

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
    fs.unlinkSync(req.file.path);

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

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server körs på http://localhost:${PORT}`);
  console.log('Öppna index.html i din webbläsare');
});







