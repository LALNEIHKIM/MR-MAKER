const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const storeDir = path.join(__dirname, 'store-data');
const usersFile = path.join(storeDir, 'users.json');
const submissionsFile = path.join(storeDir, 'submissions.json');
const requestsFile = path.join(storeDir, 'edit-requests.json');
fs.mkdirSync(storeDir, { recursive: true });
function readStore(file, fallback = []) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return fallback; }
}
function writeStore(file, value) { fs.writeFileSync(file, JSON.stringify(value, null, 2)); }
function isAdmin(req) { return req.get('x-cpl-admin') === 'admin'; }

app.post('/api/register', (req, res) => {
  const name = String(req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Name is required' });
  const users = readStore(usersFile);
  if (users.some(u => u.name.toLowerCase() === name.toLowerCase())) return res.status(409).json({ error: 'Name already submitted. Please contact admin for approval.' });
  users.push({ name, approved: false, createdAt: new Date().toISOString() });
  writeStore(usersFile, users);
  res.json({ message: 'Please contact admin for approval.' });
});
app.get('/api/users', (req, res) => isAdmin(req) ? res.json(readStore(usersFile)) : res.status(403).json({ error: 'Admin only' }));
app.post('/api/users/:name/approve', (req, res) => {
  if (!isAdmin(req)) return res.status(403).json({ error: 'Admin only' });
  const users = readStore(usersFile); const user = users.find(u => u.name.toLowerCase() === decodeURIComponent(req.params.name).toLowerCase());
  if (!user) return res.status(404).json({ error: 'User not found' }); user.approved = true; writeStore(usersFile, users); res.json({ ok: true });
});
app.post('/api/login', (req, res) => {
  const name = String(req.body.name || '').trim();
  const user = readStore(usersFile).find(u => u.name.toLowerCase() === name.toLowerCase() && u.approved);
  if (!user) return res.status(403).json({ error: 'Please contact admin for approval.' });
  res.json({ ok: true, name: user.name });
});
app.post('/api/submit', (req, res) => {
  const name = String(req.body.name || '').trim();
  const user = readStore(usersFile).find(u => u.name.toLowerCase() === name.toLowerCase() && u.approved);
  if (!user && !isAdmin(req)) return res.status(403).json({ error: 'Please contact admin for approval.' });
  const { period, data } = req.body;
  if (!period || !data || !Array.isArray(data.mrSheets)) return res.status(400).json({ error: 'Invalid store data' });
  const submissions = readStore(submissionsFile); const index = submissions.findIndex(s => s.period === period && String(s.owner || '').toLowerCase() === name.toLowerCase());
  let mergedSheets = data.mrSheets;
  if (index >= 0) {
    const previousSheets = submissions[index].data?.mrSheets || [];
    const byLabel = new Map(previousSheets.map(sheet => [sheet.label, sheet]));
    for (const sheet of data.mrSheets) byLabel.set(sheet.label, sheet);
    mergedSheets = Array.from(byLabel.values());
  }
  const record = { period, owner: name || 'ADMIN', data: { ...data, mrSheets: mergedSheets }, submittedAt: new Date().toISOString() };
  if (index >= 0) submissions[index] = record; else submissions.push(record);
  writeStore(submissionsFile, submissions); res.json({ ok: true });
});
app.get('/api/store-data/:period', (req, res) => {
  const period = decodeURIComponent(req.params.period);
  const records = readStore(submissionsFile).filter(s => s.period === period);
  res.json({ period, data: records.flatMap(s => s.data.mrSheets || []), found: records.length > 0 });
});
app.post('/api/edit-requests', (req, res) => {
  const name = String(req.body.name || '').trim(); const { period, mrLabel, message } = req.body;
  if (!name || !message) return res.status(400).json({ error: 'Name and message are required' });
  const requests = readStore(requestsFile); requests.push({ id: Date.now().toString(), name, period, mrLabel, message, status: 'pending', createdAt: new Date().toISOString() });
  writeStore(requestsFile, requests); res.json({ ok: true });
});
app.get('/api/edit-requests', (req, res) => isAdmin(req) ? res.json(readStore(requestsFile)) : res.status(403).json({ error: 'Admin only' }));
app.post('/api/edit-requests/:id/approve', (req, res) => {
  if (!isAdmin(req)) return res.status(403).json({ error: 'Admin only' });
  const requests = readStore(requestsFile); const item = requests.find(r => r.id === req.params.id);
  if (!item) return res.status(404).json({ error: 'Request not found' }); item.status = 'approved'; item.approvedAt = new Date().toISOString(); writeStore(requestsFile, requests); res.json({ ok: true });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Excel file se master data fetch karne ka API endpoint
app.get('/api/master-data', (req, res) => {
  try {
    const excelFilePath = path.join(__dirname, 'SEP 2026.xlsx');
    
    if (!fs.existsSync(excelFilePath)) {
      return res.json({ headers: [], data: [] });
    }

    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0]; 
    const worksheet = workbook.Sheets[sheetName];
    const jsonRows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (jsonRows.length === 0) {
      return res.json({ headers: [], data: [] });
    }

    const headers = jsonRows[0] || [];
    const dataRows = [];

    // Rows ko loop karke data extract karna
    for (let i = 1; i < jsonRows.length; i++) {
      let rVals = jsonRows[i];
      if (rVals && rVals.length > 0) {
        let code = String(rVals[1] || '').trim(); // Column B
        let name = String(rVals[3] || '').trim(); // Column D
        
        if (code || name) {
          let designation = String(rVals[2] || 'MAZ').trim();
          let fatherName = String(rVals[5] || '').trim(); // Column F -> Index 5
          let gender = String(rVals[6] || 'MALE').trim();      // Column G -> Index 6
          let accountNo = String(rVals[15] || '').trim(); // Column P -> Index 15
          let ifsCode = String(rVals[16] || '').trim();     // Column Q -> Index 16
          let enrolmentDate = String(rVals[22] || '').trim(); // Column W -> Index 22
          
          // Agar master excel me kisi aur column par discharge date hai toh uska index yahan dein, 
          // filhal ise blank ya master data ke anusaar set kar rahe hain taaki 21-Aug-2026 automatically na aaye.
          let dischargeDate = String(rVals[23] || '').trim(); 

          let rowArr = new Array(44).fill(''); // ✅ Total 44 columns (0 to 43)
          rowArr[1] = code;
          rowArr[2] = designation;
          rowArr[3] = name;
          rowArr[4] = fatherName;    // Table Column E
          rowArr[5] = gender ? gender : 'MALE'; // Table Column F
          rowArr[15] = accountNo; // Column P (Index 15) - Bank A/C No
          rowArr[16] = ifsCode;   // Column Q (Index 16) - IFS Code
          rowArr[17] = "31"; // Default 31 days
          rowArr[22] = enrolmentDate;
          rowArr[43] = dischargeDate; // ✅ Date of Discharge ab apne sahi index (43) par hai

          dataRows.push({
            code: code,
            name: name,
            designation: designation,
            fatherName: fatherName,
            gender: gender,
            accountNo: accountNo,
            ifsCode: ifsCode,
            enrolmentDate: enrolmentDate,
            dischargeDate: dischargeDate,
            rowData: rowArr
          });
        }
      }
    }

    res.json({ headers: headers, data: dataRows });
  } catch (error) {
    console.error("Excel read error:", error);
    res.status(500).json({ error: "Failed to read excel file" });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
});
