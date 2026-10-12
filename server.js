const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

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

          let rowArr = new Array(24).fill('');
rowArr[1] = code;
rowArr[2] = designation;
rowArr[3] = name;
rowArr[4] = fatherName;    // Table Column E
rowArr[5] = gender ? gender : 'MALE'; // Table Column F
rowArr[15] = accountNo; // ✅ Column P (Index 15) - Bank A/C No
rowArr[16] = ifsCode;   // ✅ Column Q (Index 16) - IFS Code
rowArr[17] = "31"; // Default 31 days
rowArr[22] = enrolmentDate;
rowArr[23] = "21-Aug-2026";

          dataRows.push({
            code: code,
            name: name,
            designation: designation,
            fatherName: fatherName,
            gender: gender,
            accountNo: accountNo,
            ifsCode: ifsCode,
            enrolmentDate: enrolmentDate,
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