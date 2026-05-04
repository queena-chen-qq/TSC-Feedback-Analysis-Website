const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());

const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'feedbacks.json');
const upload = multer({ dest: path.join(__dirname, 'uploads') });

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, '[]', 'utf-8');

function readData() { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8')); }
function writeData(data) { fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8'); }

const SCORE_MAP = { '非常滿意': 5, '滿意': 4, '普通': 3, '不滿意': 2, '非常不滿意': 1 };

// Auto-detect column types from header names
function classifyColumns(headers) {
  const ratingCols = [];    // satisfaction-scale columns
  const suggestionCols = []; // free-text suggestion columns
  const extraCols = [];      // other interesting columns (difficulty, application, etc.)
  const skipCols = ['姓名', '參加組別', '課程', '上課時段'];

  headers.forEach(h => {
    if (skipCols.includes(h)) return;

    // Suggestion columns: contain 建議
    if (h.includes('建議')) {
      suggestionCols.push(h);
    }
    // Rating columns: contain 課程安排 or 講師授課 (satisfaction scale)
    else if (h.includes('課程安排') || h.includes('講師授課')) {
      ratingCols.push(h);
    }
    // Extra columns: difficulty, application scenarios, future learning, etc.
    else if (h.includes('難易度') || h.includes('應用') || h.includes('最想學') || h.includes('請問')) {
      extraCols.push(h);
    }
  });

  return { ratingCols, suggestionCols, extraCols };
}

function parseRow(row, fileDate, colInfo) {
  const ratings = {};
  colInfo.ratingCols.forEach(col => {
    const val = String(row[col] ?? '').trim();
    ratings[col] = { label: val || '未填', score: SCORE_MAP[val] || 0 };
  });

  const suggestions = {};
  colInfo.suggestionCols.forEach(col => {
    const val = String(row[col] ?? '').trim();
    suggestions[col] = (!val || val === '無' || val === 'Na' || val === 'NA') ? '' : val;
  });

  const extras = {};
  colInfo.extraCols.forEach(col => {
    const val = String(row[col] ?? '').trim();
    if (val && val !== '無' && val !== 'Na' && val !== 'NA') extras[col] = val;
  });

  // Detect course name from either '課程' or '上課時段'
  const course = row['課程'] || row['上課時段'] || '';

  return {
    id: Date.now() + Math.random(),
    date: fileDate,
    name: String(row['姓名'] ?? '未知').trim(),
    group: String(row['參加組別'] ?? '').trim(),
    course: String(course).trim(),
    ratings,
    suggestions,
    extras
  };
}

function extractDateFromFilename(filename) {
  const match = filename.match(/(\d{2})(\d{2})/);
  if (match) return `${match[1]}/${match[2]}`;
  return new Date().toISOString().slice(5, 10).replace('-', '/');
}

// Upload Excel
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: '請上傳檔案' });

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    if (rows.length === 0) return res.status(400).json({ error: 'Excel 中沒有資料' });

    const headers = Object.keys(rows[0]);
    const colInfo = classifyColumns(headers);
    const fileDate = extractDateFromFilename(req.file.originalname);

    const existing = readData();
    const newRecords = rows.map(row => parseRow(row, fileDate, colInfo));
    writeData([...existing, ...newRecords]);

    fs.unlinkSync(req.file.path);
    res.json({ message: `成功匯入 ${newRecords.length} 筆回饋 (${fileDate})`, count: newRecords.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '檔案處理失敗: ' + err.message });
  }
});

// Batch import from JSON (for direct data injection)
app.post('/api/import', (req, res) => {
  try {
    const { records } = req.body;
    if (!records || !Array.isArray(records)) return res.status(400).json({ error: '無效資料' });
    const existing = readData();
    writeData([...existing, ...records]);
    res.json({ message: `成功匯入 ${records.length} 筆`, count: records.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get available dates
app.get('/api/dates', (req, res) => {
  const data = readData();
  const dates = [...new Set(data.map(item => item.date))].sort();
  res.json(dates);
});

// Get feedbacks by date and group
app.get('/api/feedbacks', (req, res) => {
  const { date, group } = req.query;
  let data = readData();
  if (date) data = data.filter(item => item.date === date);
  if (group) data = data.filter(item => item.group === group);
  res.json(data);
});

// Get stats for a date and group
app.get('/api/stats', (req, res) => {
  const { date, group } = req.query;
  let data = readData();
  if (date) data = data.filter(item => item.date === date);
  if (group) data = data.filter(item => item.group === group);
  const filtered = data;

  if (filtered.length === 0) {
    return res.json({ total: 0, groups: {}, ratingAvg: {}, distribution: {}, overallAvg: 0, suggestions: [], extras: [] });
  }

  const total = filtered.length;

  // Group count
  const groups = {};
  filtered.forEach(f => { if (f.group) groups[f.group] = (groups[f.group] || 0) + 1; });

  // Collect all unique rating column names across all records
  const allRatingCols = new Set();
  filtered.forEach(f => Object.keys(f.ratings || {}).forEach(k => allRatingCols.add(k)));

  // Average scores per rating column
  const ratingAvg = {};
  allRatingCols.forEach(col => {
    const scores = filtered.map(f => f.ratings?.[col]?.score).filter(s => s && s > 0);
    ratingAvg[col] = scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
  });

  // Overall satisfaction distribution
  const distribution = { '非常滿意': 0, '滿意': 0, '普通': 0, '不滿意': 0, '非常不滿意': 0 };
  filtered.forEach(f => {
    Object.values(f.ratings || {}).forEach(r => {
      if (distribution[r.label] !== undefined) distribution[r.label]++;
    });
  });

  // Overall average
  const allScores = filtered.flatMap(f => Object.values(f.ratings || {}).map(r => r.score).filter(s => s > 0));
  const overallAvg = allScores.length ? +(allScores.reduce((a, b) => a + b, 0) / allScores.length).toFixed(2) : 0;

  // Suggestions
  const suggestions = [];
  filtered.forEach(f => {
    Object.entries(f.suggestions || {}).forEach(([col, content]) => {
      if (content) {
        const type = col.replace(/對於/g, '').replace(/的建議[：:]?/g, '').replace(/：/g, '').trim();
        suggestions.push({ name: f.name, type: type || col, content });
      }
    });
  });

  // Extras (difficulty, application, etc.)
  const extras = [];
  filtered.forEach(f => {
    Object.entries(f.extras || {}).forEach(([col, val]) => {
      if (val) extras.push({ name: f.name, field: col, value: val });
    });
  });

  res.json({ total, groups, ratingAvg, distribution, overallAvg, suggestions, extras });
});

// Delete all
app.delete('/api/feedbacks', (req, res) => {
  writeData([]);
  res.json({ message: '所有資料已清除' });
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
