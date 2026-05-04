import * as XLSX from 'xlsx';

const STORAGE_KEY = 'tsc_feedbacks';

const SCORE_MAP = { '非常滿意': 5, '滿意': 4, '普通': 3, '不滿意': 2, '非常不滿意': 1 };

export function readData() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
  } catch { return []; }
}

export function writeData(data) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

export function clearData() {
  localStorage.removeItem(STORAGE_KEY);
}

function classifyColumns(headers) {
  const ratingCols = [], suggestionCols = [], extraCols = [];
  const skip = ['姓名', '參加組別', '課程', '上課時段'];
  headers.forEach(h => {
    if (skip.includes(h)) return;
    if (h.includes('建議')) suggestionCols.push(h);
    else if (h.includes('課程安排') || h.includes('講師授課')) ratingCols.push(h);
    else if (h.includes('難易度') || h.includes('應用') || h.includes('最想學') || h.includes('請問')) extraCols.push(h);
  });
  return { ratingCols, suggestionCols, extraCols };
}

function extractDateFromFilename(filename) {
  const match = filename.match(/(\d{2})(\d{2})/);
  if (match) return `${match[1]}/${match[2]}`;
  return new Date().toISOString().slice(5, 10).replace('-', '/');
}

export function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        if (rows.length === 0) { reject('Excel 中沒有資料'); return; }

        const headers = Object.keys(rows[0]);
        const colInfo = classifyColumns(headers);
        const fileDate = extractDateFromFilename(file.name);

        const records = rows.map((row, i) => {
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

          return {
            id: Date.now() + i + Math.random(),
            date: fileDate,
            name: String(row['姓名'] ?? '未知').trim(),
            group: String(row['參加組別'] ?? '').trim(),
            course: String(row['課程'] || row['上課時段'] || '').trim(),
            ratings, suggestions, extras
          };
        });

        const existing = readData();
        writeData([...existing, ...records]);
        resolve({ count: records.length, date: fileDate });
      } catch (err) { reject(err.message); }
    };
    reader.onerror = () => reject('檔案讀取失敗');
    reader.readAsArrayBuffer(file);
  });
}

export function getDates() {
  const data = readData();
  return [...new Set(data.map(item => item.date))].sort();
}

export function getFeedbacks(date, group) {
  let data = readData();
  if (date) data = data.filter(item => item.date === date);
  if (group) data = data.filter(item => item.group === group);
  return data;
}

export function getStats(date, group) {
  const filtered = getFeedbacks(date, group);
  if (filtered.length === 0) {
    return { total: 0, groups: {}, ratingAvg: {}, distribution: {}, overallAvg: 0, suggestions: [], extras: [] };
  }

  const total = filtered.length;
  const groups = {};
  filtered.forEach(f => { if (f.group) groups[f.group] = (groups[f.group] || 0) + 1; });

  const allRatingCols = new Set();
  filtered.forEach(f => Object.keys(f.ratings || {}).forEach(k => allRatingCols.add(k)));

  const ratingAvg = {};
  allRatingCols.forEach(col => {
    const scores = filtered.map(f => f.ratings?.[col]?.score).filter(s => s && s > 0);
    ratingAvg[col] = scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
  });

  const distribution = { '非常滿意': 0, '滿意': 0, '普通': 0, '不滿意': 0, '非常不滿意': 0 };
  filtered.forEach(f => {
    Object.values(f.ratings || {}).forEach(r => {
      if (distribution[r.label] !== undefined) distribution[r.label]++;
    });
  });

  const allScores = filtered.flatMap(f => Object.values(f.ratings || {}).map(r => r.score).filter(s => s > 0));
  const overallAvg = allScores.length ? +(allScores.reduce((a, b) => a + b, 0) / allScores.length).toFixed(2) : 0;

  const suggestions = [];
  filtered.forEach(f => {
    Object.entries(f.suggestions || {}).forEach(([col, content]) => {
      if (content) {
        const type = col.replace(/對於/g, '').replace(/的建議[：:]?/g, '').replace(/：/g, '').trim();
        suggestions.push({ name: f.name, type: type || col, content });
      }
    });
  });

  const extras = [];
  filtered.forEach(f => {
    Object.entries(f.extras || {}).forEach(([col, val]) => {
      if (val) extras.push({ name: f.name, field: col, value: val });
    });
  });

  return { total, groups, ratingAvg, distribution, overallAvg, suggestions, extras };
}
