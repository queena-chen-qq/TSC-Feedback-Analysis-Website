// Test script to verify Excel parsing logic
// Usage: node scripts/test-parse.js <path-to-excel>
const XLSX = require('xlsx');
const path = require('path');

const file = process.argv[2];
if (!file) { console.log('Usage: node scripts/test-parse.js <excel-file>'); process.exit(1); }

const wb = XLSX.readFile(file);
console.log('Sheet names:', wb.SheetNames);
console.log('---');

wb.SheetNames.forEach((name, i) => {
  console.log(`\n=== Sheet ${i}: "${name}" ===`);
  const allRows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
  // Show first 5 rows
  allRows.slice(0, 5).forEach((row, ri) => {
    console.log(`  Row ${ri}:`, row.slice(0, 6).map(c => String(c ?? '').substring(0, 20)));
  });
  console.log(`  ... total ${allRows.length} rows`);
});
