const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const excelFilename = 'हिन्दी संस्कृत शब्दकोश.xlsx';
const excelFilePath = path.join(__dirname, excelFilename);
const jsonFilename = 'dictionary.json';
const jsonFilePath = path.join(__dirname, jsonFilename);

if (!fs.existsSync(excelFilePath)) {
  console.error(`Error: Excel file not found at ${excelFilePath}`);
  process.exit(1);
}

console.log(`Reading Excel file: ${excelFilename}...`);
const workbook = xlsx.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert to array of arrays
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
console.log(`Loaded ${rows.length} rows.`);

const dictionary = [];
const seenKeys = new Set();
let duplicatesCount = 0;

// Skip the first row if it contains headers
let startIndex = 0;
if (rows.length > 0) {
  const firstRow = rows[0];
  const colA = String(firstRow[0] || '').trim();
  if (colA.startsWith('हिन्दी') || colA.includes('शब्द') || colA.toLowerCase() === 'hindi') {
    console.log('Skipping header row:', firstRow);
    startIndex = 1;
  }
}

for (let i = startIndex; i < rows.length; i++) {
  const row = rows[i];
  const hindi = String(row[0] || '').trim();
  const sanskrit = String(row[1] || '').trim();

  // Skip completely empty lines
  if (!hindi && !sanskrit) {
    continue;
  }

  // Deduplicate based on Column A (hindi)
  if (!seenKeys.has(hindi)) {
    seenKeys.add(hindi);
    dictionary.push({
      hindi: hindi,
      sanskrit: sanskrit
    });
  } else {
    duplicatesCount++;
  }
}

console.log(`Removed ${duplicatesCount} duplicate keys.`);
console.log(`Writing ${dictionary.length} unique entries to ${jsonFilename}...`);
fs.writeFileSync(jsonFilePath, JSON.stringify(dictionary, null, 2), 'utf-8');
console.log('Conversion completed successfully!');
