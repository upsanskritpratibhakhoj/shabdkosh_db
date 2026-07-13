const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const filename = 'हिन्दी संस्कृत शब्दकोश.xlsx';
const filePath = path.join(__dirname, filename);

if (!fs.existsSync(filePath)) {
  console.error(`Error: File not found at ${filePath}`);
  process.exit(1);
}

console.log(`Reading workbook: ${filename}...`);
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert worksheet to raw array of arrays
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
console.log(`Loaded ${rows.length} rows.`);

if (rows.length === 0) {
  console.log('The sheet is empty.');
  process.exit(0);
}

// Print the first 5 rows to verify the structure
console.log('\nFirst 5 rows of the sheet:');
for (let i = 0; i < Math.min(5, rows.length); i++) {
  console.log(`Row ${i + 1}:`, rows[i]);
}

const seenKeys = new Set();
const uniqueRows = [];
let duplicatesRemoved = 0;

for (let i = 0; i < rows.length; i++) {
  const row = rows[i];
  const key = row[0]; // Column A is index 0
  
  // Check if row is completely empty
  const isRowEmpty = row.every(val => val === undefined || val === null || val === '');
  if (isRowEmpty) {
    continue;
  }
  
  // Clean the key (trim whitespace)
  const cleanKey = String(key || '').trim();

  if (cleanKey === '') {
    // If the key is empty but the row contains other data, keep it
    uniqueRows.push(row);
  } else if (!seenKeys.has(cleanKey)) {
    seenKeys.add(cleanKey);
    uniqueRows.push(row);
  } else {
    duplicatesRemoved++;
  }
}

console.log(`\nUnique rows kept: ${uniqueRows.length}`);
console.log(`Duplicate rows removed: ${duplicatesRemoved}`);

// Recreate the workbook with deduplicated rows
const newWorksheet = xlsx.utils.aoa_to_sheet(uniqueRows);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

// Write to a temporary file
const outputFilename = 'हिन्दी संस्कृत शब्दकोश_cleaned.xlsx';
const outputPath = path.join(__dirname, outputFilename);

xlsx.writeFile(newWorkbook, outputPath);
console.log(`\nCleaned excel file saved to: ${outputPath}`);

// Overwrite the original file with the cleaned one
try {
  fs.copyFileSync(outputPath, filePath);
  console.log(`Successfully updated the original file: ${filename}`);
  // Optionally clean up the temp file
  fs.unlinkSync(outputPath);
  console.log('Cleaned up temporary file.');
} catch (err) {
  console.error(`Error updating original file: ${err.message}`);
}
