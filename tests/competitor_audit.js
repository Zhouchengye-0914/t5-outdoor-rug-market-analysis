'use strict';

// Verifies the seven competitor Excel snapshots, raw_* tables, and canonical dedup_* tables.

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const RAW_DIR = path.resolve(ROOT, 'data/raw');
const DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const EXPECTED_DEDUP = { '202601': 64, '202602': 74, '202603': 82, '202604': 79, '202605': 88, '202606': 91, '202607': 94 };

if (!fs.existsSync(DB_PATH)) {
  console.error('Missing competitor DB: ' + DB_PATH);
  process.exit(1);
}

const files = fs.readdirSync(RAW_DIR)
  .filter((name) => /^Competitor-US-2026\.\d{2}-\d+\.xlsx$/i.test(name))
  .sort();
let checks = 0;
let failures = 0;
let cellsChecked = 0;
let mismatchSamples = 0;

function check(label, ok, detail) {
  checks++;
  if (!ok) failures++;
  console.log((ok ? '[PASS] ' : '[FAIL] ') + label + (detail ? ' (' + detail + ')' : ''));
}

function colNorm(name) {
  if (!name) return '';
  let value = String(name).trim();
  value = value.replace(/[\$\(\)\'`]/g, '_').replace(/&/g, '_and_').replace(/\s+/g, '_');
  value = value.replace(/[^A-Za-z0-9_\u4e00-\u9fff]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '');
  if (/^\d/.test(value)) value = '_' + value;
  return value || 'col';
}

function dedupColumns(columns) {
  const seen = {};
  return columns.map((column, index) => {
    seen[column] = (seen[column] || 0) + 1;
    const deduped = seen[column] === 1 ? column : column + '_' + seen[column];
    return deduped || '_col_' + index;
  });
}

function equivalent(expected, actual) {
  if (expected === null || expected === undefined || expected === '') return actual === null;
  if (typeof expected === 'number') {
    return actual !== null && Number.isFinite(Number(actual)) && Math.abs(Number(actual) - expected) <= 1e-9;
  }
  if (typeof actual === 'number' && /^-?\d+(?:\.\d+)?$/.test(String(expected).trim())) {
    return Math.abs(actual - Number(expected)) <= 1e-9;
  }
  return String(actual) === String(expected);
}

function listingKey(row) {
  return String(row['父ASIN'] || '').trim() || String(row.ASIN || '').trim();
}

const db = new DatabaseSync(DB_PATH, { readOnly: true });
check('seven competitor snapshots discovered', files.length === 7, 'files=' + files.length);

for (const file of files) {
  const match = file.match(/2026\.(\d{2})/);
  const month = '2026' + match[1];
  const workbook = xlsx.readFile(path.join(RAW_DIR, file), { cellDates: false, cellNF: false, cellFormula: false });
  const dataSheetName = workbook.SheetNames.find((name) => !/^Notes$/i.test(name));
  const sheet = workbook.Sheets[dataSheetName];
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const excelHeaders = [];
  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: range.s.r, c: column })];
    excelHeaders.push(colNorm(cell ? cell.v : ''));
  }
  const normalizedHeaders = dedupColumns(excelHeaders);
  const rawTable = 'raw_' + month;
  const dedupTable = 'dedup_' + month;
  const rawColumns = db.prepare('PRAGMA table_info(' + rawTable + ')').all()
    .map((column) => column.name).filter((name) => name !== 'row_id');
  const rawRows = db.prepare('SELECT * FROM ' + rawTable + ' ORDER BY row_id').all();
  const expectedRawRows = range.e.r - range.s.r;
  check(month + ' raw schema', JSON.stringify(rawColumns) === JSON.stringify(normalizedHeaders),
    `excel=${normalizedHeaders.length} db=${rawColumns.length}`);
  check(month + ' raw row count', rawRows.length === expectedRawRows && rawRows.length === 3000,
    `excel=${expectedRawRows} db=${rawRows.length}`);

  let valueMismatches = 0;
  for (let row = range.s.r + 1, rowIndex = 0; row <= range.e.r; row++, rowIndex++) {
    for (let column = range.s.c; column <= range.e.c; column++) {
      const cell = sheet[xlsx.utils.encode_cell({ r: row, c: column })];
      const expected = cell ? cell.v : null;
      const dbColumn = rawColumns[column - range.s.c];
      const actual = rawRows[rowIndex] && rawRows[rowIndex][dbColumn];
      cellsChecked++;
      if (!equivalent(expected, actual)) {
        valueMismatches++;
        if (mismatchSamples < 10) {
          mismatchSamples++;
          console.log(`[VALUE] ${month} R${row + 1}C${column + 1} expected=${JSON.stringify(expected)} actual=${JSON.stringify(actual)}`);
        }
      }
    }
  }
  check(month + ' raw values', valueMismatches === 0, 'mismatches=' + valueMismatches);

  const groups = new Map();
  for (const row of rawRows) {
    const key = listingKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const dedupRows = db.prepare('SELECT * FROM ' + dedupTable + ' ORDER BY row_id').all();
  const dedupColumnsList = db.prepare('PRAGMA table_info(' + dedupTable + ')').all()
    .map((column) => column.name).filter((name) => name !== 'row_id');
  const dedupKeys = dedupRows.map(listingKey);
  check(month + ' unique raw listing keys', groups.size === EXPECTED_DEDUP[month],
    `unique=${groups.size} expected=${EXPECTED_DEDUP[month]}`);
  check(month + ' dedup key coverage', dedupRows.length === groups.size
    && dedupKeys.every(Boolean) && new Set(dedupKeys).size === dedupRows.length
    && dedupKeys.every((key) => groups.has(key)), 'rows=' + dedupRows.length);

  let selectedRowsNotInRaw = 0;
  for (const selected of dedupRows) {
    const candidates = groups.get(listingKey(selected)) || [];
    const found = candidates.some((candidate) => dedupColumnsList.every((column) => equivalent(candidate[column], selected[column])));
    if (!found) selectedRowsNotInRaw++;
  }
  check(month + ' every canonical representative is an exact raw row', selectedRowsNotInRaw === 0,
    'notFound=' + selectedRowsNotInRaw);
}

console.log('\n========== COMPETITOR AUDIT ==========');
console.log('Checks: ' + checks);
console.log('Cells checked: ' + cellsChecked);
console.log('Failures: ' + failures);
db.close();
process.exit(failures > 0 ? 1 : 0);
