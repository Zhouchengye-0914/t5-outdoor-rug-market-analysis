'use strict';

// Full active-source audit. Non-replaced tables are checked cell-by-cell against the base Excel;
// 2026.01-07 are checked cell-by-cell against canonical competitor dedup tables.

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const EXCEL = path.resolve(ROOT, 'data/raw/地垫-卖家精灵市场数据.xlsx');
const DB_PATH = path.resolve(ROOT, process.env.VERIFY_DB_PATH || 'data/processed/market.db');
const COMPETITOR_DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const REPLACED_MONTHS = new Set(['202601', '202602', '202603', '202604', '202605', '202606', '202607']);

for (const filePath of [EXCEL, DB_PATH, COMPETITOR_DB_PATH]) {
  if (!fs.existsSync(filePath)) {
    console.error('Missing audit input: ' + filePath);
    process.exit(1);
  }
}

function normMonth(name) {
  if (/^\d{6}$/.test(name)) return name;
  const match = name.match(/^(\d{4})\.(\d{1,2})$/);
  return match ? match[1] + match[2].padStart(2, '0') : null;
}

function tableName(sheetName) {
  const top = {
    'TOP销量 ': 'top_sales_volume',
    'TOP销量 （倍率）': 'top_sales_volume_ratio',
    'TOP总销售额': 'top_total_sales',
    'TOP平均单价': 'top_avg_price',
  };
  return top[sheetName] || (normMonth(sheetName) ? 'monthly_' + normMonth(sheetName) : null);
}

function colNorm(name) {
  if (!name) return '';
  let value = String(name).trim();
  value = value.replace(/[\$\(\)\'`]/g, '_');
  value = value.replace(/&/g, '_and_');
  value = value.replace(/\s+/g, '_');
  value = value.replace(/[^A-Za-z0-9_\u4e00-\u9fff]/g, '_');
  value = value.replace(/_+/g, '_').replace(/^_+|_+$/g, '');
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

function detectHeaderRow(sheet, isTop) {
  if (isTop) return 0;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  for (const row of [0, 1]) {
    const values = [];
    for (let column = range.s.c; column <= range.e.c; column++) {
      const cell = sheet[xlsx.utils.encode_cell({ r: row, c: column })];
      values.push(cell && cell.v);
    }
    if (values.includes('品牌') && values.includes('商品标题') && values.includes('月销量')) return row;
  }
  return -1;
}

function expectedColumns(sheet, headerRow) {
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const raw = [];
  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: headerRow, c: column })];
    raw.push(colNorm(cell ? cell.v : ''));
  }
  return dedupColumns(raw);
}

function listingKey(row) {
  return String(row['父ASIN'] || '').trim() || String(row.ASIN || '').trim();
}

function equivalent(expected, actual) {
  if (expected === null || expected === undefined || expected === '') return actual === null;
  if (typeof expected === 'number') {
    return actual !== null && Number.isFinite(Number(actual)) && Math.abs(Number(actual) - expected) <= 1e-9;
  }
  return String(actual) === String(expected);
}

const workbook = xlsx.readFile(EXCEL, { cellDates: false, cellNF: false, cellFormula: false });
const db = new DatabaseSync(DB_PATH, { readOnly: true });
const competitor = new DatabaseSync(COMPETITOR_DB_PATH, { readOnly: true });
let sheetsChecked = 0;
let cellsChecked = 0;
let structuralIssues = 0;
let valueMismatches = 0;
let numericTypeChanges = 0;

function reportMismatch(label, expected, actual) {
  valueMismatches++;
  if (valueMismatches <= 20) {
    console.log('[VALUE] ' + label + ' expected=' + JSON.stringify(expected) + ' actual=' + JSON.stringify(actual));
  }
}

for (const sheetName of workbook.SheetNames) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) continue;
  const target = tableName(sheetName);
  if (!target) {
    console.log('[STRUCT] unmapped non-empty sheet: ' + sheetName);
    structuralIssues++;
    continue;
  }

  const headerRow = detectHeaderRow(sheet, /^TOP/i.test(sheetName));
  if (headerRow < 0) {
    console.log('[STRUCT] header not found: ' + sheetName);
    structuralIssues++;
    continue;
  }
  const excelColumns = expectedColumns(sheet, headerRow);
  const targetColumns = db.prepare('PRAGMA table_info(' + target + ')').all()
    .map((column) => column.name).filter((name) => !['row_id', 'month_label'].includes(name));
  if (JSON.stringify(excelColumns) !== JSON.stringify(targetColumns)) {
    console.log('[STRUCT] schema mismatch ' + sheetName + ' expected=' + excelColumns.length + ' actual=' + targetColumns.length);
    structuralIssues++;
  }

  const month = target.startsWith('monthly_') ? target.slice('monthly_'.length) : null;
  const targetRows = db.prepare('SELECT * FROM ' + target + ' ORDER BY row_id').all();
  if (month && REPLACED_MONTHS.has(month)) {
    const sourceTable = 'dedup_' + month;
    const sourceRows = competitor.prepare('SELECT * FROM ' + sourceTable + ' ORDER BY row_id').all();
    const sourceColumns = new Set(competitor.prepare('PRAGMA table_info(' + sourceTable + ')').all().map((column) => column.name));
    if (targetRows.length !== sourceRows.length) {
      console.log('[STRUCT] replacement row mismatch ' + month + ' expected=' + sourceRows.length + ' actual=' + targetRows.length);
      structuralIssues++;
    }
    const targetByKey = new Map(targetRows.map((row) => [listingKey(row), row]));
    if (targetByKey.size !== targetRows.length || targetRows.some((row) => !listingKey(row))) {
      console.log('[STRUCT] blank/duplicate listing key in ' + target);
      structuralIssues++;
    }
    for (const sourceRow of sourceRows) {
      const key = listingKey(sourceRow);
      const targetRow = targetByKey.get(key);
      if (!targetRow) {
        reportMismatch(target + ' missing key ' + key, 'present', 'missing');
        continue;
      }
      if (targetRow.month_label !== month) reportMismatch(target + ' key=' + key + ' month_label', month, targetRow.month_label);
      for (const column of targetColumns) {
        const expected = sourceColumns.has(column) ? sourceRow[column] : null;
        const actual = targetRow[column];
        cellsChecked++;
        if (!equivalent(expected, actual)) reportMismatch(target + ' key=' + key + ' column=' + column, expected, actual);
        else if (typeof expected === 'number' && typeof actual !== 'number') numericTypeChanges++;
      }
    }
  } else {
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const expectedRows = range.e.r - headerRow;
    if (targetRows.length !== expectedRows) {
      console.log('[STRUCT] row mismatch ' + sheetName + ' expected=' + expectedRows + ' actual=' + targetRows.length);
      structuralIssues++;
    }
    for (let row = headerRow + 1, rowIndex = 0; row <= range.e.r; row++, rowIndex++) {
      for (let column = range.s.c; column <= range.e.c; column++) {
        const cell = sheet[xlsx.utils.encode_cell({ r: row, c: column })];
        const expected = cell ? cell.v : null;
        const targetColumn = targetColumns[column - range.s.c];
        const actual = targetRows[rowIndex] ? targetRows[rowIndex][targetColumn] : undefined;
        cellsChecked++;
        if (!equivalent(expected, actual)) {
          reportMismatch(sheetName + ' R' + (row + 1) + 'C' + (column + 1), expected, actual);
        } else if (typeof expected === 'number' && typeof actual !== 'number') {
          numericTypeChanges++;
        }
      }
    }
  }
  sheetsChecked++;
}

console.log('\n========== FULL ACTIVE-SOURCE AUDIT ==========');
console.log('Sheets checked: ' + sheetsChecked);
console.log('Cells checked: ' + cellsChecked);
console.log('Structural issues: ' + structuralIssues);
console.log('Value mismatches: ' + valueMismatches);
console.log('Numeric type changes: ' + numericTypeChanges + ' (informational)');
competitor.close();
db.close();
process.exit(structuralIssues || valueMismatches ? 1 : 0);
