'use strict';

const path = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const EXCEL = path.resolve(ROOT, 'data/raw/地垫-卖家精灵市场数据.xlsx');
const DB_PATH = path.resolve(ROOT, process.env.VERIFY_DB_PATH || 'data/processed/market.db');

function normMonth(name) {
  if (/^\d{6}$/.test(name)) return name;
  const m = name.match(/^(\d{4})\.(\d+)$/);
  return m ? m[1] + m[2].padStart(2, '0') : name;
}

function tableName(sheetName) {
  const top = {
    'TOP销量 ': 'top_sales_volume',
    'TOP销量 （倍率）': 'top_sales_volume_ratio',
    'TOP总销售额': 'top_total_sales',
    'TOP平均单价': 'top_avg_price',
  };
  if (top[sheetName]) return top[sheetName];
  if (/^\d{6}$/.test(sheetName) || /^\d{4}\.\d{1,2}$/.test(sheetName)) {
    return 'monthly_' + normMonth(sheetName);
  }
  return null;
}

function detectHeaderRow(sheet, isTop) {
  if (isTop) return 0;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  for (const r of [0, 1]) {
    const values = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      values.push(cell && cell.v);
    }
    if (values.includes('品牌') && values.includes('商品标题') && values.includes('月销量')) return r;
  }
  return -1;
}

const wb = xlsx.readFile(EXCEL, { cellDates: false, cellNF: false, cellFormula: false });
const db = new DatabaseSync(DB_PATH, { readOnly: true });
let sheetsChecked = 0;
let cellsChecked = 0;
let structuralIssues = 0;
let valueMismatches = 0;
let numericTypeChanges = 0;

for (const sheetName of wb.SheetNames) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) continue;
  const target = tableName(sheetName);
  if (!target) {
    console.log('[STRUCT] unmapped sheet: ' + sheetName);
    structuralIssues++;
    continue;
  }

  const isTop = /^TOP/i.test(sheetName);
  const headerRow = detectHeaderRow(sheet, isTop);
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const columns = db.prepare('PRAGMA table_info(' + target + ')').all()
    .map((r) => r.name)
    .filter((name) => name !== 'row_id' && name !== 'month_label');
  const rows = db.prepare('SELECT * FROM ' + target + ' ORDER BY row_id').all();
  const expectedRows = range.e.r - headerRow;
  const expectedCols = range.e.c + 1;
  if (headerRow < 0 || rows.length !== expectedRows || columns.length !== expectedCols) {
    console.log('[STRUCT] ' + sheetName + ' expected rows/cols=' + expectedRows + '/' + expectedCols
      + ' actual=' + rows.length + '/' + columns.length);
    structuralIssues++;
  }

  for (let r = headerRow + 1, rowIndex = 0; r <= range.e.r; r++, rowIndex++) {
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      const excelValue = cell ? cell.v : null;
      const dbValue = rows[rowIndex] ? rows[rowIndex][columns[c]] : undefined;
      let equal;
      if (excelValue === null || excelValue === undefined || excelValue === '') {
        equal = dbValue === null;
      } else if (typeof excelValue === 'number') {
        equal = dbValue !== null && !Number.isNaN(Number(dbValue))
          && Math.abs(Number(dbValue) - excelValue) <= 1e-9;
        if (equal && typeof dbValue !== 'number') numericTypeChanges++;
      } else {
        equal = String(dbValue) === String(excelValue);
      }
      cellsChecked++;
      if (!equal) {
        valueMismatches++;
        if (valueMismatches <= 20) {
          console.log('[VALUE] ' + sheetName + ' R' + (r + 1) + 'C' + (c + 1)
            + ' excel=' + JSON.stringify(excelValue) + ' db=' + JSON.stringify(dbValue));
        }
      }
    }
  }
  sheetsChecked++;
}

console.log('\n========== FULL AUDIT ==========');
console.log('Sheets checked: ' + sheetsChecked);
console.log('Cells checked: ' + cellsChecked);
console.log('Structural issues: ' + structuralIssues);
console.log('Value mismatches: ' + valueMismatches);
console.log('Numeric type changes: ' + numericTypeChanges);
db.close();
process.exit(structuralIssues || valueMismatches ? 1 : 0);
