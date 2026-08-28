'use strict';

// Fast acceptance checks for the two-stage data pipeline:
// base workbook -> market.db, then canonical competitor parent-level replacements for 2026.01-07.

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const EXCEL = path.resolve(ROOT, 'data/raw/地垫-卖家精灵市场数据.xlsx');
const DB_PATH = path.resolve(ROOT, process.env.VERIFY_DB_PATH || 'data/processed/market.db');
const COMPETITOR_DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const REPLACED_MONTHS = new Set(['202601', '202602', '202603', '202604', '202605', '202606', '202607']);

for (const [label, filePath] of [['Excel', EXCEL], ['market DB', DB_PATH], ['competitor DB', COMPETITOR_DB_PATH]]) {
  if (!fs.existsSync(filePath)) {
    console.error(`FAIL: ${label} not found: ${filePath}`);
    process.exit(1);
  }
}

let pass = 0;
let fail = 0;
function check(label, ok, detail) {
  if (ok) {
    pass++;
    console.log('[PASS] ' + label + (detail ? ' (' + detail + ')' : ''));
  } else {
    fail++;
    console.log('[FAIL] ' + label + (detail ? ' (' + detail + ')' : ''));
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

function detectHeaderRow(sheet, isTop = false) {
  if (isTop) return 0;
  if (!sheet || !sheet['!ref']) return -1;
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

function excelRowCount(sheetName, sheet) {
  if (!sheet || !sheet['!ref']) return 0;
  const header = detectHeaderRow(sheet, /^TOP/i.test(sheetName));
  const range = xlsx.utils.decode_range(sheet['!ref']);
  return header < 0 ? 0 : range.e.r - header;
}

function equivalent(expected, actual) {
  if (expected === null || expected === undefined || expected === '') return actual === null;
  if (typeof expected === 'number') {
    return actual !== null && Number.isFinite(Number(actual)) && Math.abs(Number(actual) - expected) <= 1e-9;
  }
  return String(actual) === String(expected);
}

function excelSamples(sheetName, sheet) {
  const header = detectHeaderRow(sheet, /^TOP/i.test(sheetName));
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const picks = [];
  for (let row = header + 1; row <= Math.min(header + 3, range.e.r); row++) picks.push(row);
  for (let row = Math.max(header + 1, range.e.r - 2); row <= range.e.r; row++) picks.push(row);
  return picks.map((row) => {
    const values = [];
    for (let column = range.s.c; column <= range.e.c; column++) {
      const cell = sheet[xlsx.utils.encode_cell({ r: row, c: column })];
      values.push(cell ? cell.v : null);
    }
    return values;
  });
}

function dbSamples(db, targetTable, columns) {
  const first = db.prepare('SELECT * FROM ' + targetTable + ' ORDER BY row_id LIMIT 3').all();
  const last = db.prepare('SELECT * FROM ' + targetTable + ' ORDER BY row_id DESC LIMIT 3').all().reverse();
  return first.concat(last).map((row) => columns.map((column) => row[column]));
}

const workbook = xlsx.readFile(EXCEL, { cellDates: false, cellNF: false, cellFormula: false });
const db = new DatabaseSync(DB_PATH, { readOnly: true });
const competitor = new DatabaseSync(COMPETITOR_DB_PATH, { readOnly: true });

const metaRows = db.prepare('SELECT * FROM meta').all();
check('meta has 1 row', metaRows.length === 1, 'rows=' + metaRows.length);
check('meta imported_at is string', metaRows[0] && typeof metaRows[0].imported_at === 'string');
check('meta source_file = xlsx', metaRows[0] && metaRows[0].source_file === '地垫-卖家精灵市场数据.xlsx');
check('meta schema_version = 1.1.0', metaRows[0] && metaRows[0].schema_version === '1.1.0');
check('meta total_sheets = 55', metaRows[0] && metaRows[0].total_sheets === 55);
check('meta visible/hidden = 23/32', metaRows[0] && metaRows[0].visible_sheets === 23 && metaRows[0].hidden_sheets === 32);
check('meta effective/skipped = 54/1', metaRows[0] && metaRows[0].effective_sheets === 54 && metaRows[0].skipped_sheets === 1);

const catalogRows = db.prepare('SELECT * FROM sheet_catalog ORDER BY sheet_order').all();
check('sheet_catalog has 55 rows', catalogRows.length === 55, 'rows=' + catalogRows.length);
check('sheet_catalog visible/hidden = 23/32', catalogRows.filter((row) => row.visibility === 'visible').length === 23
  && catalogRows.filter((row) => row.visibility !== 'visible').length === 32);
check('sheet_catalog effective/skipped = 54/1', catalogRows.filter((row) => row.target_table).length === 54
  && catalogRows.filter((row) => row.skip_reason === 'no_effective_range').length === 1);

const replacements = db.prepare('SELECT * FROM analysis_replacements ORDER BY month').all();
check('analysis_replacements has 7 months', replacements.length === 7, 'rows=' + replacements.length);
check('replacement months are 202601-202607', replacements.every((row) => REPLACED_MONTHS.has(row.month)));
check('replacement raw source has 3000 rows/month', replacements.every((row) => row.source_raw_rows === 3000));
check('replacement source hash is recorded', replacements.every((row) => /^[0-9a-f]{64}$/.test(row.source_sha256)));

const tableNames = new Set(db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all().map((row) => row.name));
let checkedTables = 0;
let rowMismatches = 0;
for (const sheetName of workbook.SheetNames) {
  const target = tableName(sheetName);
  if (!target) continue;
  checkedTables++;
  if (!tableNames.has(target)) {
    check('table exists ' + target, false);
    continue;
  }
  const month = target.startsWith('monthly_') ? target.slice('monthly_'.length) : null;
  const expectedRows = month && REPLACED_MONTHS.has(month)
    ? competitor.prepare('SELECT COUNT(*) AS count FROM dedup_' + month).get().count
    : excelRowCount(sheetName, workbook.Sheets[sheetName]);
  const actualRows = db.prepare('SELECT COUNT(*) AS count FROM ' + target).get().count;
  if (expectedRows !== actualRows) rowMismatches++;
  check('row count ' + target, expectedRows === actualRows, `expected=${expectedRows} actual=${actualRows}`);
  const catalog = catalogRows.find((row) => row.target_table === target);
  check('catalog current row count ' + target, catalog && catalog.imported_rows === actualRows,
    'catalog=' + (catalog && catalog.imported_rows) + ' actual=' + actualRows);
}
check('all 54 effective tables checked', checkedTables === 54, 'tables=' + checkedTables);
check('all table row counts match active source', rowMismatches === 0, 'mismatches=' + rowMismatches);

for (const sheetName of ['202206', '2025.6']) {
  const target = tableName(sheetName);
  const columns = db.prepare('PRAGMA table_info(' + target + ')').all()
    .map((column) => column.name).filter((name) => !['row_id', 'month_label'].includes(name));
  const expected = excelSamples(sheetName, workbook.Sheets[sheetName]);
  const actual = dbSamples(db, target, columns);
  const same = expected.length === actual.length && expected.every((row, rowIndex) =>
    row.length === actual[rowIndex].length && row.every((value, columnIndex) => equivalent(value, actual[rowIndex][columnIndex])));
  check('base workbook sample ' + target, same, 'rows=' + expected.length);
}

for (const month of ['202601', '202607']) {
  const target = 'monthly_' + month;
  const source = 'dedup_' + month;
  const targetColumns = db.prepare('PRAGMA table_info(' + target + ')').all()
    .map((column) => column.name).filter((name) => !['row_id', 'month_label'].includes(name));
  const sourceColumns = new Set(competitor.prepare('PRAGMA table_info(' + source + ')').all().map((column) => column.name));
  const sourceSampleRows = competitor.prepare('SELECT * FROM ' + source + ' ORDER BY row_id LIMIT 3').all()
    .concat(competitor.prepare('SELECT * FROM ' + source + ' ORDER BY row_id DESC LIMIT 3').all().reverse());
  const actualRows = dbSamples(db, target, targetColumns);
  const same = sourceSampleRows.length === actualRows.length && sourceSampleRows.every((sourceRow, rowIndex) =>
    targetColumns.every((column, columnIndex) => equivalent(sourceColumns.has(column) ? sourceRow[column] : null, actualRows[rowIndex][columnIndex])));
  check('competitor replacement sample ' + target, same, 'rows=' + sourceSampleRows.length);
}

console.log('\n========== VERIFY SUMMARY ==========');
console.log('PASS: ' + pass);
console.log('FAIL: ' + fail);
competitor.close();
db.close();
process.exit(fail > 0 ? 1 : 0);
