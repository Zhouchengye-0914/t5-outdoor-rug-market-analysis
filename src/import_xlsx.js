'use strict';
// src/import_xlsx.js
// Excel -> SQLite converter (no data loss)
// Ref: docs/SPEC.md

const fsMod = require('fs');
const pathMod = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');
require('dotenv').config({ path: pathMod.resolve(__dirname, '..', '.env') });

// ====== Config ======
const ROOT = pathMod.resolve(__dirname, '..');
const RAW_EXCEL_PATH = pathMod.resolve(ROOT, process.env.RAW_EXCEL_PATH || 'data/raw/地垫-卖家精灵市场数据.xlsx');
const DB_PATH = pathMod.resolve(ROOT, (process.env.DATABASE_URL || 'sqlite:///data/processed/market.db').replace(/^sqlite:\/\/\/?/, ''));
const IMPORT_OVERWRITE = (process.env.IMPORT_OVERWRITE || 'true').toLowerCase() !== 'false';
const LOG_LEVEL = process.env.LOG_LEVEL || 'info';
const SCHEMA_VERSION = '1.0.0';

// ====== Logger ======
const LOG_ORDER = { debug: 0, info: 1, warn: 2, error: 3 };
function log(level, msg) {
  if (LOG_ORDER[level] >= LOG_ORDER[LOG_LEVEL]) console.log('[' + level.toUpperCase() + '] ' + msg);
}

// ====== Utils ======
function normMonth(name) {
  if (/^\d{6}$/.test(name)) return name;
  const m = name.match(/^(\d{4})\.(\d+)$/);
  if (m) return m[1] + m[2].padStart(2, '0');
  return name;
}

function colNorm(name) {
  if (!name) return '';
  let s = String(name).trim();
  s = s.replace(/[\$\(\)\'`]/g, '_');
  s = s.replace(/&/g, '_and_');
  s = s.replace(/\s+/g, '_');
  s = s.replace(/[^A-Za-z0-9_\u4e00-\u9fff]/g, '_');
  s = s.replace(/_+/g, '_');
  s = s.replace(/^_+|_+$/g, '');
  if (/^\d/.test(s)) s = '_' + s;
  if (!s) s = 'col';
  return s;
}

function parseNum(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  if (s === '' || s === '-' || s === 'N/A') return null;
  if (s.endsWith('%')) {
    const n = parseFloat(s.slice(0, -1));
    return isNaN(n) ? null : n / 100;
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function inferType(values) {
  let hasInt = false, hasFloat = false, hasText = false, nonNullCount = 0;
  for (const v of values) {
    if (v === null || v === undefined || v === '') continue;
    nonNullCount++;
    if (typeof v === 'number') {
      if (Number.isInteger(v)) hasInt = true; else hasFloat = true;
      continue;
    }
    const s = String(v).trim();
    if (s === '') continue;
    if (/^-?\d+$/.test(s)) hasInt = true;
    else if (/^-?\d+(\.\d+)?$/.test(s)) hasFloat = true;
    else { hasText = true; return 'TEXT'; }
  }
  if (hasText) return 'TEXT';
  if (nonNullCount === 0) return 'TEXT';
  if (hasFloat) return 'REAL';
  if (hasInt) return 'INTEGER';
  return 'TEXT';
}

function dedupCols(cols) {
  const seen = {};
  return cols.map((c) => {
    if (!seen[c]) { seen[c] = 1; return c; }
    seen[c]++;
    return c + '_' + seen[c];
  });
}

function detectHeaderRow(sheet) {
  if (!sheet['!ref']) return -1;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  for (const r of [0, 1]) {
    if (r > range.e.r) continue;
    const cols = {};
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      if (cell) cols[cell.v] = c;
    }
    if (cols['品牌'] !== undefined && cols['商品标题'] !== undefined && cols['月销量'] !== undefined) {
      return r;
    }
  }
  return -1;
}

function classifySheet(name) {
  if (name === 'Sheet6') return 'empty';
  if (/^TOP/i.test(name)) return 'top';
  return 'monthly';
}

function safeTableName(type, rawName) {
  if (type === 'top') {
    const map = {
      'TOP销量 ': 'top_sales_volume',
      'TOP销量 （倍率）': 'top_sales_volume_ratio',
      'TOP总销售额': 'top_total_sales',
      'TOP平均单价': 'top_avg_price'
    };
    return map[rawName] || 'top_' + colNorm(rawName);
  }
  if (type === 'monthly') return 'monthly_' + normMonth(rawName);
  return null;
}

function processMonthlySheet(db, sheet, sheetName) {
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const headerRow = detectHeaderRow(sheet);
  if (headerRow < 0) { log('warn', 'No header for ' + sheetName); return 0; }

  const rawHeaders = [];
  for (let c = 0; c <= range.e.c; c++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: headerRow, c })];
    rawHeaders.push(cell ? String(cell.v) : '');
  }
  let headers = rawHeaders.map(colNorm);
  headers = dedupCols(headers);

  const tableName = safeTableName('monthly', sheetName);
  const monthLabel = normMonth(sheetName);

  const allRows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : null);
    }
    allRows.push(row);
  }

  const colTypes = [];
  for (let c = 0; c < headers.length; c++) {
    colTypes.push(inferType(allRows.map((row) => row[c])));
  }

  db.exec('DROP TABLE IF EXISTS ' + tableName + ';');
  let ddl = 'CREATE TABLE ' + tableName + ' (row_id INTEGER PRIMARY KEY AUTOINCREMENT, month_label TEXT';
  for (let i = 0; i < headers.length; i++) {
    const colName = headers[i] || '_col_' + i;
    ddl += ', [' + colName + '] ' + colTypes[i];
  }
  ddl += ');';
  db.exec(ddl);

  const validIdx = headers.map((h, i) => i);
  const insertCols = ['month_label', ...validIdx.map((i) => headers[i] || '_col_' + i)];
  const placeholders = insertCols.map(() => '?').join(', ');
  const colList = insertCols.map((c) => '[' + c + ']').join(', ');
  const stmt = db.prepare('INSERT INTO ' + tableName + ' (' + colList + ') VALUES (' + placeholders + ')');

  db.exec('BEGIN');
  for (let i = 0; i < allRows.length; i++) {
    const row = [monthLabel];
    for (const c of validIdx) {
      let v = allRows[i][c];
      if (v === null || v === undefined || v === '') { row.push(null); continue; }
      if (colTypes[c] === 'INTEGER' || colTypes[c] === 'REAL') {
        row.push(parseNum(v));
      } else {
        row.push(String(v));
      }
    }
    stmt.run(...row);
  }
  db.exec('COMMIT');

  log('info', 'monthly ' + sheetName + ' -> ' + tableName + ' (' + allRows.length + ' rows, ' + headers.length + ' cols)');
  return allRows.length;
}

function processTopSheet(db, sheet, sheetName) {
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const tableName = safeTableName('top', sheetName);
  if (!tableName) return;
  const headerRow = 0;

  const rawHeaders = [];
  for (let c = 0; c <= range.e.c; c++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: headerRow, c })];
    rawHeaders.push(cell ? String(cell.v) : '');
  }
  let headers = rawHeaders.map(colNorm);
  headers = dedupCols(headers);

  const allRows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : null);
    }
    allRows.push(row);
  }

  const colTypes = [];
  for (let c = 0; c < headers.length; c++) {
    colTypes.push(inferType(allRows.map((row) => row[c])));
  }

  db.exec('DROP TABLE IF EXISTS ' + tableName + ';');
  let ddl = 'CREATE TABLE ' + tableName + ' (row_id INTEGER PRIMARY KEY AUTOINCREMENT, month_label TEXT';
  for (let i = 0; i < headers.length; i++) {
    const colName = headers[i] || '_col_' + i;
    ddl += ', [' + colName + '] ' + colTypes[i];
  }
  ddl += ');';
  db.exec(ddl);

  const validIdx = headers.map((h, i) => i);
  const insertCols = ['month_label', ...validIdx.map((i) => headers[i] || '_col_' + i)];
  const placeholders = insertCols.map(() => '?').join(', ');
  const colList = insertCols.map((c) => '[' + c + ']').join(', ');
  const stmt = db.prepare('INSERT INTO ' + tableName + ' (' + colList + ') VALUES (' + placeholders + ')');

  db.exec('BEGIN');
  for (let i = 0; i < allRows.length; i++) {
    const row = [sheetName];
    for (const c of validIdx) {
      let v = allRows[i][c];
      if (v === null || v === undefined || v === '') { row.push(null); continue; }
      if (colTypes[c] === 'INTEGER' || colTypes[c] === 'REAL') {
        row.push(parseNum(v));
      } else {
        row.push(String(v));
      }
    }
    stmt.run(...row);
  }
  db.exec('COMMIT');

  log('info', 'top ' + sheetName + ' -> ' + tableName + ' (' + allRows.length + ' rows, ' + headers.length + ' cols)');
}

function main() {
  log('info', 'Starting Excel -> SQLite conversion');
  log('info', 'Excel: ' + RAW_EXCEL_PATH);
  log('info', 'DB: ' + DB_PATH);

  if (!fsMod.existsSync(RAW_EXCEL_PATH)) {
    log('error', 'Excel not found: ' + RAW_EXCEL_PATH);
    process.exit(1);
  }

  const dbDir = pathMod.dirname(DB_PATH);
  if (!fsMod.existsSync(dbDir)) fsMod.mkdirSync(dbDir, { recursive: true });

  if (IMPORT_OVERWRITE && fsMod.existsSync(DB_PATH)) {
    fsMod.unlinkSync(DB_PATH);
    log('info', 'Removed existing DB');
  }
  for (const ext of ['-wal', '-shm']) {
    const p = DB_PATH + ext;
    if (fsMod.existsSync(p)) fsMod.unlinkSync(p);
  }

  const t0 = Date.now();
  const wb = xlsx.readFile(RAW_EXCEL_PATH, { cellDates: false, cellNF: false, cellFormula: false });
  const t1 = Date.now();
  log('info', 'Excel loaded in ' + ((t1 - t0) / 1000).toFixed(1) + 's');

  const db = new DatabaseSync(DB_PATH);
  db.exec('PRAGMA journal_mode = WAL;');
  db.exec('PRAGMA synchronous = NORMAL;');

  db.exec('DROP TABLE IF EXISTS meta;');
  db.exec('CREATE TABLE meta ('
    + 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
    + 'imported_at TEXT NOT NULL,'
    + 'source_file TEXT,'
    + 'source_size_bytes INTEGER,'
    + 'total_sheets INTEGER,'
    + 'monthly_tables INTEGER,'
    + 'summary_tables INTEGER,'
    + 'total_rows INTEGER,'
    + 'db_size_bytes INTEGER,'
    + 'schema_version TEXT);');

  let totalRows = 0;
  let monthlyCount = 0;
  let topCount = 0;
  let skipCount = 0;

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const type = classifySheet(sheetName);
    if (type === 'empty') { skipCount++; continue; }
    if (type === 'top') { topCount++; processTopSheet(db, sheet, sheetName); continue; }
    if (type === 'monthly') {
      monthlyCount++;
      const r = processMonthlySheet(db, sheet, sheetName);
      totalRows += r;
    }
  }

  const dbSize = fsMod.existsSync(DB_PATH) ? fsMod.statSync(DB_PATH).size : 0;
  db.prepare('INSERT INTO meta (imported_at, source_file, source_size_bytes, total_sheets, monthly_tables, summary_tables, total_rows, db_size_bytes, schema_version) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)').run(
    new Date().toISOString(),
    pathMod.basename(RAW_EXCEL_PATH),
    fsMod.statSync(RAW_EXCEL_PATH).size,
    wb.SheetNames.length,
    monthlyCount,
    topCount,
    totalRows,
    dbSize,
    SCHEMA_VERSION,
  );

  db.close();
  const t2 = Date.now();
  log('info', 'Done in ' + ((t2 - t0) / 1000).toFixed(1) + 's');
  log('info', 'monthly=' + monthlyCount + ' top=' + topCount + ' skipped=' + skipCount + ' rows=' + totalRows);
  log('info', 'DB size: ' + (dbSize / 1024 / 1024).toFixed(2) + ' MB');
}

try { main(); } catch (e) { console.error('FATAL:', e); process.exit(1); }
