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
const SCHEMA_VERSION = '1.1.0';

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

// Excel exports in this project store many ASIN/parent-ASIN cells as rich
// text with a hyperlink. SheetJS exposes cell.v as an empty string in those
// cells while the displayed value is in cell.l.display (and/or cell.r).
// Resolve the displayed value before type inference and insertion so the
// SQLite mirror does not silently lose identifiers.
function decodeXmlEntities(value) {
  return String(value)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/gi, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCharCode(parseInt(code, 16)));
}

function cellDisplayValue(cell) {
  if (!cell) return null;
  if (cell.v !== undefined && cell.v !== null && cell.v !== '') return cell.v;
  if (cell.l && cell.l.display !== undefined && cell.l.display !== '') return cell.l.display;
  if (cell.r) {
    const richText = decodeXmlEntities(String(cell.r).replace(/<[^>]*>/g, '')).trim();
    if (richText) return richText;
  }
  if (cell.w !== undefined && cell.w !== null && cell.w !== '') return cell.w;
  return cell.v === undefined ? null : cell.v;
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
      if (cell) cols[cellDisplayValue(cell)] = c;
    }
    if (cols['品牌'] !== undefined && cols['商品标题'] !== undefined && cols['月销量'] !== undefined) {
      return r;
    }
  }
  return -1;
}

function classifySheet(name, sheet) {
  if (!sheet || !sheet['!ref']) return 'empty';
  if (/^TOP/i.test(name)) return 'top';
  if (/^\d{6}$/.test(name) || /^\d{4}\.\d{1,2}$/.test(name)) return 'monthly';
  return 'unknown';
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

function finiteNumber(value) {
  return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
}

function sumColumn(rows, column) {
  let total = 0;
  let count = 0;
  for (const row of rows) {
    if (!finiteNumber(row[column])) continue;
    total += Number(row[column]);
    count++;
  }
  return count ? total : null;
}

// The supplied workbook contains formulas in the TOP sheets without cached <v>
// values. SheetJS therefore exposes those cells as undefined when reading the
// workbook. Reconstruct the small, auditable formula set before type inference
// and insertion so the SQLite mirror contains the displayed/calculated values.
function restoreTopFormulaValues(sheetName, allRows, context) {
  const groupLabel = (rankIndex) => {
    const start = Math.floor(rankIndex / 10) * 10 + 1;
    return start + ' - ' + (start + 9);
  };

  if (sheetName === 'TOP销量 ') {
    for (let row = 0; row < Math.min(100, allRows.length); row++) {
      if (allRows[row][1] === null || allRows[row][1] === undefined) allRows[row][1] = groupLabel(row);
    }
    const totalRow = allRows[100];
    if (totalRow) {
      for (let column = 2; column < totalRow.length; column++) totalRow[column] = sumColumn(allRows.slice(0, 100), column);
    }
    return;
  }

  if (sheetName === 'TOP销量 （倍率）') {
    for (let row = 0; row < Math.min(100, allRows.length); row++) {
      if (allRows[row][1] === null || allRows[row][1] === undefined) allRows[row][1] = groupLabel(row);
      // S=R/Q, U=T/R, W=V/T, Y=X/V in the source workbook (zero-based columns).
      for (const [target, numerator, denominator] of [[18, 17, 16], [20, 19, 17], [22, 21, 19], [24, 23, 21]]) {
        const n = allRows[row][numerator];
        const d = allRows[row][denominator];
        allRows[row][target] = finiteNumber(n) && finiteNumber(d) && Number(d) !== 0 ? Number(n) / Number(d) : null;
      }
    }
    const totalRow = allRows[100];
    if (totalRow) {
      for (let column = 2; column < totalRow.length; column++) totalRow[column] = sumColumn(allRows.slice(0, 100), column);
    }
    return;
  }

  if (sheetName === 'TOP总销售额') {
    const totalRow = allRows[100];
    if (totalRow) {
      for (let column = 1; column < totalRow.length; column++) totalRow[column] = sumColumn(allRows.slice(0, 100), column);
    }
    return;
  }

  if (sheetName === 'TOP平均单价') {
    const salesRows = context.get('TOP总销售额');
    const volumeRows = context.get('TOP销量 ');
    if (!salesRows || !volumeRows) throw new Error('TOP平均单价 formula dependencies are not available');
    // B:Y (2022.06-2024.05) and AV:AY (2026.04-2026.07) are formula columns.
    const formulaColumns = [];
    for (let column = 1; column <= 24; column++) formulaColumns.push(column);
    for (let column = 47; column <= 50; column++) formulaColumns.push(column);
    for (let row = 0; row < Math.min(100, allRows.length); row++) {
      for (const column of formulaColumns) {
        const revenue = salesRows[row] && salesRows[row][column];
        const volume = volumeRows[row] && volumeRows[row][column + 1];
        allRows[row][column] = finiteNumber(revenue) && finiteNumber(volume) && Number(volume) !== 0
          ? Number(revenue) / Number(volume)
          : null;
      }
    }
  }
}

function processMonthlySheet(db, sheet, sheetName) {
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const headerRow = detectHeaderRow(sheet);
  if (headerRow < 0) { log('warn', 'No header for ' + sheetName); return 0; }

  const rawHeaders = [];
  for (let c = 0; c <= range.e.c; c++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: headerRow, c })];
    rawHeaders.push(cell ? String(cellDisplayValue(cell) ?? '') : '');
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
      row.push(cellDisplayValue(cell));
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

function processTopSheet(db, sheet, sheetName, topFormulaContext) {
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const tableName = safeTableName('top', sheetName);
  if (!tableName) return;
  const headerRow = 0;

  const rawHeaders = [];
  for (let c = 0; c <= range.e.c; c++) {
    const cell = sheet[xlsx.utils.encode_cell({ r: headerRow, c })];
    rawHeaders.push(cell ? String(cellDisplayValue(cell) ?? '') : '');
  }
  let headers = rawHeaders.map(colNorm);
  headers = dedupCols(headers);

  const allRows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      row.push(cellDisplayValue(cell));
    }
    allRows.push(row);
  }

  restoreTopFormulaValues(sheetName, allRows, topFormulaContext);

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
  topFormulaContext.set(sheetName, allRows.map((row) => row.slice()));
  return allRows.length;
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

  if (fsMod.existsSync(DB_PATH)) {
    if (!IMPORT_OVERWRITE) {
      throw new Error('Target DB already exists and IMPORT_OVERWRITE=false: ' + DB_PATH);
    }
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
    + 'visible_sheets INTEGER,'
    + 'hidden_sheets INTEGER,'
    + 'effective_sheets INTEGER,'
    + 'skipped_sheets INTEGER,'
    + 'monthly_tables INTEGER,'
    + 'summary_tables INTEGER,'
    + 'total_rows INTEGER,'
    + 'db_size_bytes INTEGER,'
    + 'schema_version TEXT);');

  db.exec('DROP TABLE IF EXISTS sheet_catalog;');
  db.exec('CREATE TABLE sheet_catalog ('
    + 'sheet_order INTEGER PRIMARY KEY,'
    + 'sheet_name TEXT NOT NULL,'
    + 'visibility TEXT NOT NULL,'
    + 'hidden_code INTEGER NOT NULL,'
    + 'sheet_ref TEXT,'
    + 'classification TEXT NOT NULL,'
    + 'target_table TEXT,'
    + 'imported_rows INTEGER NOT NULL DEFAULT 0,'
    + 'skip_reason TEXT);');

  let totalRows = 0;
  let monthlyCount = 0;
  let topCount = 0;
  let skipCount = 0;
  let visibleCount = 0;
  let hiddenCount = 0;
  const topFormulaContext = new Map();

  const sheetMetaByName = new Map(
    ((wb.Workbook && wb.Workbook.Sheets) || []).map((s) => [s.name, s]),
  );
  const catalogStmt = db.prepare('INSERT INTO sheet_catalog '
    + '(sheet_order, sheet_name, visibility, hidden_code, sheet_ref, classification, target_table, imported_rows, skip_reason) '
    + 'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)');

  for (let sheetIndex = 0; sheetIndex < wb.SheetNames.length; sheetIndex++) {
    const sheetName = wb.SheetNames[sheetIndex];
    const sheet = wb.Sheets[sheetName];
    const sheetMeta = sheetMetaByName.get(sheetName) || {};
    const hiddenCode = Number(sheetMeta.Hidden || 0);
    const visibility = hiddenCode === 0 ? 'visible' : hiddenCode === 2 ? 'very_hidden' : 'hidden';
    if (hiddenCode === 0) visibleCount++; else hiddenCount++;

    const type = classifySheet(sheetName, sheet);
    if (type === 'empty') {
      skipCount++;
      catalogStmt.run(sheetIndex + 1, sheetName, visibility, hiddenCode, null, type, null, 0, 'no_effective_range');
      continue;
    }
    if (type === 'unknown') {
      throw new Error('Unrecognized non-empty sheet: ' + sheetName);
    }
    if (type === 'top') {
      const tableName = safeTableName(type, sheetName);
      if (!tableName) throw new Error('Unmapped TOP sheet: ' + sheetName);
      topCount++;
      const importedRows = processTopSheet(db, sheet, sheetName, topFormulaContext);
      catalogStmt.run(sheetIndex + 1, sheetName, visibility, hiddenCode, sheet['!ref'], type, tableName, importedRows, null);
      continue;
    }
    if (type === 'monthly') {
      const r = processMonthlySheet(db, sheet, sheetName);
      monthlyCount++;
      totalRows += r;
      catalogStmt.run(sheetIndex + 1, sheetName, visibility, hiddenCode, sheet['!ref'], type, safeTableName(type, sheetName), r, null);
    }
  }

  db.prepare('INSERT INTO meta (imported_at, source_file, source_size_bytes, total_sheets, visible_sheets, hidden_sheets, effective_sheets, skipped_sheets, monthly_tables, summary_tables, total_rows, db_size_bytes, schema_version) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)').run(
    new Date().toISOString(),
    pathMod.basename(RAW_EXCEL_PATH),
    fsMod.statSync(RAW_EXCEL_PATH).size,
    wb.SheetNames.length,
    visibleCount,
    hiddenCount,
    monthlyCount + topCount,
    skipCount,
    monthlyCount,
    topCount,
    totalRows,
    0,
    SCHEMA_VERSION,
  );

  db.exec('PRAGMA wal_checkpoint(TRUNCATE);');
  db.close();
  let dbSize = fsMod.statSync(DB_PATH).size;
  const metaDb = new DatabaseSync(DB_PATH);
  metaDb.prepare('UPDATE meta SET db_size_bytes = ? WHERE id = 1').run(dbSize);
  metaDb.exec('PRAGMA wal_checkpoint(TRUNCATE);');
  metaDb.close();
  dbSize = fsMod.statSync(DB_PATH).size;
  const t2 = Date.now();
  log('info', 'Done in ' + ((t2 - t0) / 1000).toFixed(1) + 's');
  log('info', 'monthly=' + monthlyCount + ' top=' + topCount + ' skipped=' + skipCount + ' rows=' + totalRows);
  log('info', 'DB size: ' + (dbSize / 1024 / 1024).toFixed(2) + ' MB');
}

try { main(); } catch (e) { console.error('FATAL:', e); process.exit(1); }
