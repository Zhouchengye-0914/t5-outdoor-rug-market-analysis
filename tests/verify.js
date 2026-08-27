'use strict';
// tests/verify.js
// 验证 data/processed/market.db 与 data/raw/地垫-卖家精灵市场数据.xlsx 完全一致
// Ref: docs/SPEC.md 6. 验收标准

const fsMod = require('fs');
const pathMod = require('path');
const xlsx = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const ROOT = pathMod.resolve(__dirname, '..');
const EXCEL = pathMod.resolve(ROOT, 'data/raw/地垫-卖家精灵市场数据.xlsx');
const DB = pathMod.resolve(ROOT, process.env.VERIFY_DB_PATH || 'data/processed/market.db');

if (!fsMod.existsSync(EXCEL)) { console.error('FAIL: Excel not found'); process.exit(1); }
if (!fsMod.existsSync(DB)) { console.error('FAIL: DB not found'); process.exit(1); }

let pass = 0, fail = 0;
function check(label, ok, detail) {
  if (ok) { pass++; console.log('[PASS] ' + label + (detail ? ' (' + detail + ')' : '')); }
  else { fail++; console.log('[FAIL] ' + label + (detail ? ' (' + detail + ')' : '')); }
}

// Load Excel sheet info
const wb = xlsx.readFile(EXCEL, { cellDates: false, cellNF: false, cellFormula: false });
const sheetExcelRows = {};
for (const name of wb.SheetNames) {
  const sh = wb.Sheets[name];
  if (!sh['!ref']) { sheetExcelRows[name] = 0; continue; }
  const r = xlsx.utils.decode_range(sh['!ref']);
  // 检测 header 行 (0-indexed)
  let headerRow = -1;
  for (const hr of [0, 1]) {
    const cols = {};
    for (let c = 0; c <= r.e.c; c++) {
      const cell = sh[xlsx.utils.encode_cell({ r: hr, c })];
      if (cell) cols[cell.v] = c;
    }
    if (cols['品牌'] !== undefined && cols['商品标题'] !== undefined && cols['月销量'] !== undefined) { headerRow = hr; break; }
  }
  // TOP 表 header 在 row 0
  if (/^TOP/i.test(name) && headerRow < 0) headerRow = 0;
  // 数据行数 = total - (headerRow + 1)
  sheetExcelRows[name] = (r.e.r - r.s.r + 1) - (headerRow + 1);
}

// Open DB
const db = new DatabaseSync(DB);

// meta
const metaRows = db.prepare('SELECT * FROM meta').all();
check('meta has 1 row', metaRows.length === 1, 'rows=' + metaRows.length);
check('meta imported_at is string', metaRows[0] && typeof metaRows[0].imported_at === 'string');
check('meta source_file = xlsx', metaRows[0] && metaRows[0].source_file === '地垫-卖家精灵市场数据.xlsx');
check('meta schema_version = 1.1.0', metaRows[0] && metaRows[0].schema_version === '1.1.0');
check('meta total_sheets = 55', metaRows[0] && metaRows[0].total_sheets === 55);
check('meta visible_sheets = 23', metaRows[0] && metaRows[0].visible_sheets === 23);
check('meta hidden_sheets = 32', metaRows[0] && metaRows[0].hidden_sheets === 32);
check('meta effective_sheets = 54', metaRows[0] && metaRows[0].effective_sheets === 54);
check('meta skipped_sheets = 1', metaRows[0] && metaRows[0].skipped_sheets === 1);

const catalogRows = db.prepare('SELECT * FROM sheet_catalog ORDER BY sheet_order').all();
check('sheet_catalog has 55 rows', catalogRows.length === 55, 'rows=' + catalogRows.length);
check('sheet_catalog visible = 23', catalogRows.filter((r) => r.visibility === 'visible').length === 23);
check('sheet_catalog hidden = 32', catalogRows.filter((r) => r.visibility !== 'visible').length === 32);
check('sheet_catalog effective = 54', catalogRows.filter((r) => r.target_table).length === 54);
check('sheet_catalog skipped no-range = 1', catalogRows.filter((r) => r.skip_reason === 'no_effective_range').length === 1);

// tables
const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").all();
const tableNames = tables.map((t) => t.name);
console.log('\nTotal tables in DB: ' + tableNames.length);

function safeSheetName(sn) {
  if (sn === 'Sheet6') return null;
  if (/^TOP/i.test(sn)) {
    const map = {
      'TOP销量 ': 'top_sales_volume',
      'TOP销量 （倍率）': 'top_sales_volume_ratio',
      'TOP总销售额': 'top_total_sales',
      'TOP平均单价': 'top_avg_price'
    };
    return map[sn];
  }
  if (/^\d{6}$/.test(sn)) return 'monthly_' + sn;
  const m = sn.match(/^(\d{4})\.(\d+)$/);
  if (m) return 'monthly_' + m[1] + m[2].padStart(2, '0');
  return null;
}

let totalMismatch = 0;
let totalChecked = 0;
for (const sn of wb.SheetNames) {
  const tn = safeSheetName(sn);
  if (!tn) continue;
  if (!tableNames.includes(tn)) {
    fail++;
    console.log('[FAIL] table missing: ' + tn);
    continue;
  }
  const excelRows = sheetExcelRows[sn];
  const dbRow = db.prepare('SELECT COUNT(*) AS c FROM ' + tn).get();
  const dbRows = dbRow.c;
  totalChecked++;
  if (excelRows === dbRows) { pass++; }
  else { fail++; totalMismatch++; console.log('[FAIL] ' + tn + ' excel=' + excelRows + ' db=' + dbRows); }
}
console.log('\nChecked ' + totalChecked + ' tables, mismatches=' + totalMismatch);

// sample
function pickSample(sheetName) {
  const sheet = wb.Sheets[sheetName];
  const range = xlsx.utils.decode_range(sheet['!ref']);
  let headerRow = -1;
  for (const r of [0, 1]) {
    const cols = {};
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      if (cell) cols[cell.v] = c;
    }
    if (cols['品牌'] !== undefined && cols['商品标题'] !== undefined && cols['月销量'] !== undefined) { headerRow = r; break; }
  }
  if (headerRow < 0) return [];
  const picks = [];
  for (let r = headerRow + 1; r <= Math.min(headerRow + 3, range.e.r); r++) picks.push(r);
  for (let r = Math.max(headerRow + 1, range.e.r - 2); r <= range.e.r; r++) picks.push(r);
  const samples = [];
  for (const r of picks) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : null);
    }
    samples.push(row);
  }
  return samples;
}

function getDbSample(tableName) {
  const totalRow = db.prepare('SELECT COUNT(*) AS c FROM ' + tableName).get().c;
  const firstThree = db.prepare('SELECT * FROM ' + tableName + ' ORDER BY row_id ASC LIMIT 3').all();
  const lastThree = db.prepare('SELECT * FROM ' + tableName + ' ORDER BY row_id DESC LIMIT 3').all().reverse();
  const rows = firstThree.concat(lastThree);
  return rows.map((r) => {
    const vals = [];
    for (const k of Object.keys(r)) { if (k !== 'row_id' && k !== 'month_label') vals.push(r[k]); }
    return vals;
  });
}

console.log('\n--- Sample data comparison ---');
const sampleSheets = ['202206', '2025.6', '2026.7'];
for (const sn of sampleSheets) {
  const tn = safeSheetName(sn);
  if (!tableNames.includes(tn)) continue;
  const es = pickSample(sn);
  const ds = getDbSample(tn);
  let same = es.length === ds.length;
  if (same) {
    for (let i = 0; i < es.length; i++) {
      const e = es[i];
      const d = ds[i];
      if (e.length !== d.length) { same = false; break; }
      for (let j = 0; j < e.length; j++) {
        const ev = e[j];
        const dv = d[j];
        if (ev === null || ev === undefined || ev === '') {
          if (dv !== null) { same = false; break; }
        } else if (typeof ev === 'number') {
          if (Math.abs(Number(dv) - ev) > 0.001) { same = false; break; }
        } else {
          if (String(dv) !== String(ev)) { same = false; break; }
        }
      }
      if (!same) break;
    }
  }
  check('sample ' + tn, same, 'rows=' + es.length);
}

console.log('\n========== Summary ==========');
console.log('PASS: ' + pass);
console.log('FAIL: ' + fail);
db.close();
process.exit(fail > 0 ? 1 : 0);
