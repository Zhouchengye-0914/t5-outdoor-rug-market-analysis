'use strict';

// 从 data/raw 的 2026 竞品快照（Competitor-US-2026.XX-*.xlsx）重建竞品审计库（raw_YYYYMM + dedup_YYYYMM）。
// 该库是 2026.01-07 指标代表行的唯一事实来源：apply_competitor_2026.js 从 dedup_YYYYMM 确定性重放到 market.db。
// dedup 规则（确定性、可复跑）：每个 Listing 键（父ASIN 优先，缺省用 ASIN）取原始表中第一条记录为代表行。
// 2026.01-06 为全市场 listing 级导出（每月约 1000-2000 个父体）；2026.07 为子体级展开导出（3000 行、94 个父体）。

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { DatabaseSync } = require('node:sqlite');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const ROOT = path.resolve(__dirname, '..');
const RAW_DIR = path.resolve(ROOT, 'data/raw');
const DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const DEDUP_RULE = '每个Listing键（父ASIN优先，缺省ASIN）取原始表第一条记录（确定性，可复跑）';

function colNorm(name) {
  if (!name) return '';
  let value = String(name).trim();
  value = value.replace(/[\$\(\)'`]/g, '_').replace(/&/g, '_and_').replace(/\s+/g, '_');
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

function listingKey(row) {
  return String(row['父ASIN'] || '').trim() || String(row.ASIN || '').trim();
}

function quoteIdentifier(value) {
  return '"' + String(value).replace(/"/g, '""') + '"';
}

function main() {
  if (fs.existsSync(DB_PATH)) {
    const backup = path.join(ROOT, 'tmp', path.basename(DB_PATH) + '.old-' + new Date().toISOString().replace(/[:.]/g, '-'));
    fs.mkdirSync(path.dirname(backup), { recursive: true });
    fs.copyFileSync(DB_PATH, backup);
    console.log('Backed up existing DB -> ' + path.relative(ROOT, backup));
  }
  fs.rmSync(DB_PATH, { force: true });

  const files = fs.readdirSync(RAW_DIR)
    .filter((name) => /^Competitor-US-2026\.\d{2}-\d+\.xlsx$/i.test(name))
    .sort();
  if (files.length !== 7) throw new Error('Expected 7 competitor snapshots, found ' + files.length);

  const db = new DatabaseSync(DB_PATH);
  db.exec('PRAGMA journal_mode=WAL');
  db.exec('PRAGMA foreign_keys=OFF');
  db.exec('CREATE TABLE meta (id INTEGER PRIMARY KEY, built_at TEXT NOT NULL, source_dir TEXT NOT NULL, files TEXT NOT NULL, dedup_rule TEXT NOT NULL)');

  for (const file of files) {
    const month = '2026' + file.match(/2026\.(\d{2})/)[1];
    const workbook = XLSX.readFile(path.join(RAW_DIR, file), { cellDates: false, cellNF: false, cellFormula: false });
    const sheetName = workbook.SheetNames.find((name) => !/^Notes$/i.test(name));
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headers = [];
    for (let column = range.s.c; column <= range.e.c; column++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })];
      headers.push(cell ? cell.v : '');
    }
    const columns = dedupColumns(headers.map(colNorm));
    const rawTable = 'raw_' + month;
    const dedupTable = 'dedup_' + month;
    // NUMERIC 亲和列：数字保持数字（避免 node:sqlite 以 REAL 绑定写入 TEXT 亲和列后被转成 "591.0" 文本），
    // 文本列（标题/链接等）原样保留；与旧库"数字列存数字"的行为一致。
    db.exec('CREATE TABLE ' + quoteIdentifier(rawTable) + ' (row_id INTEGER PRIMARY KEY AUTOINCREMENT, ' +
      columns.map((column) => quoteIdentifier(column) + ' NUMERIC').join(', ') + ')');
    db.exec('CREATE TABLE ' + quoteIdentifier(dedupTable) + ' (row_id INTEGER PRIMARY KEY AUTOINCREMENT, ' +
      columns.map((column) => quoteIdentifier(column) + ' NUMERIC').join(', ') + ')');

    // raw:true 取单元格原始值（数字保持数字，避免格式化成 "591.0" 这类文本）；
    // 空字符串统一转 NULL（与竞品审计 equivalent() 约定一致：Excel 空值对应 DB NULL）。
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const insertRaw = db.prepare('INSERT INTO ' + quoteIdentifier(rawTable) + ' (' + columns.map(quoteIdentifier).join(', ') + ') VALUES (' + columns.map(() => '?').join(', ') + ')');
    const insertDedup = db.prepare('INSERT INTO ' + quoteIdentifier(dedupTable) + ' (' + columns.map(quoteIdentifier).join(', ') + ') VALUES (' + columns.map(() => '?').join(', ') + ')');
    const seen = new Set();
    let rawRows = 0;
    let dedupRows = 0;
    db.exec('BEGIN');
    // 逐行镜像源表（含全空行，全空行存为全 NULL 行，保证行数与 Excel 一致）
    for (let i = range.s.r + 1; i <= range.e.r; i++) {
      const row = rows[i] || [];
      const values = columns.map((_, column) => {
        const value = row[column];
        if (value === undefined || value === null || value === '') return null;
        return value;
      });
      insertRaw.run(...values);
      rawRows++;
      const key = String(values[columns.indexOf('父ASIN')] || '').trim() || String(values[columns.indexOf('ASIN')] || '').trim();
      if (key && !seen.has(key)) {
        seen.add(key);
        insertDedup.run(...values);
        dedupRows++;
      }
    }
    db.exec('COMMIT');
    const sha = crypto.createHash('sha256').update(fs.readFileSync(path.join(RAW_DIR, file))).digest('hex');
    db.prepare('INSERT INTO meta (built_at, source_dir, files, dedup_rule) VALUES (?, ?, ?, ?)')
      .run(new Date().toISOString(), path.relative(ROOT, RAW_DIR), file + '|sha256=' + sha, DEDUP_RULE);
    console.log(month + ': raw=' + rawRows + ' rows, dedup=' + dedupRows + ' rows (' + file + ')');
  }
  db.close();
  console.log('Built competitor DB: ' + path.relative(ROOT, DB_PATH));
}

main();
