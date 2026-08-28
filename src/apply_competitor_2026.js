'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const ROOT = path.resolve(__dirname, '..');
const TARGET_DB = path.resolve(
  ROOT,
  (process.env.DATABASE_URL || 'sqlite:///data/processed/market.db').replace(/^sqlite:\/\/\/?/, ''),
);
const COMPETITOR_DB = path.resolve(
  ROOT,
  process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db',
);
// 2026.01-06 已更新为全市场 listing 级导出（每父体一行代表子体）；2026.07 仍为子体级展开（94 父体）。
const EXPECTED_ROWS = {
  '202601': 1134,
  '202602': 1038,
  '202603': 1766,
  '202604': 1744,
  '202605': 1690,
  '202606': 1993,
  '202607': 94,
};

function quoteIdentifier(value) {
  return '"' + String(value).replace(/"/g, '""') + '"';
}

function tableExists(db, tableName) {
  return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(tableName));
}

function listingKey(row) {
  return String(row['父ASIN'] || '').trim() || String(row.ASIN || '').trim();
}

function sha256(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function main() {
  if (!fs.existsSync(TARGET_DB)) throw new Error('Target market DB not found: ' + TARGET_DB);
  if (!fs.existsSync(COMPETITOR_DB)) throw new Error('Competitor DB not found: ' + COMPETITOR_DB);

  const target = new DatabaseSync(TARGET_DB);
  const source = new DatabaseSync(COMPETITOR_DB, { readOnly: true });
  const sourceHash = sha256(COMPETITOR_DB);
  const appliedAt = new Date().toISOString();

  target.exec(`CREATE TABLE IF NOT EXISTS analysis_replacements (
    month TEXT PRIMARY KEY,
    target_table TEXT NOT NULL,
    source_database TEXT NOT NULL,
    source_table TEXT NOT NULL,
    source_sha256 TEXT NOT NULL,
    base_imported_rows INTEGER NOT NULL,
    source_raw_rows INTEGER NOT NULL,
    replacement_rows INTEGER NOT NULL,
    selection_rule TEXT NOT NULL,
    applied_at TEXT NOT NULL
  )`);

  try {
    target.exec('BEGIN IMMEDIATE');
    for (const [month, expectedRows] of Object.entries(EXPECTED_ROWS)) {
      const targetTable = 'monthly_' + month;
      const sourceTable = 'dedup_' + month;
      const rawTable = 'raw_' + month;
      if (!tableExists(target, targetTable)) throw new Error('Missing target table: ' + targetTable);
      if (!tableExists(source, sourceTable)) throw new Error('Missing source table: ' + sourceTable);
      if (!tableExists(source, rawTable)) throw new Error('Missing raw audit table: ' + rawTable);

      const targetColumns = target.prepare('PRAGMA table_info(' + quoteIdentifier(targetTable) + ')').all()
        .map((column) => column.name)
        .filter((name) => name !== 'row_id');
      const sourceColumns = new Set(source.prepare('PRAGMA table_info(' + quoteIdentifier(sourceTable) + ')').all()
        .map((column) => column.name));
      const sourceRows = source.prepare('SELECT * FROM ' + quoteIdentifier(sourceTable) + ' ORDER BY row_id').all();
      const rawRows = source.prepare('SELECT COUNT(*) AS count FROM ' + quoteIdentifier(rawTable)).get().count;
      if (sourceRows.length !== expectedRows) {
        throw new Error(`${sourceTable} row count ${sourceRows.length}, expected ${expectedRows}`);
      }
      const keys = sourceRows.map(listingKey);
      if (keys.some((key) => !key) || new Set(keys).size !== keys.length) {
        throw new Error(sourceTable + ' contains blank or duplicate parent/ASIN listing keys');
      }

      const previous = target.prepare('SELECT base_imported_rows FROM analysis_replacements WHERE month=?').get(month);
      const catalog = target.prepare('SELECT imported_rows FROM sheet_catalog WHERE target_table=?').get(targetTable);
      if (!catalog) throw new Error('sheet_catalog missing target table: ' + targetTable);
      const baseImportedRows = previous ? previous.base_imported_rows : catalog.imported_rows;

      target.exec('DELETE FROM ' + quoteIdentifier(targetTable));
      if (tableExists(target, 'sqlite_sequence')) {
        target.prepare('DELETE FROM sqlite_sequence WHERE name=?').run(targetTable);
      }
      const columnSql = targetColumns.map(quoteIdentifier).join(', ');
      const placeholders = targetColumns.map(() => '?').join(', ');
      const insert = target.prepare(
        'INSERT INTO ' + quoteIdentifier(targetTable) + ' (' + columnSql + ') VALUES (' + placeholders + ')',
      );
      for (const row of sourceRows) {
        insert.run(...targetColumns.map((column) => {
          if (column === 'month_label') return month;
          return sourceColumns.has(column) && row[column] !== undefined ? row[column] : null;
        }));
      }

      target.prepare('UPDATE sheet_catalog SET imported_rows=? WHERE target_table=?')
        .run(sourceRows.length, targetTable);
      target.prepare(`INSERT INTO analysis_replacements
        (month, target_table, source_database, source_table, source_sha256,
         base_imported_rows, source_raw_rows, replacement_rows, selection_rule, applied_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(month) DO UPDATE SET
          target_table=excluded.target_table,
          source_database=excluded.source_database,
          source_table=excluded.source_table,
          source_sha256=excluded.source_sha256,
          source_raw_rows=excluded.source_raw_rows,
          replacement_rows=excluded.replacement_rows,
          selection_rule=excluded.selection_rule,
          applied_at=excluded.applied_at`)
        .run(
          month,
          targetTable,
          path.basename(COMPETITOR_DB),
          sourceTable,
          sourceHash,
          baseImportedRows,
          rawRows,
          sourceRows.length,
          'canonical dedup snapshot; one representative row per parent/ASIN; analysis derives best category BSR from raw variants',
          appliedAt,
        );
      console.log(`[APPLY] ${sourceTable} -> ${targetTable}: ${sourceRows.length} rows`);
    }
    target.exec('COMMIT');
    target.exec('PRAGMA wal_checkpoint(TRUNCATE)');
  } catch (error) {
    try { target.exec('ROLLBACK'); } catch (_) { /* no-op */ }
    throw error;
  } finally {
    source.close();
    target.close();
  }
  console.log('[DONE] 2026.01-07 canonical competitor replacements applied reproducibly.');
}

try {
  main();
} catch (error) {
  console.error('FATAL:', error.message);
  process.exit(1);
}
