'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const TMP_DIR = path.join(ROOT, 'tmp');
const TARGET = path.join(TMP_DIR, 'overwrite_guard.db');
fs.mkdirSync(TMP_DIR, { recursive: true });
fs.writeFileSync(TARGET, 'do-not-overwrite');

function hashFile(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

const before = hashFile(TARGET);
const result = spawnSync(process.execPath, ['src/import_xlsx.js'], {
  cwd: ROOT,
  encoding: 'utf8',
  env: {
    ...process.env,
    DATABASE_URL: 'sqlite:///tmp/overwrite_guard.db',
    IMPORT_OVERWRITE: 'false',
  },
});
const after = hashFile(TARGET);

const rejected = result.status !== 0 && /IMPORT_OVERWRITE=false/.test(result.stderr + result.stdout);
const unchanged = before === after;
console.log((rejected ? '[PASS]' : '[FAIL]') + ' existing DB is rejected when overwrite=false');
console.log((unchanged ? '[PASS]' : '[FAIL]') + ' existing DB hash is unchanged');
process.exit(rejected && unchanged ? 0 : 1);
