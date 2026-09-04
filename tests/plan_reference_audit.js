'use strict';

// Read-only audit for the plan-team reference workbook.  The workbook is a
// reference model, not an import source; formula cells are independently
// recomputed because the generated file has incomplete/empty cached results.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.resolve(__dirname, '..');
const PLAN_PATH = path.resolve(ROOT, '新增参考的材料和内容/销量预测计划部底表-户外地垫.xlsx');

let checks = 0;
let failures = 0;
function check(label, ok, detail) {
  checks++;
  if (!ok) failures++;
  console.log((ok ? '[PASS] ' : '[FAIL] ') + label + (detail ? ' (' + detail + ')' : ''));
}

function cell(sheet, address) {
  return sheet[address] || null;
}

function value(sheet, address) {
  const c = cell(sheet, address);
  return c && typeof c.v === 'number' && Number.isFinite(c.v) ? c.v : null;
}

function formula(sheet, address) {
  const c = cell(sheet, address);
  return c && typeof c.f === 'string' ? c.f : null;
}

function present(sheet, address) {
  const c = cell(sheet, address);
  return Boolean(c && c.v !== undefined && c.v !== null && c.v !== '');
}

function sumColumnRange(sheet, column, start, end) {
  let total = 0;
  for (let row = start; row <= end; row++) total += value(sheet, column + row) || 0;
  return total;
}

function sumRowRange(sheet, row, startColumn, endColumn) {
  const start = XLSX.utils.decode_col(startColumn);
  const end = XLSX.utils.decode_col(endColumn);
  let total = 0;
  for (let column = start; column <= end; column++) {
    total += value(sheet, XLSX.utils.encode_cell({ r: row - 1, c: column })) || 0;
  }
  return total;
}

function close(left, right, tolerance = 1e-9) {
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) <= tolerance;
}

if (!fs.existsSync(PLAN_PATH)) {
  console.error('FAIL: plan reference workbook not found: ' + PLAN_PATH);
  process.exit(1);
}

const workbook = XLSX.readFile(PLAN_PATH, {
  bookFiles: true,
  cellFormula: true,
  cellNF: true,
  cellStyles: true,
  sheetStubs: true,
});
const industry = workbook.Sheets['行业大盘数据'];
const plan = workbook.Sheets['基础销量预测-计划端'];
const bsr = workbook.Sheets['BSR底表-美国'];

// The workbook uses shared-formula XML.  The `xlsx` parser exposes the
// anchor formula but does not expand non-self-closing shared formula members
// in this file, so formula coverage is audited against the original OOXML.
function sheetXml(sheetName) {
  const meta = (workbook.Workbook.Sheets || []).find((item) => item.name === sheetName);
  const file = meta && workbook.files && workbook.files['xl/worksheets/sheet' + meta.sheetId + '.xml'];
  return file && Buffer.isBuffer(file.content) ? file.content.toString('utf8') : null;
}

function xmlCellInner(sheetName, address) {
  const xml = sheetXml(sheetName);
  if (!xml) return null;
  const escaped = address.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = xml.match(new RegExp('<c\\b[^>]*\\br="' + escaped + '"[^>]*>([\\s\\S]*?)</c>'));
  return match ? match[1] : null;
}

function xmlHasFormula(sheetName, address) {
  const inner = xmlCellInner(sheetName, address);
  return Boolean(inner && /<f\b/.test(inner));
}

function normalizedFormula(sheet, address) {
  const raw = formula(sheet, address);
  return raw ? raw.replace(/^=/, '').replace(/\s+/g, '').toUpperCase() : null;
}

check('reference workbook contains required sheets', Boolean(industry && plan && bsr));
check('reference workbook has 15 sheets', workbook.SheetNames.length === 15, 'sheets=' + workbook.SheetNames.length);

const market2025H1 = ['H4', 'H5', 'H6', 'H7', 'H8', 'H9'].reduce((sum, address) => sum + (value(industry, address) || 0), 0);
const market2026H1 = ['I4', 'I5', 'I6', 'I7', 'I8', 'I9'].reduce((sum, address) => sum + (value(industry, address) || 0), 0);
const marketGrowth = market2026H1 / market2025H1 - 1;
check('industry H1 inputs are present', market2025H1 === 1781765 && market2026H1 === 1831843,
  market2025H1 + '/' + market2026H1);
check('industry H1 2.8% formula independently recomputes', close(marketGrowth, 0.028105838873252), marketGrowth.toFixed(12));
check('industry D6 stores the independently recomputed result', close(value(industry, 'D6'), marketGrowth), 'D6=' + value(industry, 'D6'));
check('plan C3 agrees with industry D6', close(value(plan, 'C3'), marketGrowth), 'C3=' + value(plan, 'C3'));
check('industry monthly YoY formulas are present', ['J4', 'J5', 'J6', 'J7', 'J8', 'J9'].every((address) => xmlHasFormula('行业大盘数据', address))
  && normalizedFormula(industry, 'J4') === 'IFERROR(I4/H4-1,0)');
check('industry 2027 forecast formulas are present', ['K4', 'K5', 'K6', 'K7', 'K8', 'K9'].every((address) => xmlHasFormula('行业大盘数据', address))
  && normalizedFormula(industry, 'K4') === '2.8%+0.3*(J4-2.8%)');

const c2 = value(plan, 'C2');
const c3 = value(plan, 'C3');
const d3 = (c3 - c2) / c2;
const d4 = c3 * (1 + d3);
check('plan D3 formula is present', formula(plan, 'D3') === '(C3-C2)/C2');
check('plan D4/C4 2027 baseline independently recomputes', formula(plan, 'D4') === 'C3*(1+D3)' && formula(plan, 'C4') === 'D4'
  && close(d4, 0.015691748432596353), d4.toFixed(12));

const topRows = [];
for (let row = 2; row <= 101; row++) {
  const rank = value(bsr, 'Q' + row);
  if (rank !== null && rank >= 1 && rank <= 100) {
    topRows.push({ row, rank,
      h1_2025: sumRowRange(bsr, row, 'AX', 'BC'),
      h1_2026: sumRowRange(bsr, row, 'BJ', 'BO'),
      workbookP: sumRowRange(bsr, row, 'BJ', 'BN') + (value(bsr, 'BP' + row) || 0),
    });
  }
}
const bsr2025H1 = topRows.reduce((sum, item) => sum + item.h1_2025, 0);
const bsr2026H1 = topRows.reduce((sum, item) => sum + item.h1_2026, 0);
const workbookP2026 = topRows.reduce((sum, item) => sum + item.workbookP, 0);
check('BSR Q=1..100 contains exactly 100 ranked rows', topRows.length === 100, 'rows=' + topRows.length);
check('BSR 2025H1 independent sum is 1,066,818', bsr2025H1 === 1066818, 'sum=' + bsr2025H1);
check('BSR 2026H1 independent sum is 1,121,130', bsr2026H1 === 1121130, 'sum=' + bsr2026H1);
check('BSR H1 growth independently recomputes +5.0910%', close(bsr2026H1 / bsr2025H1 - 1, 0.05091027710443585), (bsr2026H1 / bsr2025H1 - 1).toFixed(12));
check('BSR P formula defect is detected', formula(bsr, 'P2') === 'SUM(BJ2:BN2)+BP2', 'P2=' + formula(bsr, 'P2'));
check('BSR defective P formula produces a non-authoritative result', workbookP2026 === 1018734 && !close(workbookP2026, bsr2026H1), 'P=' + workbookP2026);
const missingH1Rows = topRows.filter((item) => !xmlHasFormula('BSR底表-美国', 'M' + item.row) || !xmlHasFormula('BSR底表-美国', 'N' + item.row)
  || !xmlHasFormula('BSR底表-美国', 'O' + item.row) || !xmlHasFormula('BSR底表-美国', 'P' + item.row));
check('BSR incomplete M:P coverage is detected', missingH1Rows.length === 22, 'missingRows=' + missingH1Rows.length + ' ranks=' + missingH1Rows.map((item) => item.rank).join(','));
const emptyAnnualInputs = ['B', 'C', 'D', 'E'].every((column) => topRows.every((item) => value(bsr, column + item.row) === null));
check('BSR F:H source columns B:E are empty and therefore unsafe', emptyAnnualInputs);
const missingGroupKeys = topRows.filter((item) => !present(bsr, 'L' + item.row)).length;
const missingWideGroups = topRows.filter((item) => !present(bsr, 'R' + item.row)).length;
check('BSR group-key gaps are detected', missingGroupKeys === 45 && missingWideGroups === 2,
  'L_missing=' + missingGroupKeys + ' R_missing=' + missingWideGroups);

console.log(`Plan reference audit: ${checks} checks, ${failures} failures`);
process.exitCode = failures ? 1 : 0;
