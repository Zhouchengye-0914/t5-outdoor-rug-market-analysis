'use strict';

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');
const MD_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.md');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const DB_PATH = path.resolve(ROOT, process.env.ANALYSIS_DB_PATH || 'data/processed/market.db');

let checks = 0;
let failures = 0;
function check(label, ok, detail) {
  checks++;
  if (!ok) failures++;
  console.log((ok ? '[PASS] ' : '[FAIL] ') + label + (detail ? ' (' + detail + ')' : ''));
}

function close(left, right, tolerance = 1e-8) {
  if (left === null || right === null) return left === right;
  return Number.isFinite(Number(left)) && Number.isFinite(Number(right))
    && Math.abs(Number(left) - Number(right)) <= tolerance;
}

function pct(current, previous) {
  return previous ? (current / previous - 1) * 100 : null;
}

function byMonth(rows) {
  return new Map(rows.map((row) => [row.month, row]));
}

const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
const markdown = fs.readFileSync(MD_PATH, 'utf8');
const html = fs.readFileSync(HTML_PATH, 'utf8');
const db = new DatabaseSync(DB_PATH, { readOnly: true });

check('core cutoff ends at 202606', data.analysisMonths.at(-1) === '202606');
check('202607 is source-only appendix', data.sourceMonths.includes('202607')
  && !data.analysisMonths.includes('202607')
  && data.excludedFromComparableReport.includes('202607'));
check('49 core analysis months', data.analysisMonths.length === 49, 'months=' + data.analysisMonths.length);
check('seven replacement metadata rows', data.replacementMetadata.length === 7, 'rows=' + data.replacementMetadata.length);
check('best-BSR enrichment covers 202601-202607', data.dataQuality.competitorDatabaseAvailable
  && data.dataQuality.bestBsrEnrichedMonths.length === 7);

const catalog = db.prepare("SELECT target_table FROM sheet_catalog WHERE target_table IS NOT NULL").all();
let actualRows = 0;
let actualCells = 0;
for (const { target_table: table } of catalog) {
  const rows = db.prepare('SELECT COUNT(*) AS count FROM ' + table).get().count;
  const columns = db.prepare('PRAGMA table_info(' + table + ')').all()
    .filter((column) => !['row_id', 'month_label'].includes(column.name)).length;
  actualRows += rows;
  actualCells += rows * columns;
}
check('reported actual row count matches DB', data.currentDataRowCount === actualRows,
  `json=${data.currentDataRowCount} db=${actualRows}`);
check('reported verified cell count matches DB', data.verifiedDataCellCount === actualCells,
  `json=${data.verifiedDataCellCount} db=${actualCells}`);

for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const current = data.categories[category];
  check(category + ' monthly covers core range', current.monthly.length === data.analysisMonths.length
    && current.monthly[0].month === data.analysisMonths[0]
    && current.monthly.at(-1).month === data.analysisMonths.at(-1));

  for (const collectionName of ['monthly', 'bsrTop100']) {
    const rows = collectionName === 'monthly' ? current.monthly : current.bsrTop100.monthly;
    const index = byMonth(rows);
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      const previous = rows[rowIndex - 1];
      const priorMonth = String(Number(row.month.slice(0, 4)) - 1) + row.month.slice(4);
      const prior = index.get(priorMonth);
      if (prior) {
        if (row.momBasis !== priorMonth || !close(row.momSales, pct(row.sales, prior.sales))
          || !close(row.momRevenue, pct(row.revenue, prior.revenue))) {
          check(category + ' ' + collectionName + ' MOM formulas', false, 'month=' + row.month);
          break;
        }
      } else if (row.momBasis !== null || row.momSales !== null || row.momRevenue !== null) {
        check(category + ' ' + collectionName + ' missing MOM basis remains null', false, 'month=' + row.month);
        break;
      }
      if (previous) {
        if (row.chainBasis !== previous.month || !close(row.chainSales, pct(row.sales, previous.sales))
          || !close(row.chainRevenue, pct(row.revenue, previous.revenue))) {
          check(category + ' ' + collectionName + ' chain formulas', false, 'month=' + row.month);
          break;
        }
      }
      if (rowIndex === rows.length - 1) {
        check(category + ' ' + collectionName + ' MOM/chain formulas', true);
      }
    }
  }

  const topByMonth = byMonth(current.bsrTop100.monthly);
  const groupsByMonth = new Map();
  const segmentsByMonth = new Map();
  for (const row of current.bsrGroups.monthly) {
    if (!groupsByMonth.has(row.month)) groupsByMonth.set(row.month, []);
    groupsByMonth.get(row.month).push(row);
  }
  for (const row of current.bsrSegments.monthly) {
    if (!segmentsByMonth.has(row.month)) segmentsByMonth.set(row.month, []);
    segmentsByMonth.get(row.month).push(row);
  }
  let reconciled = true;
  for (const month of data.analysisMonths) {
    const top = topByMonth.get(month);
    const groups = groupsByMonth.get(month) || [];
    const segments = segmentsByMonth.get(month) || [];
    const totals = (rows) => ({
      count: rows.reduce((sum, row) => sum + row.skuCount, 0),
      sales: rows.reduce((sum, row) => sum + row.sales, 0),
      revenue: rows.reduce((sum, row) => sum + row.revenue, 0),
    });
    const groupTotals = totals(groups);
    const segmentTotals = totals(segments);
    if (!top || top.skuCount > 100 || groups.length !== 3 || segments.length !== 5
      || groupTotals.count !== top.skuCount || segmentTotals.count !== top.skuCount
      || !close(groupTotals.sales, top.sales) || !close(segmentTotals.sales, top.sales)
      || !close(groupTotals.revenue, top.revenue) || !close(segmentTotals.revenue, top.revenue)) {
      reconciled = false;
      break;
    }
  }
  check(category + ' Top100 cap and tier reconciliation', reconciled);

  const annual2026 = current.annual.find((row) => row.year === '2026');
  check(category + ' 2026 annual period is Jan-Jun with scope warning', annual2026.period === '202601-202606'
    && annual2026.comparison === '202601-202606 vs 202501-202506'
    && annual2026.scopeComparable === false && Boolean(annual2026.scopeNote));
}

const overall = byMonth(data.categories.overall.monthly);
const pp = byMonth(data.categories.pp.monthly);
const high = byMonth(data.categories.high.monthly);
let partitioned = true;
for (const month of data.analysisMonths) {
  const total = overall.get(month);
  const plastic = pp.get(month);
  const nonPlastic = high.get(month);
  if (total.skuCount !== plastic.skuCount + nonPlastic.skuCount
    || !close(total.sales, plastic.sales + nonPlastic.sales)
    || !close(total.revenue, plastic.revenue + nonPlastic.revenue)) {
    partitioned = false;
    break;
  }
}
check('PP and high form an exact overall partition', partitioned);

const benchmark = overall.get('202602');
check('202602 overall benchmark MOM sales ≈ -27.7%', close(benchmark.momSales, -27.7, 0.1), benchmark.momSales.toFixed(3));
check('202602 overall benchmark MOM revenue ≈ -0.1%', close(benchmark.momRevenue, -0.1, 0.1), benchmark.momRevenue.toFixed(3));
check('202602 overall benchmark chain sales ≈ +9.3%', close(benchmark.chainSales, 9.3, 0.1), benchmark.chainSales.toFixed(3));
check('202602 overall benchmark chain revenue ≈ +33.5%', close(benchmark.chainRevenue, 33.5, 0.1), benchmark.chainRevenue.toFixed(3));

let ppSales2025 = 0;
let ppRevenue2025 = 0;
let genimoPpSales2025 = 0;
let genimoPpRevenue2025 = 0;
for (const month of data.analysisMonths.filter((value) => value.startsWith('2025'))) {
  const rows = db.prepare('SELECT 品牌 brand, 商品标题 title, 月销量 sales, 月销售额 revenue FROM monthly_' + month).all();
  for (const row of rows) {
    if (!/\bplastic\b/i.test(String(row.title || ''))) continue;
    ppSales2025 += Number(row.sales || 0);
    ppRevenue2025 += Number(row.revenue || 0);
    if (String(row.brand || '').trim().toLowerCase() === 'genimo') {
      genimoPpSales2025 += Number(row.sales || 0);
      genimoPpRevenue2025 += Number(row.revenue || 0);
    }
  }
}
check('GENIMO PP sales share uses PP-only numerator', close(data.insights.genimoPpShare2025,
  genimoPpSales2025 / ppSales2025 * 100));
check('GENIMO PP revenue share uses PP-only numerator', close(data.insights.genimoPpRevenueShare2025,
  genimoPpRevenue2025 / ppRevenue2025 * 100));

for (const [name, output] of [['Markdown', markdown], ['HTML', html]]) {
  check(name + ' has no stale high-price classification', !output.includes('材质关键词或价格≥$40'));
  check(name + ' has no random representative wording', !output.includes('随机保留'));
  check(name + ' contains exact 90-day exit gate', output.includes('连续 90 天无法进入前100') || output.includes('连续90天无法进入前100'));
  check(name + ' contains cross-scope warning', output.includes('2026竞品父体口径与2025全市场SKU口径不同')
    || output.includes('2026与2025数据范围不同'));
}
check('HTML has exactly one coverage navigation entry', (html.match(/href="#coverage"/g) || []).length === 1);

console.log('\n========== ANALYSIS AUDIT ==========');
console.log('Checks: ' + checks);
console.log('Failures: ' + failures);
db.close();
process.exit(failures > 0 ? 1 : 0);
