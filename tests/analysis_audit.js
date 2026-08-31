'use strict';

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');

const ROOT = path.resolve(__dirname, '..');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');
const MD_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.md');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const QUICK_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-极速版.md');
const DB_PATH = path.resolve(ROOT, process.env.ANALYSIS_DB_PATH || 'data/processed/market.db');
const COMPETITOR_DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');

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
const quick = fs.readFileSync(QUICK_PATH, 'utf8');
const db = new DatabaseSync(DB_PATH, { readOnly: true });
const competitorDb = new DatabaseSync(COMPETITOR_DB_PATH, { readOnly: true });

check('core cutoff ends at 202606', data.analysisMonths.at(-1) === '202606');
check('202607 is merged display but excluded from comparable core', data.sourceMonths.includes('202607')
  && !data.analysisMonths.includes('202607')
  && data.excludedFromComparableReport.includes('202607'));
check('49 core analysis months', data.analysisMonths.length === 49, 'months=' + data.analysisMonths.length);
check('seven replacement metadata rows', data.replacementMetadata.length === 7, 'rows=' + data.replacementMetadata.length);
check('best-BSR enrichment covers 202601-202607', data.dataQuality.competitorDatabaseAvailable
  && data.dataQuality.bestBsrEnrichedMonths.length === 7);
check('historical BSR quality diagnostics cover all 43 pre-2026 months',
  data.dataQuality.historicalBsrTop100Quality.length === 43
  && data.dataQuality.historicalBsrTop100Quality.every((row) => row.month < '202601'
    && row.identifierCoveragePct === 0 && row.strictListingPool === false));
check('merged 2026 overall trend covers Jan-Jul in order', data.overallMarketTrend2026.length === 7
  && data.overallMarketTrend2026.map((row) => row.month).join(',') === '202601,202602,202603,202604,202605,202606,202607');

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
  check(category + ' 2026 annual period is Jan-Jun and explicitly directional only', annual2026.period === '202601-202606'
    && annual2026.comparison === '202601-202606 vs 202501-202506'
    && annual2026.timeComparable === true && annual2026.scopeComparable === false
    && annual2026.scopeNote.includes('不构成严格同口径同比'));
}

let priceAveragesReconcile = true;
let missingPriceRows = 0;
for (const month of data.analysisMonths) {
  const prices = db.prepare('SELECT 价格 price FROM monthly_' + month).all();
  const numeric = prices.filter((row) => row.price !== null && row.price !== undefined && row.price !== ''
    && Number.isFinite(Number(row.price))).map((row) => Number(row.price));
  missingPriceRows += prices.length - numeric.length;
  const expectedAverage = numeric.length ? numeric.reduce((sum, value) => sum + value, 0) / numeric.length : null;
  const actual = data.categories.overall.monthly.find((row) => row.month === month);
  if (!actual || actual.pricedSkuCount !== numeric.length || !close(actual.avgListPrice, expectedAverage)) {
    priceAveragesReconcile = false;
    break;
  }
}
check('overall average list price excludes missing/non-numeric prices and reconciles to DB',
  priceAveragesReconcile && missingPriceRows > 0, 'missingExcluded=' + missingPriceRows);

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

let mergedTrendReconciles = true;
for (const row of data.overallMarketTrend2026.filter((item) => item.month <= '202606')) {
  const source = overall.get(row.month);
  if (!source || !row.coreComparable || row.skuCount !== source.skuCount
    || !close(row.sales, source.sales) || !close(row.revenue, source.revenue)
    || !close(row.avgListPrice, source.avgListPrice) || !close(row.weightedPrice, source.weightedPrice)
    || !close(row.momSales, source.momSales) || !close(row.momRevenue, source.momRevenue)
    || !close(row.chainSales, source.chainSales) || !close(row.chainRevenue, source.chainRevenue)) {
    mergedTrendReconciles = false;
    break;
  }
}
check('merged Jan-Jun trend reconciles to core overall monthly data', mergedTrendReconciles);
const julyTrend = data.overallMarketTrend2026.find((row) => row.month === '202607');
const julyDiagnostic = data.sourceDiagnostics.find((row) => row.month === '202607');
check('merged July trend reconciles to source and blocks cross-caliber deltas', julyTrend && julyDiagnostic
  && julyTrend.coreComparable === false && julyTrend.skuCount === 94
  && close(julyTrend.sales, julyDiagnostic.sales) && close(julyTrend.revenue, julyDiagnostic.revenue)
  && julyTrend.momBasis === null && julyTrend.chainBasis === null
  && julyTrend.momSales === null && julyTrend.momRevenue === null
  && julyTrend.chainSales === null && julyTrend.chainRevenue === null
  && julyTrend.scopeStatus.includes('不参与同比/环比和累计'));
for (const [name, output] of [['Markdown', markdown], ['HTML', html]]) {
  check(name + ' merges Jan-Jul trend and has no separate July appendix', output.includes('2026.01-07整体市场趋势（合并展示）')
    && !output.includes('2026.07附录/参考') && !output.includes('参考附录'));
}
check('极速版包含核心结论与所有关键限制', quick.includes('户外地垫市场分析报告（极速版）')
  && quick.includes('2026.07仅94父体') && quick.includes('不是严格同口径同比')
  && quick.includes('2022-2025 BSR Top100是源表行代理') && quick.includes('SKU平均标价排除空值'));

// 无对应数据（基准分层为空）披露校验：2025.05 源数据小类BSR重复 → 2026.05 中部/尾部 MOM 无基准
const overallGroups = data.categories.overall.bsrGroups.monthly;
const mayGap = overallGroups.filter((r) => r.month === '202605' && (r.segment === '中部（21-50）' || r.segment === '尾部（51-100）'));
check('202605 mid/tail MOM basis=202505 is null with gap reason',
  mayGap.length === 2
  && mayGap.every((r) => r.momBasis === '202505' && r.momSales === null && r.momRevenue === null
    && Boolean(r.momGapReason) && r.momGapReason.includes('202505')));
check('202605 mid/tail chain (vs 202604) still computed',
  mayGap.every((r) => r.chainBasis === '202604' && r.chainSales !== null && r.chainRevenue !== null && !r.chainGapReason));
check('gap reason only when basis segment is empty',
  overallGroups.filter((r) => r.momGapReason && (r.momSales !== null || r.momRevenue !== null)).length === 0
  && overallGroups.filter((r) => !r.momGapReason && r.momBasis && r.momSales === null).length === 0);
for (const [name, output] of [['Markdown', markdown], ['HTML', html]]) {
  check(name + ' discloses 2025.05 BSR duplication and no-data cells', output.includes('小类BSR同值重复') && output.includes('无对应数据'));
}

const may2026 = overall.get('202605');
check('202605 overall uses new full-market snapshot (1690 listings)', may2026.skuCount === 1690, 'sku=' + may2026.skuCount);
check('202601 overall uses new full-market snapshot (1134 listings)', overall.get('202601').skuCount === 1134, 'sku=' + overall.get('202601').skuCount);

const benchmark = overall.get('202602');
check('202602 overall benchmark MOM sales ≈ -14.8%', close(benchmark.momSales, -14.8, 0.1), benchmark.momSales.toFixed(3));
check('202602 overall benchmark MOM revenue ≈ -20.5%', close(benchmark.momRevenue, -20.5, 0.1), benchmark.momRevenue.toFixed(3));
check('202602 overall benchmark chain sales ≈ +9.6%', close(benchmark.chainSales, 9.6, 0.1), benchmark.chainSales.toFixed(3));
check('202602 overall benchmark chain revenue ≈ +23.8%', close(benchmark.chainRevenue, 23.8, 0.1), benchmark.chainRevenue.toFixed(3));

function listingKey(row) {
  return String(row.parent || '').trim() || String(row.asin || '').trim() || 'row-' + row.row_id;
}
let ppSales2026 = 0;
let ppRevenue2026 = 0;
let genimoPpSales2026 = 0;
let genimoPpRevenue2026 = 0;
for (const month of data.analysisMonths.filter((value) => value >= '202601' && value <= '202606')) {
  const profiles = new Map();
  const rawRows = competitorDb.prepare('SELECT row_id, ASIN asin, "父ASIN" parent, 品牌 brand, 商品标题 title FROM raw_' + month).all();
  for (const row of rawRows) {
    const key = listingKey(row);
    const profile = profiles.get(key) || { plastic: false, genimo: false };
    profile.plastic = profile.plastic || /\bplastic\b/i.test(String(row.title || ''));
    profile.genimo = profile.genimo || String(row.brand || '').trim().toLowerCase() === 'genimo';
    profiles.set(key, profile);
  }
  const rows = db.prepare('SELECT row_id, ASIN asin, "父ASIN" parent, 月销量 sales, 月销售额 revenue FROM monthly_' + month).all();
  for (const row of rows) {
    const profile = profiles.get(listingKey(row));
    if (!profile || !profile.plastic) continue;
    ppSales2026 += Number(row.sales || 0);
    ppRevenue2026 += Number(row.revenue || 0);
    if (profile.genimo) {
      genimoPpSales2026 += Number(row.sales || 0);
      genimoPpRevenue2026 += Number(row.revenue || 0);
    }
  }
}
check('2026 GENIMO PP sales share uses family-aware PP-only numerator', close(data.insights.genimoPpShare2026,
  genimoPpSales2026 / ppSales2026 * 100));
check('2026 GENIMO PP revenue share uses family-aware PP-only numerator', close(data.insights.genimoPpRevenueShare2026,
  genimoPpRevenue2026 / ppRevenue2026 * 100));
check('insight payload is 2026-only and omits legacy 2025 cards', data.insights.period === '202601-202606'
  && data.insights.overall2026 && data.insights.pp2026 && data.insights.high2026 && data.insights.genimo2026
  && !Object.prototype.hasOwnProperty.call(data.insights, 'overall2025')
  && !Object.prototype.hasOwnProperty.call(data.insights, 'genimoPpShare2025'));
check('GENIMO Top products are restricted to 2026 Jan-Jun', data.genimoTopProductsPeriod === '202601-202606'
  && data.genimoTopProducts.length > 0
  && data.genimoTopProducts.every((row) => row.months >= 1 && row.months <= 6 && row.parent && row.listingKey));

for (const [name, output] of [['Markdown', markdown], ['HTML', html]]) {
  check(name + ' has no stale high-price classification', !output.includes('材质关键词或价格≥$40'));
  check(name + ' has no random representative wording', !output.includes('随机保留'));
  check(name + ' contains exact 90-day exit gate', output.includes('连续 90 天无法进入前100') || output.includes('连续90天无法进入前100'));
  check(name + ' explicitly blocks strict 2025→2026同比 interpretation', output.includes('不是严格同口径同比')
    || output.includes('不构成严格同口径同比') || output.includes('跨年变化仅作方向性参考'));
  check(name + ' trend and GENIMO recommendations use 2026 actuals', output.includes('趋势结论与GENIMO建议（基于2026.01-06实绩）')
    && output.includes('GENIMO 2026.01-06累计Top父体')
    && !output.includes('2025年PP销量份额'));
  check(name + ' has no obsolete 64-94 monthly caliber claim', !output.includes('2026年数据源已切换为竞品父ASIN去重口径（64-94父商品/月）'));
}
check('forecast includes all four required 2026 Q4 ranges', ['104,000-125,000', '89,000-107,000', '85,000-102,000', '91,000-110,000']
  .every((range) => markdown.includes(range) && html.includes(range)));
check('HTML includes three accessible static SVG trend charts', (html.match(/<figure class="chart-card">/g) || []).length === 3
  && (html.match(/<svg /g) || []).length === 3 && (html.match(/role="img"/g) || []).length === 3);
check('HTML exposes 43-row historical BSR quality table', html.includes('2022-2025 BSR Top100逐月质量诊断')
  && html.includes('源表行代理（无Listing标识）'));
check('HTML has exactly one coverage navigation entry', (html.match(/href="#coverage"/g) || []).length === 1);

console.log('\n========== ANALYSIS AUDIT ==========');
console.log('Checks: ' + checks);
console.log('Failures: ' + failures);
db.close();
competitorDb.close();
process.exit(failures > 0 ? 1 : 0);
