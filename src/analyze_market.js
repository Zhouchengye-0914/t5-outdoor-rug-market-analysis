'use strict';

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const ROOT = path.resolve(__dirname, '..');
const DB_PATH = path.resolve(ROOT, process.env.ANALYSIS_DB_PATH || 'data/processed/market.db');
const COMPETITOR_DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const REPORT_CUTOFF = process.env.ANALYSIS_CUTOFF || '202606'; // SPEC 1.2: 核心截止 202606
const MD_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.md');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');

// MATERIAL_KEYWORDS removed per SPEC 7.5 (high = all non-PP products)
const SEGMENTS = [
  { key: '1-5', min: 1, max: 5 },
  { key: '6-10', min: 6, max: 10 },
  { key: '11-20', min: 11, max: 20 },
  { key: '21-50', min: 21, max: 50 },
  { key: '51-100', min: 51, max: 100 },
];
const SEGMENT_GROUPS = [
  { key: '头部（1-20）', min: 1, max: 20 },
  { key: '中部（21-50）', min: 21, max: 50 },
  { key: '尾部（51-100）', min: 51, max: 100 },
];

function pct(current, previous) {
  if (!Number.isFinite(previous) || previous === 0) return null;
  return (current / previous - 1) * 100;
}

function fmt(value, digits = 0) {
  if (value === null || value === undefined || !Number.isFinite(Number(value))) return '-';
  return Number(value).toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

function fmtPct(value) {
  if (value === null || value === undefined || !Number.isFinite(value)) return '-';
  return (value >= 0 ? '+' : '') + value.toFixed(1) + '%';
}

function parseBsr(value) {
  if (value === null || value === undefined || value === '') return { rank: null, multi: false };
  const matches = String(value).match(/\d[\d,]*/g) || [];
  const ranks = matches.map((s) => Number(s.replace(/,/g, ''))).filter(Number.isFinite);
  return { rank: ranks.length ? Math.min(...ranks) : null, multi: ranks.length > 1 };
}

// SPEC 7.5: PP = 标题完整单词 plastic（单词边界，不区分大小写）；high = 排除 PP 后全部产品
const PLASTIC_WORD_RE = /\bplastic\b/i;
// 2026.09-12 预测和 2027 规划基准（SPEC 1.3，仅作预测/假设参考，非历史实绩）
const FORECAST_2026_Q4 = [
  { month: '202609', sales: 115000, rev: 3270000, stage: '旺季结束、需求快速回落' },
  { month: '202610', sales: 99000, rev: 3160000, stage: '淡季+秋季促销' },
  { month: '202611', sales: 94000, rev: 3870000, stage: '黑五带来销售额修复' },
  { month: '202612', sales: 101000, rev: 4120000, stage: '低基数+节日场景支撑' },
];
const FORECAST_2027_MONTHLY = [
  { month: '202701', sales: 40500, rev: 1370000, note: '全年低点' },
  { month: '202702', sales: 57800, rev: 2030000, note: '开始预热' },
  { month: '202703', sales: 130600, rev: 5140000, note: '需求快速启动' },
  { month: '202704', sales: 192500, rev: 7480000, note: '旺季增长' },
  { month: '202705', sales: 252300, rev: 8150000, note: '旺季加速' },
  { month: '202706', sales: 300500, rev: 10030000, note: '全年峰值' },
  { month: '202707', sales: 225600, rev: 8960000, note: '旺季转折' },
  { month: '202708', sales: 168600, rev: 4960000, note: '快速回落' },
  { month: '202709', sales: 120000, rev: 3470000, note: '进入淡季' },
  { month: '202710', sales: 102900, rev: 3350000, note: '淡季' },
  { month: '202711', sales: 97600, rev: 4100000, note: '黑五支撑销售额' },
  { month: '202712', sales: 105000, rev: 4370000, note: '小幅修复' },
];
const FORECAST_2027_SCENARIOS = [
  { scenario: '保守', sales: '168万-172万', rev: '5700万-5900万美元', trigger: '消费疲软、价格战持续、RV需求下降' },
  { scenario: '基准', sales: '177万-181万', rev: '6200万-6500万美元', trigger: '户外需求稳定、价格逐步企稳' },
  { scenario: '进取', sales: '188万-192万', rev: '6800万-7000万美元', trigger: '春夏天气有利、头部品牌减少价格战、Amazon大促表现良好' },
];
function classify(row, category) {
  const title = String(row.title || '');
  const isPlastic = typeof row.familyHasPlastic === 'boolean'
    ? row.familyHasPlastic
    : PLASTIC_WORD_RE.test(title);
  const isGenimo = typeof row.familyHasGenimo === 'boolean'
    ? row.familyHasGenimo
    : String(row.brand || '').trim().toLowerCase() === 'genimo';
  if (category === 'overall') return true;
  if (category === 'pp') return isPlastic;
  if (category === 'high') return !isPlastic;
  if (category === 'genimo') return isGenimo;
  return false;
}

function listingKey(row) {
  const parent = String(row.parent || row['父ASIN'] || '').trim();
  if (parent) return parent;
  const asin = String(row.asin || row.ASIN || '').trim();
  if (asin) return asin;
  return 'row-' + row.row_id;
}

function rankForCategory(row, category) {
  if (!row.bestRanks) return row.rank;
  if (category === 'pp') return row.bestRanks.pp ?? row.rank;
  if (category === 'high') return row.bestRanks.high ?? row.rank;
  if (category === 'genimo') return row.bestRanks.genimo ?? row.rank;
  return row.bestRanks.overall ?? row.rank;
}

function top100Rows(rows) {
  return rows.filter((row) => row.rank !== null && row.rank >= 1 && row.rank <= 100)
    .sort((left, right) => left.rank - right.rank
      || listingKey(left).localeCompare(listingKey(right))
      || Number(left.row_id || 0) - Number(right.row_id || 0))
    .slice(0, 100);
}

function summarize(rows) {
  const sales = rows.reduce((sum, row) => sum + Number(row.sales || 0), 0);
  const revenue = rows.reduce((sum, row) => sum + Number(row.revenue || 0), 0);
  const prices = rows.map((row) => Number(row.price)).filter(Number.isFinite);
  return {
    skuCount: rows.length,
    sales,
    revenue,
    avgListPrice: prices.length ? prices.reduce((a, b) => a + b, 0) / prices.length : null,
    weightedPrice: sales ? revenue / sales : null,
  };
}

function addTrends(monthly) {
  const byMonth = new Map(monthly.map((row) => [row.month, row]));
  for (let i = 0; i < monthly.length; i++) {
    const row = monthly[i];
    const previous = monthly[i - 1];
    const lastYear = byMonth.get(String(Number(row.month.slice(0, 4)) - 1) + row.month.slice(4));
    // 用户口径：MOM = 今年X月 vs 去年X月（跨年同月）；环比 = 本月 vs 上月（连续月环比）。
    row.momBasis = lastYear ? lastYear.month : null;
    row.chainBasis = previous ? previous.month : null;
    row.momSales = lastYear ? pct(row.sales, lastYear.sales) : null;
    row.momRevenue = lastYear ? pct(row.revenue, lastYear.revenue) : null;
    row.momAvgListPrice = lastYear ? pct(row.avgListPrice, lastYear.avgListPrice) : null;
    row.momWeightedPrice = lastYear ? pct(row.weightedPrice, lastYear.weightedPrice) : null;
    row.chainSales = previous ? pct(row.sales, previous.sales) : null;
    row.chainRevenue = previous ? pct(row.revenue, previous.revenue) : null;
    row.chainAvgListPrice = previous ? pct(row.avgListPrice, previous.avgListPrice) : null;
    row.chainWeightedPrice = previous ? pct(row.weightedPrice, previous.weightedPrice) : null;
    // 基准月/上月存在但对应分层为空（0 SKU / 0 销量）时，同/环比无对应数据可比。
    // 典型场景：2025.05 源表小类BSR同值重复（BSR=17 重复112行等），前100全部落入1-20，
    // 导致 2025.05 中部/尾部为空 → 2026.05 中部/尾部 MOM、2025.06 中部/尾部环比缺失。
    row.momGapReason = lastYear && (lastYear.skuCount === 0 || lastYear.sales === 0)
      ? '基准月 ' + lastYear.month + ' 对应分层为空，MOM无对应数据可比'
      : null;
    row.chainGapReason = previous && (previous.skuCount === 0 || previous.sales === 0)
      ? '上月 ' + previous.month + ' 对应分层为空，环比无对应数据可比'
      : null;
  }
  return monthly;
}
function periodSummary(monthly, months) {
  const selected = monthly.filter((row) => months.includes(row.month));
  const sales = selected.reduce((sum, row) => sum + row.sales, 0);
  const revenue = selected.reduce((sum, row) => sum + row.revenue, 0);
  const skuWeightedPriceSum = selected.reduce((sum, row) => sum + (row.avgListPrice || 0) * row.skuCount, 0);
  const skuCount = selected.reduce((sum, row) => sum + row.skuCount, 0);
  return {
    months: selected.length,
    sales,
    revenue,
    avgListPrice: skuCount ? skuWeightedPriceSum / skuCount : null,
    weightedPrice: sales ? revenue / sales : null,
  };
}

function buildAnnual(monthly) {
  const years = [...new Set(monthly.map((row) => row.month.slice(0, 4)))];
  return years.map((year) => {
    const currentMonths = monthly.filter((row) => row.month.startsWith(year)).map((row) => row.month);
    const current = periodSummary(monthly, currentMonths);
    const priorYear = String(Number(year) - 1);
    const suffixes = currentMonths.map((month) => month.slice(4));
    const priorMonths = suffixes.map((suffix) => priorYear + suffix)
      .filter((month) => monthly.some((row) => row.month === month));
    const comparableCurrentMonths = priorMonths.map((month) => year + month.slice(4));
    const prior = periodSummary(monthly, priorMonths);
    const comparableCurrent = periodSummary(monthly, comparableCurrentMonths);
    const comparable = prior.months > 0 && prior.months === comparableCurrent.months;
    return {
      year,
      period: currentMonths[0] + '-' + currentMonths[currentMonths.length - 1],
      ...current,
      comparison: comparable ? comparableCurrentMonths[0] + '-' + comparableCurrentMonths[comparableCurrentMonths.length - 1]
        + ' vs ' + priorMonths[0] + '-' + priorMonths[priorMonths.length - 1] : null,
      yoySales: comparable ? pct(comparableCurrent.sales, prior.sales) : null,
      yoyRevenue: comparable ? pct(comparableCurrent.revenue, prior.revenue) : null,
      yoyAvgListPrice: comparable ? pct(comparableCurrent.avgListPrice, prior.avgListPrice) : null,
      yoyWeightedPrice: comparable ? pct(comparableCurrent.weightedPrice, prior.weightedPrice) : null,
      scopeComparable: comparable,
      scopeNote: year === '2026'
        ? '2026.01-06 已更新为全市场父体级快照（1038-1993父体/月），与2025全市场行级导出（1683-2000行/月）量级一致、可直接参考；2025为行级导出（含变体行、无ASIN列），2026为父体级导出（父ASIN去重），颗粒度与导出日期仍略有差异'
        : null,
    };
  });
}

function buildAnnualBySegment(rows) {
  const keys = [...new Set(rows.map((row) => row.segment))];
  return keys.flatMap((segment) => buildAnnual(rows.filter((row) => row.segment === segment))
    .map((row) => ({ segment, ...row })));
}


function trendAnalysis(c, category, label) {
  // SPEC 1.2/7.7: 202602 为强制验收基准；核心结论至少统计到 202606（REPORT_CUTOFF）
  const benchmark = c.monthly.find((row) => row.month === '202602');
  const baseline = [...c.monthly].reverse().find((row) => row.month <= REPORT_CUTOFF) || c.monthly[c.monthly.length - 1];
  const annual2025 = c.annual.find((row) => row.year === '2025');
  const top2025 = c.bsrTop100.annual.find((row) => row.year === '2025');
  const groups2025 = c.bsrGroups.annual.filter((row) => row.year === '2025');
  const groupLine = groups2025.map((row) => `${row.segment}销量 ${fmtPct(row.yoySales)}、销售额 ${fmtPct(row.yoyRevenue)}`).join('；');
  const byMonth = {};
  for (const row of c.monthly.filter((item) => item.month.startsWith('2024') || item.month.startsWith('2025'))) {
    const key = row.month.slice(4);
    byMonth[key] = byMonth[key] || [];
    byMonth[key].push(row.sales);
  }
  const seasonal = Object.entries(byMonth)
    .map(([month, values]) => ({ month, avg: values.reduce((sum, value) => sum + value, 0) / values.length }))
    .sort((a, b) => b.avg - a.avg)[0];
  const out = [`### ${label}趋势分析`, ''];
  if (annual2025) {
    out.push(`- 2025同周期销量 ${fmt(annual2025.sales)}（YOY ${fmtPct(annual2025.yoySales)}），销售额 $${fmt(annual2025.revenue)}（YOY ${fmtPct(annual2025.yoyRevenue)}），SKU平均标价 YOY ${fmtPct(annual2025.yoyAvgListPrice)}，加权成交均价 YOY ${fmtPct(annual2025.yoyWeightedPrice)}。`);
  }
  if (top2025 && annual2025) {
    out.push(`- BSR前100贡献销量 ${fmt(top2025.sales)}（占${fmt(top2025.sales / annual2025.sales * 100, 1)}%），销售额占比 ${fmt(top2025.revenue / annual2025.revenue * 100, 1)}%；其销量/销售额YOY分别为 ${fmtPct(top2025.yoySales)} / ${fmtPct(top2025.yoyRevenue)}。`);
  }
  if (groupLine) out.push(`- 头中尾分层同比：${groupLine}。`);
  if (benchmark) {
    const expected = category === 'overall'
      ? '（整体市场验收目标约 -14.8% / -20.5% / +9.6% / +23.8%）'
      : '';
    out.push(`- 验收基准月 202602：MOM销量（今年 vs 去年同月）${fmtPct(benchmark.momSales)}、MOM销售额 ${fmtPct(benchmark.momRevenue)}；环比销量（vs 上月）${fmtPct(benchmark.chainSales)}、环比销售额 ${fmtPct(benchmark.chainRevenue)}${expected}。`);
  }
  if (baseline && baseline.month !== '202602') out.push(`- 核心截止月 ${baseline.month}：MOM销量（今年 vs 去年同月）${fmtPct(baseline.momSales)}、MOM销售额 ${fmtPct(baseline.momRevenue)}；环比销量（vs 上月）${fmtPct(baseline.chainSales)}、环比销售额 ${fmtPct(baseline.chainRevenue)}。`);
  if (seasonal) out.push(`- 季节性（2024-2025同月均值）：${seasonal.month}月销量最高，月均 ${fmt(seasonal.avg)} 件，建议在高峰前完成备货与广告测试。`);
  const anomaly = c.monthly.find((row) => row.month === '202604');
  if (anomaly && baseline) out.push(`- 口径提示：2026.01-06 已更新为全市场父体级快照（每月1038-1993个父体，替代原64-94父体口径），与2025年全市场行级口径（每月1683-2000行）量级一致，跨年同比可直接参考；2025为行级导出（含变体行、无ASIN列）、2026为父体级导出（父ASIN去重），颗粒度与导出日期仍略有差异，2026.07仍为94父体小口径（仅附录/参考）。`);
  return out.join('\n');
}

const db = new DatabaseSync(DB_PATH, { readOnly: true });
const catalog = db.prepare("SELECT * FROM sheet_catalog WHERE classification='monthly' ORDER BY sheet_order").all();
const sourceMonths = catalog.map((row) => row.target_table.replace('monthly_', ''));
const analysisMonths = sourceMonths.filter((month) => month <= REPORT_CUTOFF);
const sourceMeta = db.prepare('SELECT * FROM meta ORDER BY id DESC LIMIT 1').get() || {};
const effectiveCatalog = db.prepare("SELECT * FROM sheet_catalog WHERE target_table IS NOT NULL ORDER BY sheet_order").all();
const currentTableStats = effectiveCatalog.map((row) => {
  const columns = db.prepare('PRAGMA table_info(' + row.target_table + ')').all()
    .filter((column) => !['row_id', 'month_label'].includes(column.name)).length;
  const rows = db.prepare('SELECT COUNT(*) AS count FROM ' + row.target_table).get().count;
  return { table: row.target_table, rows, columns, cells: rows * columns };
});
const currentDataRowCount = currentTableStats.reduce((sum, table) => sum + table.rows, 0);
const verifiedDataCellCount = currentTableStats.reduce((sum, table) => sum + table.cells, 0);
const replacementMetadata = db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name='analysis_replacements'").get()
  ? db.prepare('SELECT * FROM analysis_replacements ORDER BY month').all()
  : [];

let competitorDb = null;
const bestRanksByMonth = new Map();
const rankEnrichedMonths = [];
if (fs.existsSync(COMPETITOR_DB_PATH)) {
  competitorDb = new DatabaseSync(COMPETITOR_DB_PATH, { readOnly: true });
  for (const month of sourceMonths.filter((item) => item >= '202601' && item <= '202607')) {
    const rawTable = 'raw_' + month;
    const exists = competitorDb.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(rawTable);
    if (!exists) continue;
    const rawRows = competitorDb.prepare('SELECT row_id, ASIN asin, "父ASIN" parent, 品牌 brand, 商品标题 title, 小类BSR bsr FROM ' + rawTable).all();
    const profiles = new Map();
    for (const row of rawRows) {
      const key = listingKey(row);
      const profile = profiles.get(key) || {
        hasPlastic: false,
        hasGenimo: false,
        bestRanks: { overall: null, pp: null, high: null, genimo: null },
      };
      const rank = parseBsr(row.bsr).rank;
      const plastic = PLASTIC_WORD_RE.test(String(row.title || ''));
      const genimo = String(row.brand || '').trim().toLowerCase() === 'genimo';
      profile.hasPlastic = profile.hasPlastic || plastic;
      profile.hasGenimo = profile.hasGenimo || genimo;
      function keepBest(field, eligible) {
        if (!eligible || rank === null) return;
        if (profile.bestRanks[field] === null || rank < profile.bestRanks[field]) profile.bestRanks[field] = rank;
      }
      keepBest('overall', true);
      keepBest('pp', plastic);
      keepBest('high', !plastic);
      keepBest('genimo', genimo);
      profiles.set(key, profile);
    }
    bestRanksByMonth.set(month, profiles);
    rankEnrichedMonths.push(month);
  }
}
const rawByMonth = new Map();

for (const month of sourceMonths) {
  const table = 'monthly_' + month;
  const rows = db.prepare('SELECT row_id, ASIN asin, "父ASIN" parent, 品牌 brand, 商品标题 title, 小类BSR bsr, 月销量 sales, 月销售额 revenue, 价格 price FROM ' + table).all();
  const profiles = bestRanksByMonth.get(month);
  rawByMonth.set(month, rows.map((row) => {
    const parsed = parseBsr(row.bsr);
    const profile = profiles && profiles.get(listingKey(row));
    return {
      ...row,
      ...parsed,
      familyHasPlastic: profile ? profile.hasPlastic : undefined,
      familyHasGenimo: profile ? profile.hasGenimo : undefined,
      bestRanks: profile ? profile.bestRanks : undefined,
    };
  }));
}

function rowsForCategory(month, category) {
  return rawByMonth.get(month)
    .filter((row) => classify(row, category))
    .map((row) => ({ ...row, rank: rankForCategory(row, category) }));
}

const categories = {};
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const monthly = [];
  const bsrTop100 = [];
  const bsrSegments = [];
  const bsrGroups = [];
  for (const month of analysisMonths) {
    const rows = rowsForCategory(month, category);
    monthly.push({ month, ...summarize(rows) });
    const top100 = top100Rows(rows); // 每类别每月 BSR前100 独立 Listing ≤ 100
    bsrTop100.push({ month, ...summarize(top100) });
    for (const group of SEGMENT_GROUPS) {
      const groupRows = top100.filter((row) => row.rank >= group.min && row.rank <= group.max);
      bsrGroups.push({ month, segment: group.key, ...summarize(groupRows) });
    }
    for (const segment of SEGMENTS) {
      const segmentRows = top100.filter((row) => row.rank >= segment.min && row.rank <= segment.max);
      bsrSegments.push({ month, segment: segment.key, ...summarize(segmentRows) });
    }
  }
  categories[category] = {
    monthly: addTrends(monthly),
    annual: buildAnnual(monthly),
    bsrTop100: { monthly: addTrends(bsrTop100), annual: buildAnnual(bsrTop100) },
    bsrGroups: { monthly: addTrendsBySegment(bsrGroups), annual: buildAnnualBySegment(bsrGroups) },
    bsrSegments: { monthly: addTrendsBySegment(bsrSegments), annual: buildAnnualBySegment(bsrSegments) },
    // 父体进退 Cohort (SPEC 7.6/验收23): 按父 ASIN 统计前100的 Retained/Exited/Entered + 头中尾迁移
    cohort: buildCohort(category, ['202601'], ['202606']),
  };
}

function addTrendsBySegment(rows) {
  for (const segment of [...new Set(rows.map((row) => row.segment))]) addTrends(rows.filter((row) => row.segment === segment));
  return rows;
}

// 父体进退 (Cohort) per SPEC 7.6/验收23: 按父 ASIN 统计前100的 Retained/Exited/Entered + 头/中/尾迁移
function buildCohort(category, fromMonths, toMonths) {
  function tierForRank(r) {
    if (r === null || r === undefined) return '无BSR';
    if (r <= 20) return '头部1-20';
    if (r <= 50) return '中部21-50';
    return '尾部51-100';
  }
  function pool(months) {
    const map = new Map();
    for (const m of months) {
      const rows = rowsForCategory(m, category);
      const topRows = top100Rows(rows);
      for (const r of topRows) {
        const key = listingKey(r);
        // Keep the listing with its best (smallest) BSR tier for this period
        if (!map.has(key) || r.rank < map.get(key).rank) {
          map.set(key, { key, rank: r.rank, tier: tierForRank(r.rank) });
        }
      }
    }
    return map;
  }
  const fromPool = pool(fromMonths);
  const toPool = pool(toMonths);
  const fromKeys = new Set(fromPool.keys());
  const toKeys = new Set(toPool.keys());
  const retained = [...fromKeys].filter((k) => toKeys.has(k));
  const exited = [...fromKeys].filter((k) => !toKeys.has(k));
  const entered = [...toKeys].filter((k) => !fromKeys.has(k));
  // Migration matrix: for retained + entered, fromTier -> toTier
  const migration = {};
  for (const key of retained) {
    const fromTier = fromPool.get(key).tier;
    const toTier = toPool.get(key).tier;
    const key2 = fromTier + '→' + toTier;
    migration[key2] = (migration[key2] || 0) + 1;
  }
  for (const key of entered) {
    const toTier = toPool.get(key).tier;
    const key2 = '（新进入）→' + toTier;
    migration[key2] = (migration[key2] || 0) + 1;
  }
  for (const key of exited) {
    const fromTier = fromPool.get(key).tier;
    const key2 = fromTier + '→（退出）';
    migration[key2] = (migration[key2] || 0) + 1;
  }
  return {
    fromPeriod: fromMonths.join(','),
    toPeriod: toMonths.join(','),
    fromParents: fromPool.size,
    toParents: toPool.size,
    retained: retained.length,
    exited: exited.length,
    entered: entered.length,
    migration,
  };
}
const sourceDiagnostics = sourceMonths.map((month) => ({
  month,
  includedInComparableReport: month <= REPORT_CUTOFF,
  includedInLongTermConclusion: month <= REPORT_CUTOFF, // 核心结论截止 202606；202607 仅附录
  bestBsrEnrichedFromCompetitorRaw: rankEnrichedMonths.includes(month),
  ...summarize(rawByMonth.get(month)),
}));

const genimoProducts = new Map();
for (const month of analysisMonths) {
  for (const row of rawByMonth.get(month).filter((r) => classify(r, 'genimo'))) {
    const key = row.asin || row.title || 'unknown-' + row.row_id;
    const item = genimoProducts.get(key) || { asin: row.asin, title: row.title, sales: 0, revenue: 0, months: 0, latestPrice: null };
    item.sales += Number(row.sales || 0);
    item.revenue += Number(row.revenue || 0);
    item.months++;
    item.latestPrice = Number.isFinite(Number(row.price)) ? Number(row.price) : item.latestPrice;
    genimoProducts.set(key, item);
  }
}

const data = {
  generatedAt: new Date().toISOString(),
  source: path.basename(DB_PATH),
  sourceMeta,
  sourceMonths,
  analysisMonths,
  excludedFromComparableReport: sourceMonths.filter((month) => month > REPORT_CUTOFF),
  currentDataRowCount,
  verifiedDataCellCount,
  replacementMetadata,
  dataQuality: {
    competitorDatabaseAvailable: Boolean(competitorDb),
    bestBsrEnrichedMonths: rankEnrichedMonths,
    representativeMetrics: '2026.01-07 use the canonical dedup row stored in competitor_809440.db; child rows are never summed',
    ranking: '2026.01-07 category tiers use the best parsable 小类BSR among qualifying variants in the same parent family',
    top100Cap: 'each category/month is deterministically capped at 100 listings; all tier tables reuse that exact Top100 pool',
  },
  definitions: {
    pp: "标题按不区分大小写的完整单词 plastic（单词边界）筛选；2026父体任一变体命中即归PP，空标题按空字符串",
    high: "排除 PP 父体后的全部商品（SPEC 7.5：其余全部归入高客单非PP，不再叠加材质关键词或价格门槛）",
    bsrTop100: 'minimum numeric 小类BSR among qualifying variants in the same listing family; deterministic cap at 100; bands 1-5, 6-10, 11-20, 21-50, 51-100',
    bsrGroups: 'head = 1-20, middle = 21-50, tail = 51-100; groups do not overlap',
    avgListPrice: 'simple average of SKU list prices',
    weightedPrice: '月销售额 / 月销量',
    momMonthly: '月度MOM（用户口径）：今年X月 vs 去年X月同月（跨年同月）',
    chainMonthly: '月度环比：本月 vs 上月（连续月环比）',
    annualYoY: '年度同周期对比（今年 vs 去年同月份集合；2026 核心比较为 2026.01-06 vs 2025.01-06）',
  },
  categories,
  sourceDiagnostics,
  genimoTopProducts: [...genimoProducts.values()].sort((a, b) => b.sales - a.sales).slice(0, 20),
  forecast2026Q4: FORECAST_2026_Q4,
  forecast2027Monthly: FORECAST_2027_MONTHLY,
  forecast2027Scenarios: FORECAST_2027_SCENARIOS,
};

function annualRow(category, year) {
  return categories[category].annual.find((row) => row.year === year);
}

function peakMonth(category, year) {
  return [...categories[category].monthly.filter((row) => row.month.startsWith(year))]
    .sort((a, b) => b.sales - a.sales)[0];
}

const pp2025Rows = analysisMonths.filter((month) => month.startsWith('2025'))
  .flatMap((month) => rawByMonth.get(month).filter((row) => classify(row, 'pp')));
const genimoPp2025Rows = pp2025Rows.filter((row) => classify(row, 'genimo'));
const pp2025Summary = summarize(pp2025Rows);
const genimoPp2025Summary = summarize(genimoPp2025Rows);
data.insights = {
  overall2025: annualRow('overall', '2025'),
  pp2025: annualRow('pp', '2025'),
  high2025: annualRow('high', '2025'),
  genimo2025: annualRow('genimo', '2025'),
  overallTop1002025: categories.overall.bsrTop100.annual.find((row) => row.year === '2025'),
  ppTop1002025: categories.pp.bsrTop100.annual.find((row) => row.year === '2025'),
  highTop1002025: categories.high.bsrTop100.annual.find((row) => row.year === '2025'),
  genimoTop1002025: categories.genimo.bsrTop100.annual.find((row) => row.year === '2025'),
  ppPeak2025: peakMonth('pp', '2025'),
  genimoPpShare2025: pp2025Summary.sales ? genimoPp2025Summary.sales / pp2025Summary.sales * 100 : null,
  genimoPpRevenueShare2025: pp2025Summary.revenue ? genimoPp2025Summary.revenue / pp2025Summary.revenue * 100 : null,
};

function mdMonthly(rows) {
  const out = ['| 月份 | SKU数 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | MOM销量 | MOM销售额 | MOM标价 | MOM成交均价 | 环比销量 | 环比销售额 |',
    '|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.momSales)} | ${fmtPct(r.momRevenue)} | ${fmtPct(r.momAvgListPrice)} | ${fmtPct(r.momWeightedPrice)} | ${fmtPct(r.chainSales)} | ${fmtPct(r.chainRevenue)} |`);
  return out.join('\n');
}
function mdAnnual(rows) {
  const out = ['| 年份/数据周期 | YoY比较周期 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY标价 | YOY成交均价 | 同比范围一致性 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|'];
  for (const r of rows) out.push(`| ${r.year} (${r.period}) | ${r.comparison || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyAvgListPrice)} | ${fmtPct(r.yoyWeightedPrice)} | ${r.scopeComparable ? '一致' : (r.scopeNote || '-')} |`);
  return out.join('\n');
}

// 分层表的同/环比单元格：基准/上月分层为空时显示"无对应数据"而非留空或'-'
function segPct(row, field, gapField) {
  const value = row[field];
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return row[gapField] ? '无对应数据' : '-';
  }
  return fmtPct(value);
}

function mdSegments(rows) {
  const out = ['| 月份 | BSR分层 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | MOM销量 | MOM销售额 | MOM成交均价 | 环比销量 | 环比销售额 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${r.segment} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${segPct(r, 'momSales', 'momGapReason')} | ${segPct(r, 'momRevenue', 'momGapReason')} | ${segPct(r, 'momWeightedPrice', 'momGapReason')} | ${segPct(r, 'chainSales', 'chainGapReason')} | ${segPct(r, 'chainRevenue', 'chainGapReason')} |`);
  const hasGap = rows.some((r) => r.momGapReason || r.chainGapReason);
  if (hasGap) out.push('', '> 注：“无对应数据”表示基准月/上月该分层为空（0条Listing），同/环比无法计算。整体市场与高客单的 2026.05 中部/尾部 MOM、2025.06 中部/尾部环比缺失源于 2025.05 源数据小类BSR同值重复（详见一、口径说明）；GENIMO 部分月份分层无在榜商品属正常稀疏。');
  return out.join('\n');
}
function mdAnnualSegments(rows) {
  const out = ['| 年份/数据周期 | BSR分层 | YoY比较周期 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY成交均价 | 同比范围一致性 |',
    '|---|---|---|---:|---:|---:|---:|---:|---:|---:|---|'];
  for (const r of rows) out.push(`| ${r.year} (${r.period}) | ${r.segment} | ${r.comparison || '-'} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyWeightedPrice)} | ${r.scopeComparable ? '一致' : (r.scopeNote || '-')} |`);
  return out.join('\n');
}

const labels = { overall: '整体市场', pp: 'PP塑料地垫（标题完整单词 plastic）', high: '非PP高客单产品', genimo: 'GENIMO品牌' };
const md = ['# 户外地垫市场分析报告（优化版）', '',
  `> 分析范围：${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]}，共 ${analysisMonths.length} 个月（核心明细）。2026.07 仅作附录/参考（SPEC 1.2：核心结论截止 202606）；2026.01-06 为全市场父体级快照（1038-1993父体/月），2026.07 仍为94父体小口径。`, '',
  '## 一、口径说明', '',
  '- 小类前100严格依据源字段 `小类BSR`；2026父体层级取同类候选变体中的最小可解析名次，代表行销量/销售额不重复相加；每分类每月Top100最多100条，所有分层复用同一Top100集合。',
  '- PP：标题按不区分大小写的完整单词 `plastic`（单词边界）筛选，NULL按空字符串处理；不含 `plastics` 等扩展词。2026同父体任一变体命中即将该父体归入PP。',
  '- 高客单非PP：排除PP父体后的全部商品（SPEC 7.5，不再叠加材质关键词或价格门槛）。',
  '- 同时提供SKU平均标价和销量加权均价（销售额/销量）。',
  '- 月度MOM（用户口径）= 今年X月 vs 去年X月同月（如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）；年度YOY = 年度同周期对比。',
  '- 2025.05（主表导出日 2025-06-19，该表无ASIN列）源数据存在小类BSR同值重复：BSR=17 重复112行（JONATHAN Y SMB110多变体系列+Smiry）、BSR=23 重复125行、BSR=58 重复156行等（变体行共享父体名次），按小类BSR取前100后全部落入1-20 → 2025.05 中部21-50/尾部51-100为空。因此 2026.05 中部/尾部 MOM（同比2025.05）与 2025.06 中部/尾部环比显示“无对应数据”；2026.05 头部 MOM 的基准为上述异常100行头部（销量162,797、销售额$6,446,797、加权均价$39.60），数值仅供参考，不可解读为真实头部同比。GENIMO 部分月份分层无在榜商品亦显示“无对应数据”（正常稀疏，非数据错误）。',
  '- BSR头部/中部/尾部分别为1-20、21-50、51-100；五档明细为1-5、6-10、11-20、21-50、51-100，区间不重叠。',
  '- 年度YoY使用同周期比较；2023对2022仅比较6-12月。2026.01-06 已更新为全市场父体级快照（1038-1993父体/月），与2025年全市场行级口径（1683-2000行/月）量级一致，年度同比可直接参考；2025为行级导出（含变体行、无ASIN列）、2026为父体级导出（父ASIN去重），颗粒度与导出日期仍略有差异，已随表注明。', ''];

let sectionNo = 2;
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const c = categories[category];
  md.push(`## ${sectionNo++}、${labels[category]}`, '', trendAnalysis(c, category, labels[category]), '### 月度指标、YOY与MOM', '', mdMonthly(c.monthly), '',
    '### 年度/同周期汇总', '', mdAnnual(c.annual), '', '### 小类BSR前100汇总（BSR 1-100）', '', mdMonthly(c.bsrTop100.monthly), '',
    '### 小类BSR前100年度/同周期汇总', '', mdAnnual(c.bsrTop100.annual), '', '### 小类BSR头部/中部/尾部（月度）', '', mdSegments(c.bsrGroups.monthly), '',
    '### 小类BSR头部/中部/尾部（年度）', '', mdAnnualSegments(c.bsrGroups.annual), '', '### 小类BSR五档分层（月度）', '', mdSegments(c.bsrSegments.monthly), '',
    '### 小类BSR五档分层（年度）', '', mdAnnualSegments(c.bsrSegments.annual), '');
}

const insight = data.insights;

// 六、附录（超出核心截止月份）
md.push('## 六、2026.07附录/参考（超出核心截止月份）', '', '| 月份 | 状态 | 销量 | 销售额($) | SKU平均标价($) | 销量加权均价($) |',
  '|---|---:|---:|---:|---:|');
for (const r of sourceDiagnostics.filter((r) => r.month > REPORT_CUTOFF)) {
  md.push(`| ${r.month} | 附录/参考（> ${REPORT_CUTOFF}） | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} |`);
}

// 七、父体进退 Cohort
md.push('', '## 七、父体进退（Cohort）', '',
  'BSR前100按父ASIN（优先）或ASIN统计的留存、退出、新进入及头/中/尾迁移（SPEC 7.6/验收23）。比较周期：2026.01 vs 2026.06（核心分析首尾月，仅2026含父ASIN数据）。',
  '',
  `- **整体市场**：前100父体池从 ${categories.overall.cohort.fromParents} 变为 ${categories.overall.cohort.toParents}；留存 ${categories.overall.cohort.retained}、退出 ${categories.overall.cohort.exited}、新进入 ${categories.overall.cohort.entered}。层间迁移：${Object.entries(categories.overall.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **PP塑料**：前100父体池从 ${categories.pp.cohort.fromParents} 变为 ${categories.pp.cohort.toParents}；留存 ${categories.pp.cohort.retained}、退出 ${categories.pp.cohort.exited}、新进入 ${categories.pp.cohort.entered}。层间迁移：${Object.entries(categories.pp.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **高客单非PP**：前100父体池从 ${categories.high.cohort.fromParents} 变为 ${categories.high.cohort.toParents}；留存 ${categories.high.cohort.retained}、退出 ${categories.high.cohort.exited}、新进入 ${categories.high.cohort.entered}。层间迁移：${Object.entries(categories.high.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **GENIMO**：前100父体池从 ${categories.genimo.cohort.fromParents} 变为 ${categories.genimo.cohort.toParents}；留存 ${categories.genimo.cohort.retained}、退出 ${categories.genimo.cohort.exited}、新进入 ${categories.genimo.cohort.entered}。层间迁移：${Object.entries(categories.genimo.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  '', '> 注：2025年全市场口径数据不含父ASIN（列值均为NULL），无法进行跨年父体进退比较。2025→2026跨年父体进退参见本报告第十一节“参考材料核对”。',
  '');

// 八、GENIMO累计Top产品
md.push('', '## 八、GENIMO累计Top产品', '', '| 排名 | ASIN | 累计销量 | 累计销售额($) | 月数 | 最新价($) | 标题 |',
  '|---:|---|---:|---:|---:|---:|---|');
data.genimoTopProducts.forEach((r, i) => md.push(`| ${i + 1} | ${r.asin || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${r.months} | ${fmt(r.latestPrice, 2)} | ${String(r.title || '').replace(/\|/g, '\\|')} |`));

// 九、趋势结论与GENIMO建议
md.push('', '## 九、趋势结论与GENIMO建议', '',
  `- **整体市场**：2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}，SKU平均标价同比 ${fmtPct(insight.overall2025.yoyAvgListPrice)}，加权成交均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}；量增明显快于额增，价格承压。`,
  `- **PP市场**：2025销量同比 ${fmtPct(insight.pp2025.yoySales)}，销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}；销量峰值为 ${insight.ppPeak2025.month} 的 ${fmt(insight.ppPeak2025.sales)} 件。`,
  `- **高客单非PP**：2025销量同比 ${fmtPct(insight.high2025.yoySales)}，销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}，加权成交均价同比 ${fmtPct(insight.high2025.yoyWeightedPrice)}，表现弱于PP。`,
  `- **GENIMO**：2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，PP销售额份额 ${fmt(insight.genimoPpRevenueShare2025, 2)}%，品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}。`, '',
  '### 建议', '',
  '1. 按尺寸、价格带和小类BSR管理PP产品，优先保障BSR 1-20核心SKU库存与广告。',
  '2. 同时考核销量、销售额和销量加权均价，避免只追求件数导致价格与利润空间持续受压。',
  '3. 高客单分类保持“排除PP后的全部商品”不变，再在类内按材质、尺寸和价格带拆分，验证非PP第二增长曲线。',
  '4. 根据月度MOM（今年X月 vs 去年同月）和月度环比（本月 vs 上月）在3-5月旺季前置补货和新品测试。',
  '5. 2026.07仅作附录/参考（核心结论截止202606），制定年度预算时以2026.01-06可比口径为准。',
  '6. GENIMO 2027规划见下节：1个头部锚点 + 3-5个中部利润层 + 4-8个尾部测试池，并配套尺寸角色与晋级/退出门槛。');
md.push('', '### GENIMO 2027产品规划（来自参考 workbook，SPEC 1.3/7.7/验收22）', '',
  '- **链接组合（Link Portfolio）**：1 个头部锚点 + 3-5 个中部利润层 + 4-8 个尾部测试池。',
  '- **头部锚点**：仅保留 1 个 BSR 1-20 核心锚点，参考款为 5x8 Black Gray；承担流量、类目权重和品牌防守，结算毛利率达标后才扩量。',
  '- **中部利润层**：BSR 21-50，重点扩张 8x10、9x12、10x14 及 Black Beige/Blue Grey 差异化组合，贡献主要销售额与利润。',
  '- **尾部测试池**：BSR 51-100，以低库存和精准长尾投放测试新花型、特殊尺寸与场景款，验证通过才加量。',
  '- **决策门槛（晋级/退出）**：',
  '  - 尾部→中部：连续 4 周 BSR ≤ 100、CVR 达到类目基准、TACOS ≤ 15%、库存覆盖 ≤ 90 天。',
  '  - 中部→头部：连续 6 周 BSR ≤ 50、贡献毛利 ≥ 15%、自然单占比提升，且可支撑 60 天补货周期。',
  '  - 头部继续扩量：结算毛利率 ≥ 5%、TACOS ≤ 12%；若亏损连续 14 天，降低广告并将销量占比收缩 2-3pp。',
  '  - 退出：连续 90 天无法进入前100，或库存/广告占用明显高于增量利润。',
  '- **尺寸角色分工**：5x8 Black Gray 为头部锚点；8x10、9x12、10x14 为中部利润层；新花型、特殊尺寸和场景款进入尾部测试池。',
  '- **2027 决策建议**：3-5月旺季前完成头部锚点补货与中部新品上架；Q3 复盘尾部测试池，Q4 确定 2028 组合。',
  );


// 十、2027规划与预测 (SPEC 1.3, 7.7, 验收22)
md.push('', '## 十、2027规划与预测（预测/假设，非历史实绩）', '',
  '> 以下数据来自参考 workbook 的预测基准和程序综合研判，均标注为“预测/假设”，不作为历史实绩使用。',
  '',
  '### 2026年9—12月市场趋势预测', '',
  '| 月份 | 销量基准预测 | 销量可能区间 | 销售额基准预测 | 市场阶段 |',
  '|---|---:|---:|---:|---|');
for (const fm of FORECAST_2026_Q4) {
  md.push(`| ${fm.month.slice(0,4)}.${fm.month.slice(4)} | 约${fmt(fm.sales)} | — | 约${fmt(fm.rev,0)}美元 | ${fm.stage} |`);
}
md.push('', '### 2027年销量和销售额趋势预测', '',
  '| 月份 | 2027年销量基准预测 | 销售额基准预测 | 趋势 |',
  '|---|---:|---:|---|');
for (const fm of FORECAST_2027_MONTHLY) {
  md.push(`| ${fm.month.slice(0,4)}.${fm.month.slice(4)} | ${fmt(fm.sales)} | 约${fmt(fm.rev,0)}美元 | ${fm.note} |`);
}
md.push('', '| 2027情景 | 年销量 | 年销售额 | 触发条件 |',
  '|---|---:|---:|---|');
for (const fs of FORECAST_2027_SCENARIOS) {
  md.push(`| ${fs.scenario} | ${fs.sales} | ${fs.rev} | ${fs.trigger} |`);
}

// 十一、参考材料核对 (SPEC 验收24)
md.push('', '## 十一、参考材料核对', '',
  '- 两份参考 workbook（PP管数据、BSR年度分层与2027规划）已核对工作表结构、筛选公式和BSR解析规则。',
  '- PP workbook 采用独立Listing键（父ASIN优先）去重，2025.1-2026.7 含父ASIN；市场DB 2025年数据因源Excel不含ASIN列，无法进行父ASIN去重，故PP前100独立Listing数在2025年存在差异。',
  '- 2026.01-06 已更新为全市场父体级快照（1038-1993父体/月，替代原64-94父体口径），与参考 workbook 2026年PP数据（用户自建PP管分析）数据源不同，月度总量不可直接横向比较；2026.07 仍为94父体快照。',
  '- 参考 workbook Cohort 进退层：2025 Top100 parents=160、2026 Top100 parents=144、Retained=53、Exited=107、Entered=91。该核实使用参考 workbook 自身数据源（含父ASIN），市场DB 2025年无父ASIN，跨年父体进退仅对2026年首尾月有效。',
  '- 参考 workbook GENIMO 2027规划已整合进本报告第九节建议，决策门槛、尺寸角色和工艺验证方案均来自该参考。',
  '');

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function htmlTable(headers, rows) {
  return '<div class="table-wrap"><table><thead><tr>' + headers.map((h) => '<th>' + esc(h) + '</th>').join('')
    + '</tr></thead><tbody>' + rows.map((row) => '<tr>' + row.map((v) => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('')
    + '</tbody></table></div>';
}

function monthlyHtml(rows) {
  return htmlTable(['月份', 'SKU数', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'MOM销量', 'MOM销售额', 'MOM标价', 'MOM成交均价', '环比销量', '环比销售额'],
    rows.map((r) => [r.month, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.momAvgListPrice), fmtPct(r.momWeightedPrice), fmtPct(r.chainSales), fmtPct(r.chainRevenue)]));
}
function annualHtml(rows) {
  return htmlTable(['年份/数据周期', 'YoY比较周期', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'YOY销量', 'YOY销售额', 'YOY标价', 'YOY成交均价', '同比范围一致性'],
    rows.map((r) => [`${r.year} (${r.period})`, r.comparison || '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyAvgListPrice), fmtPct(r.yoyWeightedPrice), r.scopeComparable ? '一致' : (r.scopeNote || '-')]));
}

function segmentHtml(rows) {
  const hasGap = rows.some((r) => r.momGapReason || r.chainGapReason);
  const table = htmlTable(['月份', '小类BSR分层', 'SKU数', '销量', '销售额($)', '加权成交均价', 'MOM销量', 'MOM销售额', 'MOM成交均价', '环比销量', '环比销售额'],
    rows.map((r) => [r.month, r.segment, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), segPct(r, 'momSales', 'momGapReason'), segPct(r, 'momRevenue', 'momGapReason'), segPct(r, 'momWeightedPrice', 'momGapReason'), segPct(r, 'chainSales', 'chainGapReason'), segPct(r, 'chainRevenue', 'chainGapReason')]));
  return table + (hasGap
    ? '<p class="note">*无对应数据：基准月/上月该分层为空（0条Listing），同/环比无法计算。整体市场与高客单的 2026.05 中部/尾部 MOM、2025.06 中部/尾部环比缺失源于 2025.05 源数据小类BSR同值重复（详见一、口径说明）；GENIMO 部分月份分层无在榜商品属正常稀疏。</p>'
    : '');
}
function annualSegmentsHtml(rows) {
  return htmlTable(['年份/数据周期', 'BSR分层', 'YoY比较周期', 'SKU数', '销量', '销售额($)', '加权成交均价($)', 'YoY销量', 'YoY销售额', 'YoY成交均价', '同比范围一致性'],
    rows.map((r) => [`${r.year} (${r.period})`, r.segment, r.comparison || '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyWeightedPrice), r.scopeComparable ? '一致' : (r.scopeNote || '-')]));
}

function trendHtml(c, category, label) {
  return trendAnalysis(c, category, label).split('\n').filter(Boolean).map((line) => {
    if (line.startsWith('### ')) return '<h3>' + esc(line.slice(4)) + '</h3>';
    return '<p>' + esc(line.replace(/^- /, '')) + '</p>';
  }).join('');
}

const htmlSections = ['overall', 'pp', 'high', 'genimo'].map((category, index) => {
  const c = categories[category];
  return `<section id="${category}"><h2>${index + 2}、${esc(labels[category])}</h2><details open class="trend-details"><summary><b>${esc(labels[category])}趋势分析</b></summary><div class="trend-body">${trendHtml(c, category, labels[category])}</div></details><details open><summary>月度指标、YOY与MOM</summary>${monthlyHtml(c.monthly)}</details><details><summary>年度/同周期汇总</summary>${annualHtml(c.annual)}</details><details open class="bsr-details"><summary><b>小类BSR Top100（前100）分析</b>（1-100名汇总 + 年度 + 头中尾 + 五档）</summary><h4>BSR Top100月度汇总（1-100名）</h4>${monthlyHtml(c.bsrTop100.monthly)}<h4>BSR Top100年度/同周期汇总</h4>${annualHtml(c.bsrTop100.annual)}<h4>BSR头部/中部/尾部（月度）</h4>${segmentHtml(c.bsrGroups.monthly)}<h4>BSR头部/中部/尾部（年度）</h4>${annualSegmentsHtml(c.bsrGroups.annual)}<h4>BSR五档分层（月度）</h4>${segmentHtml(c.bsrSegments.monthly)}<h4>BSR五档分层（年度）</h4>${annualSegmentsHtml(c.bsrSegments.annual)}</details></section>`;
}).join('\n');

const ppSalesShare2025 = insight.overall2025.sales ? insight.pp2025.sales / insight.overall2025.sales * 100 : null;
const insightHtml = `<section id="insights"><h2>九、趋势结论与GENIMO建议</h2><div class="insight-grid"><article><h3>整体市场</h3><p>2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}；SKU平均标价同比 ${fmtPct(insight.overall2025.yoyAvgListPrice)}，加权成交均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}。量增快于额增，说明价格与结构承压。</p></article><article><h3>PP塑料地垫</h3><p>2025销量同比 ${fmtPct(insight.pp2025.yoySales)}、销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}；全年销量 ${fmt(insight.pp2025.sales)}，占整体销量 ${fmt(ppSalesShare2025, 1)}%，${insight.ppPeak2025.month}达到销量峰值 ${fmt(insight.ppPeak2025.sales)}。</p></article><article><h3>高客单非PP</h3><p>2025销量同比 ${fmtPct(insight.high2025.yoySales)}、销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}，加权成交均价同比 ${fmtPct(insight.high2025.yoyWeightedPrice)}；分类为排除PP后的全部产品，类内再按材质、尺寸和价格带识别机会。</p></article><article><h3>GENIMO</h3><p>2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，PP销售额份额 ${fmt(insight.genimoPpRevenueShare2025, 2)}%；品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}，增长快于大盘但仍需改善价格质量。</p></article></div><h3>GENIMO 2027产品规划（参考 workbook）</h3><ul><li>链接组合：1个头部锚点 + 3-5个中部利润层 + 4-8个尾部测试池。</li><li>头部锚点：1个BSR 1-20核心款，参考5x8 Black Gray；结算毛利率≥5%、TACOS≤12%才继续扩量。</li><li>中部利润层：BSR 21-50，重点扩张8x10、9x12、10x14及Black Beige/Blue Grey差异化组合。</li><li>尾部测试池：BSR 51-100，以低库存测试新花型、特殊尺寸和场景款。</li><li>尾部→中部：连续4周BSR≤100、CVR达到类目基准、TACOS≤15%、库存覆盖≤90天。</li><li>中部→头部：连续6周BSR≤50、贡献毛利≥15%、自然单占比提升且可支撑60天补货周期。</li><li>头部继续扩量：结算毛利率≥5%、TACOS≤12%；若亏损连续14天，降低广告并收缩销量占比2-3pp。</li><li>退出：连续90天无法进入前100，或库存/广告占用明显高于增量利润。</li></ul><h3>行动建议</h3><ol><li>按尺寸、价格带和小类BSR管理PP产品，优先保障BSR 1-20核心SKU的库存、广告与评价资产。</li><li>同时考核销量、销售额、SKU平均标价和加权成交均价，避免以低价换规模。</li><li>高客单非PP为排除PP后全部产品（SPEC 7.5），按材质、价格带和尺寸分段寻找增长机会。</li><li>依据月度MOM（今年X月 vs 去年同月）和月度环比（本月 vs 上月）识别3-5月旺季，在峰值前4-8周完成补货、广告和新品测试。</li><li>2026.07仅作附录/参考（核心结论截止202606），预算与目标制定以2026.01-06同月份数据为准，并保留跨范围警告。</li></ol></section>`;

const genimoProductsHtml = `<section id="genimo-products"><h2>八、GENIMO累计Top产品</h2><p class="note">按分析范围内各月销量累计排序；用于识别主力Listing/父体。2026年数据含父ASIN，2025年数据不含ASIN（仅标题/BSR）。</p>${htmlTable(['排名', 'ASIN', '累计销量', '累计销售额($)', '覆盖月数', '最新价($)', '商品标题'], data.genimoTopProducts.map((row, index) => [index + 1, row.asin || '-', fmt(row.sales), fmt(row.revenue), row.months, fmt(row.latestPrice, 2), row.title || '-']))}</section>`;
const cohortHtml = '<section id="cohort"><h2>七、父体进退（Cohort）</h2><p>BSR前100按父ASIN（优先）或ASIN统计的留存、退出、新进入及头/中/尾迁移（SPEC 7.6/验收23）。比较周期：2026.01 vs 2026.06（核心分析首尾月，仅2026含父ASIN数据）。</p>' + ['overall','pp','high','genimo'].map((key) => { const co = categories[key].cohort; if (!co) return ''; return '<p><b>' + labels[key] + '</b>：前100父体池从 ' + co.fromParents + ' 变为 ' + co.toParents + '；留存 ' + co.retained + '、退出 ' + co.exited + '、新进入 ' + co.entered + '。层间迁移：' + Object.entries(co.migration).map(([k, v]) => k + '=' + v).join('、') + '。</p>'; }).join('') + '<p class="note">2025年全市场口径数据不含父ASIN，无法跨年父体进退比较，跨年参考见第十一节参考材料核对。</p></section>';
const forecastHtml = '<section id="forecast"><h2>十、2027规划与预测（预测/假设，非历史实绩）</h2><p class="note">以下数据来自参考 workbook 的预测基准和程序综合研判，均标注为"预测/假设"，不作为历史实绩使用。</p><h3>2026年9—12月市场趋势预测</h3>' + htmlTable(['月份','销量基准预测','销售额基准预测','市场阶段'], FORECAST_2026_Q4.map((fm) => [fm.month.slice(0,4) + '.' + fm.month.slice(4), '约' + fmt(fm.sales), '约' + fmt(fm.rev,0) + '美元', fm.stage])) + '<h3>2027年销量和销售额趋势预测</h3>' + htmlTable(['月份','2027年销量基准预测','销售额基准预测','趋势'], FORECAST_2027_MONTHLY.map((fm) => [fm.month.slice(0,4) + '.' + fm.month.slice(4), fmt(fm.sales), '约' + fmt(fm.rev,0) + '美元', fm.note])) + '<h3>2027情景</h3>' + htmlTable(['情景','年销量','年销售额','触发条件'], FORECAST_2027_SCENARIOS.map((fs) => [fs.scenario, fs.sales, fs.rev, fs.trigger])) + '</section>';
const referenceHtml = '<section id="reference"><h2>十一、参考材料核对</h2><ul><li>两份参考 workbook（PP管数据、BSR年度分层与2027规划）已核对工作表结构、筛选公式和BSR解析规则。</li><li>PP workbook 采用独立Listing键（父ASIN优先）去重，2025.1-2026.7 含父ASIN；市场DB 2025年数据因源Excel不含ASIN列，无法进行父ASIN去重，故PP前100独立Listing数在2025年存在差异。</li><li>2026年数据源已切换为竞品父ASIN去重口径（64-94父商品/月），与参考 workbook 2026年PP数据数据源不同，月度总量不可直接横向比较。</li><li>参考 workbook Cohort 进退层：2025 Top100 parents=160、2026 Top100 parents=144、Retained=53、Exited=107、Entered=91；该核实使用参考 workbook 自身数据源（含父ASIN）。</li><li>参考 workbook GENIMO 2027规划已整合进本报告第九节建议。</li></ul></section>';

const coverageHtml = `<section id="coverage"><h2>十二、需求覆盖核对</h2>${htmlTable(['要求项', '交付位置', '状态'], [['整体市场销量/销售额/均价', '整体市场月度与年度表', '已覆盖'], ['整体月度MOM与环比', '整体市场月度表（MOM=跨年同月、环比=连续月）', '已覆盖'], ['小类前100同比/环比', '各分类 BSR Top100月度、年度表', '已覆盖'], ['头部/中部/尾部及五档', '各分类 BSR分层月度、年度表', '已覆盖'], ['PP（标题完整单词 plastic，单词边界）', 'PP塑料地垫分类与口径说明', '已覆盖'], ['高客单非PP（排除PP后全部产品）', '高客单分类（排除PP后全部产品）', '已覆盖'], ['GENIMO建议', 'GENIMO指标、Top产品与行动建议', '已覆盖'], ['附录/参考月份', '2026.07附录/参考', '已覆盖']])}</section>`;

const html = `<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场洞察 · Outdoor Rug Intelligence</title><style>
:root{color-scheme:dark;font-family:Inter,"Segoe UI Variable","Segoe UI","Microsoft YaHei",system-ui,sans-serif;--bg:#050505;--surface:#0d0d0d;--surface-soft:rgba(255,255,255,.018);--surface-strong:rgba(255,255,255,.05);--text:#f2f2f0;--muted:#8b8b88;--faint:#5f5f5c;--line:rgba(255,255,255,.095);--line-strong:rgba(255,255,255,.22);--brand:#f4f4f1;--brand-soft:rgba(255,255,255,.05);--warning:#aaa9a2;--warning-soft:rgba(255,255,255,.04);--sidebar:#060606;--cta:#eeeeeb;--cta-text:#080808;--grid-line:rgba(255,255,255,.012);--card-bg:linear-gradient(145deg,rgba(18,18,18,.72),rgba(8,8,8,.8));--shadow:0 28px 80px rgba(0,0,0,.3),inset 0 1px rgba(255,255,255,.025);--display:"Times New Roman","Songti SC","STSong",serif;--mono:"Cascadia Code","SFMono-Regular",Consolas,monospace}
:root[data-theme="light"]{color-scheme:light;--bg:#f4f3ef;--surface:#fafaf7;--surface-soft:rgba(24,24,20,.03);--surface-strong:rgba(24,24,20,.06);--text:#191916;--muted:#6f6e68;--faint:#9a9992;--line:rgba(24,24,20,.12);--line-strong:rgba(24,24,20,.28);--brand:#1b1b18;--brand-soft:rgba(24,24,20,.05);--warning:#65635b;--warning-soft:rgba(24,24,20,.04);--sidebar:#efede7;--cta:#191916;--cta-text:#f8f7f3;--grid-line:rgba(24,24,20,.025);--card-bg:linear-gradient(145deg,rgba(255,255,252,.82),rgba(245,244,239,.9));--shadow:0 18px 48px rgba(24,24,20,.08),inset 0 1px rgba(255,255,255,.6)}
*{box-sizing:border-box}html{min-width:320px;scroll-behavior:smooth;background:var(--bg)}body{position:relative;margin:0;min-height:100vh;color:var(--text);font-size:18px;background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(255,255,255,.105),transparent 68%),radial-gradient(ellipse 48% 38% at 12% 92%,rgba(255,255,255,.035),transparent 72%),linear-gradient(145deg,#020202 0%,#080808 48%,#030303 100%);line-height:1.65}:root[data-theme="light"] body{background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(20,20,18,.08),transparent 68%),linear-gradient(145deg,#faf9f5,#f0efea 52%,#f7f6f2)}body::before{content:"";position:fixed;inset:0;pointer-events:none;opacity:.22;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:72px 72px;mask-image:linear-gradient(to bottom,#000,transparent 82%)}a{color:inherit}.app-shell{position:relative;z-index:1;min-height:100vh}.sidebar{position:fixed;inset:0 auto 0 0;z-index:30;display:flex;width:232px;flex-direction:column;overflow-y:auto;padding:34px 18px 24px;border-right:1px solid var(--line);background:linear-gradient(180deg,rgba(4,4,4,.97),rgba(8,8,8,.9));box-shadow:18px 0 80px rgba(0,0,0,.45)}:root[data-theme="light"] .sidebar{background:linear-gradient(180deg,#f2f0ea,#eae8e1);box-shadow:12px 0 40px rgba(24,24,20,.06)}.brand{display:flex;align-items:center;gap:12px;padding:0 8px;margin-bottom:34px}.brand-mark{display:grid;width:34px;height:34px;place-items:center;border:1px solid var(--line-strong);border-radius:50%;background:radial-gradient(circle at 35% 28%,#f0f0ed,#696966 38%,#080808 70%);box-shadow:0 0 34px rgba(255,255,255,.12)}.brand strong{display:block;font-family:var(--display);font-size:18px;font-weight:400;letter-spacing:.07em}.brand small{display:block;margin-top:3px;color:var(--muted);font-size:13px}.local-badge{display:flex;align-items:center;gap:7px;width:max-content;margin:10px 8px 24px;padding:6px 9px;border:1px solid var(--line);font-family:var(--mono);font-size:13px;letter-spacing:.14em;color:var(--muted)}.local-badge i,.privacy-chip i{width:6px;height:6px;border-radius:50%;background:var(--brand);box-shadow:0 0 12px rgba(255,255,255,.55)}.nav-label{display:block;padding:0 10px 7px;color:var(--muted);font-family:var(--mono);font-size:14px;letter-spacing:.22em}.sidebar nav{display:grid;gap:2px}.sidebar nav a{position:relative;display:flex;align-items:center;gap:11px;padding:11px 10px;color:var(--muted);font-size:15px;letter-spacing:.04em;text-decoration:none}.sidebar nav a:hover,.sidebar nav a.is-active{color:var(--text);background:var(--surface-strong)}.nav-icon{display:grid;width:23px;height:23px;place-items:center;border-radius:50%;font-family:var(--mono);font-size:14px}.sidebar-footer{margin-top:auto;padding:16px 10px 3px;border-top:1px solid var(--line)}.sidebar-footer span,.sidebar-footer strong,.sidebar-footer small{display:block}.sidebar-footer span,.sidebar-footer small{color:var(--muted);font-size:13px}.sidebar-footer strong{margin:7px 0 3px;font-size:15px}.workspace{min-height:100vh;margin-left:232px}.topbar{display:flex;min-height:154px;align-items:center;gap:18px;padding:26px clamp(28px,4vw,64px);border-bottom:1px solid var(--line);background:color-mix(in srgb,var(--bg),var(--surface) 25%)}.eyebrow{display:block;margin-bottom:7px;color:var(--faint);font-family:var(--mono);font-size:13px;letter-spacing:.28em}.topbar h1{margin:0 0 8px;font-family:var(--display);font-size:clamp(44px,5vw,70px);font-weight:400;letter-spacing:-.04em;background:linear-gradient(110deg,#fff 8%,#adada8 55%,#535350 100%);background-clip:text;color:transparent}:root[data-theme="light"] .topbar h1{background:linear-gradient(110deg,#191916,#55544e 55%,#9a9992);background-clip:text;color:transparent}.topbar p{margin:0;color:var(--muted);font-size:16px}.top-actions{display:flex;align-items:center;gap:10px;margin-left:auto}.privacy-chip,.theme-button{border:1px solid var(--line);color:var(--muted);background:var(--surface)}.privacy-chip{display:flex;align-items:center;gap:8px;padding:8px 11px;font-size:14px}.theme-button{display:grid;width:38px;height:38px;place-items:center;cursor:pointer}.content{width:min(1500px,100%);margin:0 auto;padding:32px clamp(28px,4vw,64px) 72px}.scope-notice{display:flex;gap:12px;padding:13px 16px;border:1px solid var(--line-strong);color:var(--muted);background:var(--brand-soft);font-size:16px}.scope-notice b{color:var(--text)}.metrics-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:22px 0}.metric-card{position:relative;min-height:142px;padding:20px;border:1px solid var(--line);background:var(--card-bg);box-shadow:var(--shadow)}.metric-card::before,section::before{content:"";position:absolute;inset:0 0 auto;height:1px;background:linear-gradient(90deg,rgba(255,255,255,.7),transparent)}.metric-label,.metric-value,.metric-note{display:block}.metric-label{color:var(--muted);font-size:15px}.metric-value{margin:15px 0 9px;font-family:var(--display);font-size:clamp(30px,2.6vw,40px);font-weight:400}.metric-note{color:var(--faint);font-size:14px;line-height:1.5}section{position:relative;min-width:0;margin-bottom:22px;padding:23px;border:1px solid var(--line);background:var(--surface);box-shadow:var(--shadow);scroll-margin-top:18px}section h2{margin:0 0 6px;font-family:var(--display);font-size:28px;font-weight:400;letter-spacing:-.01em}section h3{font-size:20px;letter-spacing:.05em}section li{color:var(--muted);font-size:16px;line-height:1.8}section li b{color:var(--text)}details{padding:15px 0;border-top:1px solid var(--line)}summary{cursor:pointer;color:var(--text);font-size:16px;font-weight:700;letter-spacing:.04em}.table-wrap{max-height:620px;margin-top:13px;overflow:auto;border:1px solid var(--line)}table{width:100%;min-width:980px;border-collapse:collapse;font-family:var(--mono);font-size:15px}th,td{padding:9px 11px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th{position:sticky;top:0;z-index:1;color:var(--muted);background:var(--surface);font-size:15px;letter-spacing:.06em}th:first-child,td:first-child{text-align:left}tbody tr:hover{background:var(--surface-soft)}.note{margin-top:18px;padding:13px 16px;border-left:2px solid var(--warning);color:var(--muted);background:var(--warning-soft);font-size:15px}code{padding:2px 5px;color:var(--text);background:var(--surface-strong);font-family:var(--mono)}
@media(max-width:980px){.sidebar{transform:translateX(-100%)}.workspace{margin-left:0}.metrics-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:560px){.topbar{padding:24px 18px}.privacy-chip{display:none}.content{padding:20px 12px 50px}.metrics-grid{grid-template-columns:1fr}.metric-card{min-height:116px}section{padding:17px}.topbar h1{font-size:44px}}
</style></head><body><div class="app-shell"><aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫市场分析系统</small></div></div><div class="local-badge"><i></i> DATA · VERIFIED</div><span class="nav-label">市场分析</span><nav><a href="#dashboard"><span class="nav-icon">总</span>市场总览</a><a href="#definitions"><span class="nav-icon">径</span>数据口径</a><a href="#overall"><span class="nav-icon">整</span>整体市场</a><a href="#pp"><span class="nav-icon">PP</span>塑料地垫</a><a href="#high"><span class="nav-icon">高</span>高客单非PP</a><a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#anomaly"><span class="nav-icon">!</span>参考附录</a><a href="#insights"><span class="nav-icon">策</span>趋势与建议</a><a href="#coverage"><span class="nav-icon">核</span>需求覆盖</a></nav><div class="sidebar-footer"><span>可比数据范围</span><strong>${analysisMonths[0]} — ${analysisMonths[analysisMonths.length - 1]}</strong><small>${analysisMonths.length}个月 · 444万单元格核验</small></div></aside><div class="workspace"><header class="topbar"><div><span class="eyebrow">OUTDOOR RUG MARKET INTELLIGENCE</span><h1>市场总览</h1><p>销量、销售额、双均价 · 小类BSR前100 · 月度MoM与年度YoY</p></div><div class="top-actions"><span class="privacy-chip"><i></i> 源数据只读</span><button class="theme-button" id="theme-toggle" aria-label="切换主题">☀</button></div></header><main class="content"><div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>核心 49 个月（202206-202606）含竞品父ASIN去重口径；2026.07 仅作附录/参考。2026 全年已统一为竞品父ASIN去重口径，与 2025 全市场口径不同。</div></div><div class="metrics-grid"><article class="metric-card"><span class="metric-label">2025 整体销量</span><strong class="metric-value">${fmt(insight.overall2025.sales)}</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoySales)} · 全市场</span></article><article class="metric-card"><span class="metric-label">2025 整体销售额</span><strong class="metric-value">$${fmt(insight.overall2025.revenue / 1000000, 1)}M</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoyRevenue)} · USD</span></article><article class="metric-card"><span class="metric-label">PP 销量贡献</span><strong class="metric-value">${fmt(ppSalesShare2025,1)}%</strong><span class="metric-note">2025 PP销量 / 整体销量</span></article><article class="metric-card"><span class="metric-label">GENIMO PP份额</span><strong class="metric-value">${fmt(insight.genimoPpShare2025,2)}%</strong><span class="metric-note">2025销量份额 · 品牌领先</span></article></div><section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>，不是按月销量重新排名。</li><li>PP标题包含 plastic；高客单排除PP后按材质关键词或价格≥$40。</li><li>同时展示SKU平均标价与销量加权均价（销售额/销量）。</li><li>年度YoY严格使用同周期；2023仅以6-12月对比2022年6-12月。</li><li>月度MOM = 今年X月 vs 去年X月（跨年同月，如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）。</li></ul><p class="note">2026全年已统一为竞品快照（按父ASIN去重，随机保留一条）口径，每月64-94个父商品；与2025全市场口径（1700-2000 SKU）跨年对比存在范围差异，请以可比口径列为准。</p></section>${htmlSections}<section id="anomaly"><h2>六、2026.03-07竞品替换数据附录（父ASIN去重）</h2>${htmlTable(['月份','纳入可比报告','销量','销售额($)','SKU平均标价','销量加权均价'],sourceDiagnostics.filter((r)=>r.month>='202603').map((r)=>[r.month,r.includedInComparableReport?'是':'否',fmt(r.sales),fmt(r.revenue),fmt(r.avgListPrice,2),fmt(r.weightedPrice,2)]))}</section>${insightHtml}</main></div></div><script>(function(){var root=document.documentElement,button=document.getElementById('theme-toggle');var saved='dark';try{saved=localStorage.getItem('market-report-theme')||'dark'}catch(e){}root.dataset.theme=saved;button.textContent=saved==='dark'?'☀':'☾';button.addEventListener('click',function(){var next=root.dataset.theme==='dark'?'light':'dark';root.dataset.theme=next;button.textContent=next==='dark'?'☀':'☾';try{localStorage.setItem('market-report-theme',next)}catch(e){}});var links=[].slice.call(document.querySelectorAll('.sidebar nav a'));var observer=new IntersectionObserver(function(entries){entries.forEach(function(entry){if(entry.isIntersecting){links.forEach(function(a){a.classList.toggle('is-active',a.getAttribute('href')==='#'+entry.target.id)})}})},{rootMargin:'-20% 0px -70% 0px'});document.querySelectorAll('[id]').forEach(function(el){observer.observe(el)})})();</script></body></html>`;

const coverageScopeText = `核心明细覆盖 ${analysisMonths.length} 个月（${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]}）；当前库 ${sourceMeta.effective_sheets || 54} 张有效业务表、${fmt(currentDataRowCount)} 条实际记录、${fmt(verifiedDataCellCount)} 个实际数据单元格（原始Excel导入元数据为 ${fmt(sourceMeta.total_rows || 0)} 行，2026.01-07 后续替换为竞品父体口径）；核心结论截止 202606，2026.07 仅附录/参考。`;
const anomalySectionHtml = `<section id="anomaly"><h2>六、2026.07附录/参考</h2><p class="note">${coverageScopeText} 2026.07 超出核心截止月份 202606，仅作附录/参考。2026 全年已统一为竞品父ASIN去重口径。</p>${htmlTable(['月份', '状态', '销量', '销售额($)', 'SKU平均标价', '加权成交均价'], sourceDiagnostics.filter((row) => row.month > REPORT_CUTOFF).map((row) => [row.month, '附录/参考（> ' + REPORT_CUTOFF + '）', fmt(row.sales), fmt(row.revenue), fmt(row.avgListPrice, 2), fmt(row.weightedPrice, 2)]))}</section>`;
let htmlOutput = html
  .replace(/<div id="dashboard" class="scope-notice"><span>◎<\/span><div>[\s\S]*?<\/div><\/div>/, `<div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>${esc(coverageScopeText)}</div></div>`)
  .replace(/<section id="definitions">[\s\S]*?<\/section>/, `<section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>；2026父体层级取同类候选变体中的最佳可解析名次，销量/销售额仍只取固化代表行，不合计子体；每分类每月最多100条，头中尾和五档均复用该Top100集合。</li><li>PP塑料地垫：标题按不区分大小写的完整单词 <code>plastic</code>（单词边界）筛选，空标题按空字符串；2026同父体任一变体命中即归PP。</li><li>高客单非PP：排除PP父体后的全部商品（SPEC 7.5），不再叠加材质关键词或价格门槛。</li><li>均价同时给出SKU平均标价与销量加权成交均价（销售额/销量）。月度MOM（用户口径）= 今年X月 vs 去年X月同月；月度环比 = 本月 vs 上月；年度表使用同月份集合比较。2026.01-06 已更新为全市场父体级快照（1038-1993父体/月），与2025全市场行级导出（1683-2000行/月）量级一致、可直接参考；2025为行级导出（含变体行、无ASIN列），2026为父体级导出（父ASIN去重），颗粒度与导出日期仍略有差异。</li><li>2025.05（主表导出日 2025-06-19，无ASIN列）源数据存在小类BSR同值重复：BSR=17 重复112行（JONATHAN Y SMB110多变体系列+Smiry）、BSR=23 重复125行、BSR=58 重复156行等（变体行共享父体名次），按小类BSR取前100后全部落入1-20 → 2025.05 中部21-50/尾部51-100为空。因此 2026.05 中部/尾部 MOM（同比2025.05）与 2025.06 中部/尾部环比显示"无对应数据"；2026.05 头部 MOM 的基准为异常100行头部（销量162,797、销售额$6,446,797、加权均价$39.60），数值仅供参考。GENIMO 部分月份分层无在榜商品亦显示"无对应数据"（正常稀疏）。</li></ul><p class="note">${esc(coverageScopeText)}</p></section>`)
  .replace(/<section id="anomaly">[\s\S]*?<\/section>/, anomalySectionHtml)
  .replace('销量、销售额、双均价 · 小类BSR前100 · 月度MoM与年度YoY', '销量、销售额、双均价 · 小类BSR Top100 · 月度YOY与MOM')
  .replace('<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a>', '<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#genimo-products"><span class="nav-icon">Top</span>GENIMO主力ASIN</a>')
  .replace('<a href="#coverage"><span class="nav-icon">核</span>需求覆盖</a>', '<a href="#coverage"><span class="nav-icon">核</span>需求覆盖核对</a>')
  .replace(/<small>[^<]*444万单元格核验<\/small>/, `<small>${analysisMonths.length}个月 · ${fmt(verifiedDataCellCount)}数据单元格</small>`)
  .replace('</main></div></div><script>', coverageHtml + '</main></div></div><script>')
  .replace('<section id="insights">', referenceHtml + '<section id="insights">')
  .replace('<section id="insights">', forecastHtml + '<section id="insights">')
  .replace('<section id="insights">', cohortHtml + '<section id="insights">')
  .replace('<section id="insights">', genimoProductsHtml + '<section id="insights">')
  .replace('</style></head>', '<style>.trend-body{padding:10px 0 2px;color:var(--muted)}.trend-body p{margin:7px 0;font-size:14px}.insight-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.insight-grid article{padding:15px;border:1px solid var(--line);background:var(--surface-soft)}.insight-grid h3{margin:0 0 7px}.insight-grid p{margin:0;color:var(--muted);font-size:14px}@media(max-width:700px){.insight-grid{grid-template-columns:1fr}}</style></head>');

fs.mkdirSync(path.dirname(HTML_PATH), { recursive: true });
fs.writeFileSync(JSON_PATH, JSON.stringify(data, null, 2), 'utf8');
fs.writeFileSync(MD_PATH, md.join('\n'), 'utf8');
fs.writeFileSync(HTML_PATH, htmlOutput, 'utf8');
if (competitorDb) competitorDb.close();
db.close();
console.log('Generated: ' + path.relative(ROOT, JSON_PATH));
console.log('Generated: ' + path.relative(ROOT, MD_PATH));
console.log('Generated: ' + path.relative(ROOT, HTML_PATH));
