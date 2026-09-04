'use strict';

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const ROOT = path.resolve(__dirname, '..');
const DB_PATH = path.resolve(ROOT, process.env.ANALYSIS_DB_PATH || 'data/processed/market.db');
const COMPETITOR_DB_PATH = path.resolve(ROOT, process.env.COMPETITOR_DB_PATH || 'data/processed/competitor_809440.db');
const REPORT_CUTOFF = process.env.ANALYSIS_CUTOFF || '202606'; // SPEC 1.2: 核心截止 202606
const MD_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.md');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');
const QUICK_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-极速版.md');
const PLAN_REFERENCE_PATH = path.resolve(ROOT, '新增参考的材料和内容/销量预测计划部底表-户外地垫.xlsx');

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

function fmtMonth(month) {
  const value = String(month || '');
  return /^\d{6}$/.test(value) ? value.slice(0, 4) + '.' + value.slice(4) : value;
}

function fmtPeriod(period) {
  return String(period || '').replace(/\d{6}/g, (month) => fmtMonth(month));
}

function presentNumber(value) {
  return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
}

function parseBsr(value) {
  if (value === null || value === undefined || value === '') return { rank: null, multi: false };
  const matches = String(value).match(/\d[\d,]*/g) || [];
  const ranks = matches.map((s) => Number(s.replace(/,/g, ''))).filter(Number.isFinite);
  return { rank: ranks.length ? Math.min(...ranks) : null, multi: ranks.length > 1 };
}

// 领导验收基准：计划部 workbook 的 BI 全类目和 BSR Top100 H1 汇总。
// 该 workbook 是参考/核对源，不改变 market.db 的历史明细；两个来源的结果
// 必须并列保留，避免把不可比快照的负向结果直接改写成正数。
function loadLeadershipBenchmark() {
  const empty = {
    available: false,
    sourceFile: path.relative(ROOT, PLAN_REFERENCE_PATH),
    industry: null,
    bsrTop100: null,
    error: fs.existsSync(PLAN_REFERENCE_PATH) ? null : '参考 workbook 不存在',
  };
  if (!fs.existsSync(PLAN_REFERENCE_PATH)) return empty;
  try {
    const workbook = XLSX.readFile(PLAN_REFERENCE_PATH, {
      cellFormula: true,
      cellStyles: true,
      sheetStubs: true,
    });
    const industrySheet = workbook.Sheets['行业大盘数据'];
    const bsrSheet = workbook.Sheets['BSR底表-美国'];
    if (!industrySheet || !bsrSheet) return { ...empty, error: '缺少行业大盘数据或BSR底表-美国工作表' };
    const cellNumber = (sheet, address) => {
      const value = sheet[address] && sheet[address].v;
      return typeof value === 'number' && Number.isFinite(value) ? value : null;
    };
    const monthRows = [];
    for (let row = 4; row <= 9; row++) {
      const monthLabel = industrySheet['G' + row] && industrySheet['G' + row].v;
      const monthNumber = String(monthLabel || '').match(/(\d{1,2})/);
      if (!monthNumber) continue;
      const month = String(Number(monthNumber[1])).padStart(2, '0');
      const baselineSales = cellNumber(industrySheet, 'H' + row);
      const currentSales = cellNumber(industrySheet, 'I' + row);
      monthRows.push({
        month: '2026' + month,
        baselineMonth: '2025' + month,
        baselineSales,
        currentSales,
        momSales: pct(currentSales, baselineSales),
      });
    }
    const baselineSales = monthRows.reduce((sum, row) => sum + (row.baselineSales || 0), 0);
    const currentSales = monthRows.reduce((sum, row) => sum + (row.currentSales || 0), 0);
    const industry = {
      available: monthRows.length === 6 && baselineSales > 0 && currentSales > 0,
      sourceFile: path.relative(ROOT, PLAN_REFERENCE_PATH),
      sheet: '行业大盘数据',
      sourceLabel: 'BI看板-亚马逊后台',
      category: 'Outdoor Rugs',
      metric: 'sales',
      period: '202601-202606 vs 202501-202506',
      baselineSales,
      currentSales,
      growthPct: pct(currentSales, baselineSales),
      formula: 'SUM(I4:I9)/SUM(H4:H9)-1',
      months: monthRows,
      revenueAvailable: false,
      note: '仅提供全类目销量；未提供销售额/均价，不推导缺失指标。',
    };

    const bsrRows = [];
    const columnNumber = (column, row) => cellNumber(bsrSheet, column + row);
    for (let row = 2; row <= 101; row++) {
      const rank = columnNumber('Q', row);
      if (rank === null || rank < 1 || rank > 100) continue;
      const h1_2025 = ['AX', 'AY', 'AZ', 'BA', 'BB', 'BC'].reduce((sum, col) => sum + (columnNumber(col, row) || 0), 0);
      const h1_2026 = ['BJ', 'BK', 'BL', 'BM', 'BN', 'BO'].reduce((sum, col) => sum + (columnNumber(col, row) || 0), 0);
      bsrRows.push({ row, rank, h1_2025, h1_2026 });
    }
    const bsrBaselineSales = bsrRows.reduce((sum, row) => sum + row.h1_2025, 0);
    const bsrCurrentSales = bsrRows.reduce((sum, row) => sum + row.h1_2026, 0);
    const bsrTop100 = {
      available: bsrRows.length === 100 && bsrBaselineSales > 0 && bsrCurrentSales > 0,
      sourceFile: path.relative(ROOT, PLAN_REFERENCE_PATH),
      sheet: 'BSR底表-美国',
      metric: 'sales',
      period: '202601-202606 vs 202501-202506',
      rankRange: 'Q=1..100',
      baselineSales: bsrBaselineSales,
      currentSales: bsrCurrentSales,
      growthPct: pct(bsrCurrentSales, bsrBaselineSales),
      formula: 'SUM(Q=1..100, BJ:BO)/SUM(Q=1..100, AX:BC)-1',
      note: '按原始月度输入独立求和；不使用原表 P 列的漏月/混月公式。',
    };
    return { available: industry.available && bsrTop100.available, sourceFile: path.relative(ROOT, PLAN_REFERENCE_PATH), industry, bsrTop100, error: null };
  } catch (error) {
    return { ...empty, error: String(error && error.message ? error.message : error) };
  }
}

const leadershipBenchmark = loadLeadershipBenchmark();

// SPEC 7.5: PP = 标题完整单词 plastic（单词边界，不区分大小写）；high = 排除 PP 后全部产品
const PLASTIC_WORD_RE = /\bplastic\b/i;
// 2026.09-12 预测和 2027 规划基准（SPEC 1.3，仅作预测/假设参考，非历史实绩）
const FORECAST_2026_Q4 = [
  { month: '202609', sales: 115000, range: '104,000-125,000', rev: 3270000, stage: '旺季结束、需求快速回落' },
  { month: '202610', sales: 99000, range: '89,000-107,000', rev: 3160000, stage: '淡季+秋季促销' },
  { month: '202611', sales: 94000, range: '85,000-102,000', rev: 3870000, stage: '黑五带来销售额修复' },
  { month: '202612', sales: 101000, range: '91,000-110,000', rev: 4120000, stage: '低基数+节日场景支撑' },
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
const FORECAST_PARAMETERS = [
  { parameter: '需求系数', defaultValue: '1.00', effect: '调整销量；情景销量 = 基准销量 × 需求系数' },
  { parameter: '成交价系数', defaultValue: '1.00', effect: '调整价格；情景销售额 = 基准销售额 × 需求系数 × 成交价系数' },
  { parameter: '旺季前置周数', defaultValue: '4-8周', effect: '决定备货和广告启动时间，不改变历史实绩' },
  { parameter: '情景选择', defaultValue: '基准', effect: '可切换保守/基准/进取，触发条件见情景表' },
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

function multiForCategory(row, category) {
  if (!row.bestRankMulti) return Boolean(row.multi);
  return Boolean(row.bestRankMulti[category]);
}

function bsrSourceForCategory(row, category) {
  if (!row.bestRankSources) return row.bsr === null || row.bsr === undefined ? '' : String(row.bsr);
  return row.bestRankSources[category] || '';
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
  const prices = rows.filter((row) => presentNumber(row.price)).map((row) => Number(row.price));
  return {
    skuCount: rows.length,
    pricedSkuCount: prices.length,
    sales,
    revenue,
    avgListPrice: prices.length ? prices.reduce((a, b) => a + b, 0) / prices.length : null,
    weightedPrice: sales ? revenue / sales : null,
  };
}

function addTrends(monthly) {
  const byMonth = new Map(monthly.map((row) => [row.month, row]));
  for (const row of monthly) {
    const lastYear = byMonth.get(String(Number(row.month.slice(0, 4)) - 1) + row.month.slice(4));
    // 最新用户口径：月度MOM/环比 = 今年X月 vs 去年X月（跨年同月）。
    row.momBasis = lastYear ? lastYear.month : null;
    row.momSales = lastYear ? pct(row.sales, lastYear.sales) : null;
    row.momRevenue = lastYear ? pct(row.revenue, lastYear.revenue) : null;
    row.momAvgListPrice = lastYear ? pct(row.avgListPrice, lastYear.avgListPrice) : null;
    row.momWeightedPrice = lastYear ? pct(row.weightedPrice, lastYear.weightedPrice) : null;
    // 跨年同月基准存在但对应分层为空（0 SKU / 0 销量）时，变化无对应数据可比。
    // 典型场景：2025.05 源表小类BSR同值重复（BSR=17 重复112行等），前100全部落入1-20，
    // 导致 2025.05 中部/尾部为空 → 2026.05 中部/尾部跨年同月变化缺失。
    row.momGapReason = lastYear && (lastYear.skuCount === 0 || lastYear.sales === 0)
      ? '基准月 ' + lastYear.month + ' 对应分层为空，跨年同月MOM/环比无对应数据可比'
      : null;
    // 删除历史连续环比字段，避免旧的 chain/previous 比较混入当前交付物。
    delete row.chainBasis;
    delete row.chainSales;
    delete row.chainRevenue;
    delete row.chainAvgListPrice;
    delete row.chainWeightedPrice;
    delete row.chainGapReason;
  }
  return monthly;
}
function periodSummary(monthly, months) {
  const selected = monthly.filter((row) => months.includes(row.month));
  const sales = selected.reduce((sum, row) => sum + row.sales, 0);
  const revenue = selected.reduce((sum, row) => sum + row.revenue, 0);
  const pricedSkuCount = selected.reduce((sum, row) => sum + (row.pricedSkuCount || 0), 0);
  const skuWeightedPriceSum = selected.reduce((sum, row) => sum + (row.avgListPrice || 0) * (row.pricedSkuCount || 0), 0);
  return {
    months: selected.length,
    sales,
    revenue,
    pricedSkuCount,
    avgListPrice: pricedSkuCount ? skuWeightedPriceSum / pricedSkuCount : null,
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
    const timeComparable = prior.months > 0 && prior.months === comparableCurrent.months;
    const scopeComparable = timeComparable && year !== '2026';
    return {
      year,
      period: currentMonths[0] + '-' + currentMonths[currentMonths.length - 1],
      ...current,
      comparison: timeComparable ? comparableCurrentMonths[0] + '-' + comparableCurrentMonths[comparableCurrentMonths.length - 1]
        + ' vs ' + priorMonths[0] + '-' + priorMonths[priorMonths.length - 1] : null,
      yoySales: timeComparable ? pct(comparableCurrent.sales, prior.sales) : null,
      yoyRevenue: timeComparable ? pct(comparableCurrent.revenue, prior.revenue) : null,
      yoyAvgListPrice: timeComparable ? pct(comparableCurrent.avgListPrice, prior.avgListPrice) : null,
      yoyWeightedPrice: timeComparable ? pct(comparableCurrent.weightedPrice, prior.weightedPrice) : null,
      timeComparable,
      scopeComparable,
      scopeNote: year === '2026'
        ? '方向性参考：时间同周期，但2025为含ASIN/父ASIN的行级变体导出，2026为父ASIN去重快照；统计单元不同，不构成严格同口径同比'
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
  // 领导反馈（2026-08-31）：趋势结论必须以 2026 实绩为主，2025 只作为同比基准。
  const benchmark = c.monthly.find((row) => row.month === '202602');
  const baseline = [...c.monthly].reverse().find((row) => row.month <= REPORT_CUTOFF) || c.monthly[c.monthly.length - 1];
  const annual2026 = c.annual.find((row) => row.year === '2026');
  const top2026 = c.bsrTop100.annual.find((row) => row.year === '2026');
  const groups2026 = c.bsrGroups.annual.filter((row) => row.year === '2026');
  const groupLine = groups2026.map((row) => `${row.segment}销量方向变化 ${fmtPct(row.yoySales)}、销售额方向变化 ${fmtPct(row.yoyRevenue)}`).join('；');
  const peak2026 = [...c.monthly.filter((row) => row.month.startsWith('2026'))]
    .sort((left, right) => right.sales - left.sales)[0];
  const out = [`### ${label}趋势分析`, ''];
  if (category === 'overall' && leadershipBenchmark.available && leadershipBenchmark.industry) {
    const benchmark = leadershipBenchmark.industry;
    const bsrBenchmark = leadershipBenchmark.bsrTop100;
    const bsrText = bsrBenchmark && bsrBenchmark.available
      ? `；同一 workbook 的 BSR Top100 独立重算为 ${fmt(bsrBenchmark.currentSales)} vs ${fmt(bsrBenchmark.baselineSales)}，销量方向变化 ${fmtPct(bsrBenchmark.growthPct)}`
      : '';
    out.push(`- 领导验收口径（计划部 BI 全类目）：${fmtPeriod(benchmark.period)}销量 ${fmt(benchmark.currentSales)} vs ${fmt(benchmark.baselineSales)}，销量方向变化 ${fmtPct(benchmark.growthPct)}${bsrText}。该参考表仅提供销量，不推导销售额或均价。`);
  }
  if (annual2026) {
    out.push(`- 2026.01-06核心实绩：销量 ${fmt(annual2026.sales)}、销售额 $${fmt(annual2026.revenue)}；按现有混合统计单元相对2025同期的方向变化为 ${fmtPct(annual2026.yoySales)} / ${fmtPct(annual2026.yoyRevenue)}，SKU平均标价/加权成交均价方向变化为 ${fmtPct(annual2026.yoyAvgListPrice)} / ${fmtPct(annual2026.yoyWeightedPrice)}，不可解释为严格同口径同比。`);
  }
  if (top2026 && annual2026) {
    out.push(`- 2026.01-06 BSR前100贡献销量 ${fmt(top2026.sales)}（占同期${fmt(top2026.sales / annual2026.sales * 100, 1)}%），销售额占比 ${fmt(top2026.revenue / annual2026.revenue * 100, 1)}%；相对2025行代理池的销量/销售额方向变化为 ${fmtPct(top2026.yoySales)} / ${fmtPct(top2026.yoyRevenue)}。`);
  }
  if (groupLine) out.push(`- 2026.01-06头中尾分层：${groupLine}。`);
  if (benchmark) {
    const expected = category === 'overall'
      ? '（整体市场全量快照复核值约 -14.8% / -20.5%）'
      : '';
    out.push(`- ${fmtMonth(benchmark.month)}：月度MOM/环比按跨年同月口径（${fmtMonth(benchmark.month)} vs ${fmtMonth(benchmark.momBasis)}）销量 ${fmtPct(benchmark.momSales)}、销售额 ${fmtPct(benchmark.momRevenue)}${expected}。`);
  }
  if (baseline && baseline.month !== '202602') out.push(`- 2026核心截止月 ${fmtMonth(baseline.month)}：月度MOM/环比按跨年同月口径（${fmtMonth(baseline.month)} vs ${fmtMonth(baseline.momBasis)}）销量 ${fmtPct(baseline.momSales)}、销售额 ${fmtPct(baseline.momRevenue)}。`);
  if (peak2026) out.push(`- 2026.01-06核心月份中，${fmtMonth(peak2026.month)}销量最高，为 ${fmt(peak2026.sales)} 件；该峰值用于安排2027旺季前4-8周的补货、广告与新品测试。`);
  const scopeAnchor = c.monthly.find((row) => row.month === '202604');
  if (scopeAnchor && baseline) out.push(`- 口径提示：2026.01-06为全市场父体级快照（每月1038-1993个父体），2025为含ASIN/父ASIN的行级导出（每月1683-2000行，含变体行）。两者量级接近不等于统计单元一致，跨年变化仅作方向性参考；2026.07为94父体小样本，只合并展示，不参与同比、环比和累计。`);
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
        bestRankMulti: { overall: false, pp: false, high: false, genimo: false },
        bestRankSources: { overall: '', pp: '', high: '', genimo: '' },
      };
      const parsed = parseBsr(row.bsr);
      const rank = parsed.rank;
      const plastic = PLASTIC_WORD_RE.test(String(row.title || ''));
      const genimo = String(row.brand || '').trim().toLowerCase() === 'genimo';
      profile.hasPlastic = profile.hasPlastic || plastic;
      profile.hasGenimo = profile.hasGenimo || genimo;
      function keepBest(field, eligible) {
        if (!eligible || rank === null) return;
        if (profile.bestRanks[field] === null || rank < profile.bestRanks[field]) {
          profile.bestRanks[field] = rank;
          profile.bestRankMulti[field] = parsed.multi;
          profile.bestRankSources[field] = String(row.bsr || '');
        } else if (rank === profile.bestRanks[field] && parsed.multi) {
          profile.bestRankMulti[field] = true;
          profile.bestRankSources[field] = String(row.bsr || profile.bestRankSources[field] || '');
        }
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
      bestRankMulti: profile ? profile.bestRankMulti : undefined,
      bestRankSources: profile ? profile.bestRankSources : undefined,
    };
  }));
}

function rowsForCategory(month, category) {
  return rawByMonth.get(month)
    .filter((row) => classify(row, category))
    .map((row) => ({
      ...row,
      rank: rankForCategory(row, category),
      bsrMulti: multiForCategory(row, category),
      bsrSource: bsrSourceForCategory(row, category),
    }));
}

function bsrPoolQuality(month, rows, top100) {
  const identifiedRows = top100.filter((row) => String(row.parent || '').trim() || String(row.asin || '').trim()).length;
  const listingKeys = top100.map((row) => listingKey(row)).filter((key) => !key.startsWith('row-'));
  const distinctListingKeys = new Set(listingKeys).size;
  const duplicateListingRows = Math.max(0, listingKeys.length - distinctListingKeys);
  const historicalRowLevel = month < '202601';
  const distinctRanks = new Set(top100.map((row) => row.rank)).size;
  return {
    month,
    eligibleRows: rows.filter((row) => row.rank !== null && row.rank >= 1 && row.rank <= 100).length,
    selectedRows: top100.length,
    identifiedRows,
    identifierCoveragePct: top100.length ? identifiedRows / top100.length * 100 : null,
    distinctListingKeys,
    duplicateListingRows,
    distinctRanks,
    repeatedRankRows: Math.max(0, top100.length - distinctRanks),
    multiValueRows: top100.filter((row) => row.bsrMulti).length,
    multiValuePct: top100.length ? top100.filter((row) => row.bsrMulti).length / top100.length * 100 : null,
    statisticalUnit: !historicalRowLevel && identifiedRows === top100.length && duplicateListingRows === 0
      ? '父体/ASIN独立Listing'
      : identifiedRows === top100.length
        ? '含ASIN标识的行级/变体展开'
        : '源表行代理（无Listing标识）',
    strictListingPool: !historicalRowLevel && top100.length > 0 && identifiedRows === top100.length && duplicateListingRows === 0,
  };
}

const categories = {};
const bsrMultiValueAudit = [];
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const monthly = [];
  const bsrTop100 = [];
  const bsrSegments = [];
  const bsrGroups = [];
  const bsrQuality = [];
  for (const month of analysisMonths) {
    const rows = rowsForCategory(month, category);
    monthly.push({ month, ...summarize(rows) });
    const top100 = top100Rows(rows); // 每类别每月 BSR前100 独立 Listing ≤ 100
    for (const row of top100.filter((item) => item.bsrMulti)) {
      bsrMultiValueAudit.push({
        category,
        month,
        listingKey: listingKey(row),
        parent: row.parent || '',
        asin: row.asin || '',
        rank: row.rank,
        sourceBsr: row.bsrSource || String(row.bsr || ''),
        multiValue: true,
      });
    }
    bsrTop100.push({ month, ...summarize(top100) });
    bsrQuality.push(bsrPoolQuality(month, rows, top100));
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
    bsrTop100: { monthly: addTrends(bsrTop100), annual: buildAnnual(bsrTop100), quality: bsrQuality },
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
  includedInLongTermConclusion: month <= REPORT_CUTOFF, // 核心结论截止 202606；202607 仅合并展示
  bestBsrEnrichedFromCompetitorRaw: rankEnrichedMonths.includes(month),
  ...summarize(rawByMonth.get(month)),
}));

// 领导反馈：整体市场趋势必须把 2026.01-07 合并到同一张趋势表。
// 2026.07 仍为 94 父体小样本，因此只合并展示原始规模，不计算跨口径 MOM/环比。
const overallMarketTrend2026 = sourceMonths.filter((month) => month >= '202601' && month <= '202607').map((month) => {
  const diagnostic = sourceDiagnostics.find((row) => row.month === month);
  const core = categories.overall.monthly.find((row) => row.month === month);
  return {
    month,
    skuCount: diagnostic.skuCount,
    sales: diagnostic.sales,
    revenue: diagnostic.revenue,
    avgListPrice: diagnostic.avgListPrice,
    weightedPrice: diagnostic.weightedPrice,
    momBasis: core ? core.momBasis : null,
    momSales: core ? core.momSales : null,
    momRevenue: core ? core.momRevenue : null,
    coreComparable: Boolean(core),
    scopeStatus: core
      ? '核心全市场父体快照'
      : '94父体小样本；合并展示，不参与同比/环比和累计',
  };
});

const genimoProducts = new Map();
const core2026Months = analysisMonths.filter((month) => month >= '202601' && month <= REPORT_CUTOFF);
for (const month of core2026Months) {
  for (const row of rawByMonth.get(month).filter((r) => classify(r, 'genimo'))) {
    const key = listingKey(row);
    const item = genimoProducts.get(key) || { listingKey: key, parent: row.parent || '', asin: row.asin || '', title: row.title, sales: 0, revenue: 0, months: 0, latestPrice: null };
    item.sales += Number(row.sales || 0);
    item.revenue += Number(row.revenue || 0);
    item.months++;
    item.parent = row.parent || item.parent;
    item.asin = row.asin || item.asin;
    item.title = row.title || item.title;
    item.latestPrice = presentNumber(row.price) ? Number(row.price) : item.latestPrice;
    genimoProducts.set(key, item);
  }
}

const ppListingDetails = core2026Months.flatMap((month) => rowsForCategory(month, 'pp')
  .sort((left, right) => (left.rank ?? Number.MAX_SAFE_INTEGER) - (right.rank ?? Number.MAX_SAFE_INTEGER)
    || listingKey(left).localeCompare(listingKey(right)))
  .map((row) => ({
    month,
    listingKey: listingKey(row),
    parent: row.parent || '',
    asin: row.asin || '',
    title: row.title || '',
    sales: Number(row.sales || 0),
    revenue: Number(row.revenue || 0),
    price: presentNumber(row.price) ? Number(row.price) : null,
    rank: row.rank,
    bsrMulti: Boolean(row.bsrMulti),
    sourceBsr: row.bsrSource || String(row.bsr || ''),
  })));

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
  leadershipBenchmark,
  dataQuality: {
    competitorDatabaseAvailable: Boolean(competitorDb),
    bestBsrEnrichedMonths: rankEnrichedMonths,
    representativeMetrics: '2026.01-07 canonical representative = smallest parsable 小类BSR; tie-break by complete sales/revenue fields, then source row; child rows are never summed',
    ranking: '2026.01-07 category tiers use the best parsable 小类BSR among qualifying variants in the same parent family',
    top100Cap: 'each category/month is deterministically capped at 100 listings; all tier tables reuse that exact Top100 pool',
    bsrMultiValueAudit,
    historicalBsrWarning: '2022-2025 exports contain ASIN/父ASIN as rich-text hyperlinks after display-value restoration; BSR Top100 remains a row-level/variant pool and is not necessarily 100 independent parent listings. Cross-year BSR changes are directional only.',
    historicalBsrTop100Quality: categories.overall.bsrTop100.quality.filter((row) => row.month < '202601'),
  },
  definitions: {
    pp: "标题按不区分大小写的完整单词 plastic（单词边界）筛选；2026父体任一变体命中即归PP，空标题按空字符串",
    high: "排除 PP 父体后的全部商品（SPEC 7.5：其余全部归入高客单非PP，不再叠加材质关键词或价格门槛）",
    bsrTop100: '2022-2025: deterministic row proxy because ASIN/父ASIN are absent; 2026: independent parent/ASIN listings using the minimum numeric 小类BSR among qualifying variants; cap at 100',
    bsrGroups: 'head = 1-20, middle = 21-50, tail = 51-100; groups do not overlap',
    avgListPrice: 'simple average of non-null, non-blank, numeric SKU list prices; missing prices are excluded rather than treated as zero',
    weightedPrice: '月销售额 / 月销量',
    momMonthly: '月度MOM/环比（最新用户口径）：今年X月 vs 去年X月同月（跨年同月）',
    annualYoY: '年度同周期数值对比；2026.01-06 vs 2025.01-06 因统计单元从行级变为父体级，仅作方向性参考，不是严格同口径同比',
  },
  categories,
  overallMarketTrend2026,
  sourceDiagnostics,
  ppListingDetailsPeriod: core2026Months[0] + '-' + core2026Months[core2026Months.length - 1],
  ppListingDetails,
  genimoTopProductsPeriod: '202601-202606',
  genimoTopProducts: [...genimoProducts.values()].sort((a, b) => b.sales - a.sales).slice(0, 20),
  forecast2026Q4: FORECAST_2026_Q4,
  forecast2027Monthly: FORECAST_2027_MONTHLY,
  forecast2027Scenarios: FORECAST_2027_SCENARIOS,
  forecastParameters: FORECAST_PARAMETERS,
};

function annualRow(category, year) {
  return categories[category].annual.find((row) => row.year === year);
}

function peakMonth(category, year) {
  return [...categories[category].monthly.filter((row) => row.month.startsWith(year))]
    .sort((a, b) => b.sales - a.sales)[0];
}

const pp2026Rows = core2026Months
  .flatMap((month) => rawByMonth.get(month).filter((row) => classify(row, 'pp')));
const genimoPp2026Rows = pp2026Rows.filter((row) => classify(row, 'genimo'));
const pp2026Summary = summarize(pp2026Rows);
const genimoPp2026Summary = summarize(genimoPp2026Rows);
const overall2026 = annualRow('overall', '2026');
const pp2026 = annualRow('pp', '2026');
const high2026 = annualRow('high', '2026');
const genimo2026 = annualRow('genimo', '2026');
data.insights = {
  period: '202601-202606',
  overall2026,
  pp2026,
  high2026,
  genimo2026,
  overallTop1002026: categories.overall.bsrTop100.annual.find((row) => row.year === '2026'),
  ppTop1002026: categories.pp.bsrTop100.annual.find((row) => row.year === '2026'),
  highTop1002026: categories.high.bsrTop100.annual.find((row) => row.year === '2026'),
  genimoTop1002026: categories.genimo.bsrTop100.annual.find((row) => row.year === '2026'),
  ppPeak2026: peakMonth('pp', '2026'),
  ppSalesShare2026: overall2026.sales ? pp2026.sales / overall2026.sales * 100 : null,
  ppRevenueShare2026: overall2026.revenue ? pp2026.revenue / overall2026.revenue * 100 : null,
  genimoPpShare2026: pp2026Summary.sales ? genimoPp2026Summary.sales / pp2026Summary.sales * 100 : null,
  genimoPpRevenueShare2026: pp2026Summary.revenue ? genimoPp2026Summary.revenue / pp2026Summary.revenue * 100 : null,
  genimoJune2026: categories.genimo.monthly.find((row) => row.month === '202606'),
  genimoJuneGroups2026: categories.genimo.bsrGroups.monthly.filter((row) => row.month === '202606'),
  genimoTopProduct2026: data.genimoTopProducts[0] || null,
};

function mdMonthly(rows) {
  const out = ['| 月份 | MOM/环比基准月份 | SKU数 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | MOM/环比销量 | MOM/环比销售额 | MOM/环比标价 | MOM/环比成交均价 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${fmtMonth(r.month)} | ${r.momBasis ? fmtMonth(r.momBasis) : '-'} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.momSales)} | ${fmtPct(r.momRevenue)} | ${fmtPct(r.momAvgListPrice)} | ${fmtPct(r.momWeightedPrice)} |`);
  return out.join('\n');
}
function mergedTrendPct(row, field) {
  return row.coreComparable ? fmtPct(row[field]) : '不适用（口径不同）';
}
function mdOverallMarketTrend2026(rows) {
  const out = ['| 月份 | 口径状态 | MOM/环比基准月份 | Listing数 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | MOM/环比销量 | MOM/环比销售额 |',
    '|---|---|---|---:|---:|---:|---:|---:|---:|---:|'];
  for (const row of rows) out.push(`| ${fmtMonth(row.month)} | ${row.scopeStatus} | ${row.momBasis ? fmtMonth(row.momBasis) : '-'} | ${fmt(row.skuCount)} | ${fmt(row.sales)} | ${fmt(row.revenue)} | ${fmt(row.avgListPrice, 2)} | ${fmt(row.weightedPrice, 2)} | ${mergedTrendPct(row, 'momSales')} | ${mergedTrendPct(row, 'momRevenue')} |`);
  return out.join('\n');
}

function mdReplacementMetadata(rows) {
  const out = ['| 月份 | 主表替换前行数 | 快照原始行数 | 当前父体行数 | 源库SHA-256 | 应用时间 |',
    '|---|---:|---:|---:|---|---|'];
  for (const row of rows) out.push(`| ${fmtMonth(row.month)} | ${fmt(row.base_imported_rows)} | ${fmt(row.source_raw_rows)} | ${fmt(row.replacement_rows)} | ${row.source_sha256 || '-'} | ${row.applied_at || '-'} |`);
  return out.join('\n');
}

function mdPpListingDetails(rows) {
  const out = ['| 月份 | Listing键 | 父ASIN | 代表ASIN | 小类BSR | 多值解析 | 销量 | 销售额($) | 价格($) | 商品标题 |',
    '|---|---|---|---|---:|---|---:|---:|---:|---|'];
  for (const row of rows) out.push(`| ${fmtMonth(row.month)} | ${row.listingKey || '-'} | ${row.parent || '-'} | ${row.asin || '-'} | ${row.rank === null ? '-' : fmt(row.rank)} | ${row.bsrMulti ? '是' : '否'} | ${fmt(row.sales)} | ${fmt(row.revenue)} | ${fmt(row.price, 2)} | ${String(row.title || '').replace(/\|/g, '\\|')} |`);
  return out.join('\n');
}

function mdBsrMultiValueAudit(rows) {
  const out = ['| 类别 | 月份 | Listing键 | 父ASIN | ASIN | 采用名次 | 源小类BSR | 多值标记 |',
    '|---|---|---|---|---|---:|---|---|'];
  for (const row of rows) {
    const sourceBsr = String(row.sourceBsr || '').replace(/\s+/g, ' ').replace(/\|/g, '\\|').trim();
    out.push(`| ${row.category} | ${fmtMonth(row.month)} | ${row.listingKey || '-'} | ${row.parent || '-'} | ${row.asin || '-'} | ${fmt(row.rank)} | ${sourceBsr} | 是 |`);
  }
  return out.join('\n');
}
function scopeLabel(row) {
  if (row.scopeComparable) return '严格同口径';
  if (row.scopeNote) return row.scopeNote;
  return row.timeComparable ? '同周期；口径未确认' : '-';
}
function mdAnnual(rows) {
  const out = ['| 年份/数据周期 | YoY比较周期 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY标价 | YOY成交均价 | 同比范围一致性 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|'];
  for (const r of rows) out.push(`| ${r.year} (${fmtPeriod(r.period)}) | ${r.comparison ? fmtPeriod(r.comparison) : '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyAvgListPrice)} | ${fmtPct(r.yoyWeightedPrice)} | ${scopeLabel(r)} |`);
  return out.join('\n');
}

// 分层表的跨年同月MOM/环比单元格：基准分层为空时显示"无对应数据"而非留空或'-'
function segPct(row, field, gapField) {
  const value = row[field];
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return row[gapField] ? '无对应数据' : '-';
  }
  return fmtPct(value);
}

function mdSegments(rows) {
  const out = ['| 月份 | BSR分层 | MOM/环比基准月份 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | MOM/环比销量 | MOM/环比销售额 | MOM/环比成交均价 |',
    '|---|---|---|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${fmtMonth(r.month)} | ${r.segment} | ${r.momBasis ? fmtMonth(r.momBasis) : '-'} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${segPct(r, 'momSales', 'momGapReason')} | ${segPct(r, 'momRevenue', 'momGapReason')} | ${segPct(r, 'momWeightedPrice', 'momGapReason')} |`);
  const hasGap = rows.some((r) => r.momGapReason);
  if (hasGap) out.push('', '> 注：“无对应数据”表示跨年同月基准分层为空（0条Listing），MOM/环比无法计算。整体市场与高客单的 2026.05 中部/尾部缺失源于 2025.05 源数据小类BSR同值重复（详见一、口径说明）；GENIMO 部分月份分层无在榜商品属正常稀疏。');
  return out.join('\n');
}
function mdAnnualSegments(rows) {
  const out = ['| 年份/数据周期 | BSR分层 | YoY比较周期 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY成交均价 | 同比范围一致性 |',
    '|---|---|---|---:|---:|---:|---:|---:|---:|---:|---|'];
  for (const r of rows) out.push(`| ${r.year} (${fmtPeriod(r.period)}) | ${r.segment} | ${r.comparison ? fmtPeriod(r.comparison) : '-'} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyWeightedPrice)} | ${scopeLabel(r)} |`);
  return out.join('\n');
}

const labels = { overall: '整体市场', pp: 'PP塑料地垫（标题完整单词 plastic）', high: '非PP高客单产品', genimo: 'GENIMO品牌' };
const overall202505Head = categories.overall.bsrGroups.monthly.find((row) => row.month === '202505' && row.segment === '头部（1-20）');
const overall202505HeadEvidence = overall202505Head
  ? `销量${fmt(overall202505Head.sales)}、销售额$${fmt(overall202505Head.revenue)}、加权均价$${fmt(overall202505Head.weightedPrice, 2)}`
  : '该月无可用头部汇总';
const md = ['# 户外地垫市场分析报告（优化版）', '',
  `> 分析范围：核心明细为 ${fmtMonth(analysisMonths[0])}-${fmtMonth(analysisMonths[analysisMonths.length - 1])}，共 ${analysisMonths.length} 个月；整体市场趋势将 2026.01-07 合并展示。2026.01-06 为全市场父体级快照（1038-1993父体/月），2026.07 为94父体小样本，只展示规模，不参与同比/环比和累计。`, '',
  '## 一、口径说明', '',
  '- 小类前100依据源字段 `小类BSR`。2026同父体优先取最小可解析名次；同名次优先销量/销售额字段完整行，再按源表顺序稳定决胜，代表行销量/销售额不重复相加。每分类每月Top100最多100条，所有分层复用同一Top100集合。',
  '- 2022-2025源表的ASIN/父ASIN以富文本超链接保存，已恢复为显示值；历史BSR Top100仍是行级/变体池，存在同父体重复时不能直接证明是100个独立Listing；同一父体和同一名次重复情况按月写入数据JSON质量诊断。2026为父体/ASIN Listing池，因此跨年BSR变化仅作方向性参考。',
  '- PP：标题按不区分大小写的完整单词 `plastic`（单词边界）筛选，NULL按空字符串处理；不含 `plastics` 等扩展词。2026同父体任一变体命中即将该父体归入PP。',
  '- 高客单非PP：排除PP父体后的全部商品（SPEC 7.5，不再叠加材质关键词或价格门槛）。',
  '- 同时提供SKU平均标价和销量加权均价（销售额/销量）；SKU平均标价仅统计非空、可解析的价格，缺失值不按0计入。',
  '- 月度MOM/环比（用户口径）= 今年X月 vs 去年X月同月（如 2025.01 vs 2024.01）；本次交付不计算或展示本月 vs 上月的连续环比；年度YOY = 年度同周期对比。',
  '- 领导验收主基准：计划部参考 workbook「行业大盘数据」的 BI 全类目 Outdoor Rugs，独立按 `SUM(I4:I9)/SUM(H4:H9)-1` 计算 2026.01-06 相对 2025.01-06 销量方向；该表只有销量字段，不推导销售额或均价。BSR Top100 另按 Q=1..100 的 AX:BC 对 BJ:BO 原始月度输入独立求和。',
  `- 2025.05（主表导出日 2025-06-19，ASIN/父ASIN为富文本超链接）源数据存在小类BSR同值重复：BSR=17 重复112行（JONATHAN Y SMB110多变体系列+Smiry）、BSR=23 重复125行、BSR=58 重复156行等（变体行共享父体名次），按小类BSR取前100后全部落入1-20 → 2025.05 中部21-50/尾部51-100为空。因此 2026.05 中部/尾部跨年同月 MOM/环比显示“无对应数据”；2026.05 头部 MOM 的基准为上述异常100行头部（${overall202505HeadEvidence}），数值仅供参考，不可解读为真实头部同比。GENIMO 部分月份分层无在榜商品亦显示“无对应数据”（正常稀疏，非数据错误）。`,
  '- BSR头部/中部/尾部分别为1-20、21-50、51-100；五档明细为1-5、6-10、11-20、21-50、51-100，区间不重叠。',
  '- 年度数值使用同月份集合比较；2023对2022仅比较6-12月。2025是含ASIN/父ASIN的行级导出（含变体行），2026是父ASIN去重快照；即使月度量级接近，统计单元仍不同，所以2025→2026只标为方向性参考，不构成严格同比。', ''];
md.push('### BSR多值解析审计', '',
  '> 下表保留进入各分类Top100集合且源小类BSR包含多个数值的Listing标记；采用可解析最小名次，源字符串保留用于复核。', '',
  mdBsrMultiValueAudit(data.dataQuality.bsrMultiValueAudit), '');

let sectionNo = 2;
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const c = categories[category];
  md.push(`## ${sectionNo++}、${labels[category]}`, '');
  if (category === 'overall') {
    const lead = leadershipBenchmark.industry;
    const leadBsr = leadershipBenchmark.bsrTop100;
    md.push('### 领导验收口径（计划部 BI 全类目）', '',
      lead && lead.available
        ? `- ${fmtPeriod(lead.period)}销量 ${fmt(lead.currentSales)} vs ${fmt(lead.baselineSales)}，按公式“${lead.formula}”独立重算为 **${fmtPct(lead.growthPct)}**。来源：${lead.sourceLabel} / ${lead.sheet}；该表仅提供销量，不推导销售额或均价。`
        : `- 领导验收 workbook 未能读取：${leadershipBenchmark.error || '未知错误'}。`,
      leadBsr && leadBsr.available
        ? `- 同一 workbook 的 BSR Top100（${leadBsr.rankRange}）销量 ${fmt(leadBsr.currentSales)} vs ${fmt(leadBsr.baselineSales)}，按原始月度列独立重算为 **${fmtPct(leadBsr.growthPct)}**；不使用原表 P 列漏月/混月公式。`
        : '',
      '> 以上是领导验收主基准；下方 market.db 全量快照仍保留实际导入结果。由于 2025 为含ASIN/父ASIN的行级/含变体口径、2026 为父体去重快照，负向差异只能作为不可比明细参考，不能与上述 BI 基准混列。', '',
      '### 2026.01-07整体市场趋势（合并展示）', '', mdOverallMarketTrend2026(overallMarketTrend2026), '',
      '> 2026.07只有94个父体，是小范围样本；已按领导反馈并入同一趋势表，但不与1-6月直接计算同比、环比或累计。', '',
      '### 2026数据替换审计记录', '', mdReplacementMetadata(replacementMetadata), '',
      '> 每月替换记录同时保存在 market.db 的 analysis_replacements 表；源库SHA-256用于确认七个月均来自同一次确定性快照构建。', '');
  }
  md.push(trendAnalysis(c, category, labels[category]), '### 月度指标、年度YOY与用户定义MOM/环比', '', mdMonthly(c.monthly), '',
    '### 年度/同周期汇总', '', mdAnnual(c.annual), '', '### 小类BSR前100汇总（BSR 1-100）', '', mdMonthly(c.bsrTop100.monthly), '',
    '### 小类BSR前100年度/同周期汇总', '', mdAnnual(c.bsrTop100.annual), '', '### 小类BSR头部/中部/尾部（月度）', '', mdSegments(c.bsrGroups.monthly), '',
    '### 小类BSR头部/中部/尾部（年度）', '', mdAnnualSegments(c.bsrGroups.annual), '', '### 小类BSR五档分层（月度）', '', mdSegments(c.bsrSegments.monthly), '',
    '### 小类BSR五档分层（年度）', '', mdAnnualSegments(c.bsrSegments.annual), '');
  if (category === 'pp') {
    const latestPpRows = data.ppListingDetails.filter((row) => row.month === REPORT_CUTOFF);
    md.push('### PP独立Listing明细（核心截止月）', '',
      `> 完整JSON保留 ${fmtPeriod(data.ppListingDetailsPeriod)} 共 ${fmt(data.ppListingDetails.length)} 条月度Listing记录；下表展示核心截止月 ${fmtMonth(REPORT_CUTOFF)} 的 ${fmt(latestPpRows.length)} 条独立Listing。`, '',
      mdPpListingDetails(latestPpRows), '');
  }
}

// 页面外壳中的旧占位模板仍使用2025字段名；本地别名只用于模板求值，最终JSON和替换后的HTML不暴露旧口径。
const insight = {
  ...data.insights,
  overall2025: data.insights.overall2026,
  pp2025: data.insights.pp2026,
  high2025: data.insights.high2026,
  genimo2025: data.insights.genimo2026,
  ppPeak2025: data.insights.ppPeak2026,
  genimoPpShare2025: data.insights.genimoPpShare2026,
  genimoPpRevenueShare2025: data.insights.genimoPpRevenueShare2026,
};
const genimoJuneHead = insight.genimoJuneGroups2026.find((row) => row.segment === '头部（1-20）');
const genimoJuneMiddle = insight.genimoJuneGroups2026.find((row) => row.segment === '中部（21-50）');
const genimoJuneTail = insight.genimoJuneGroups2026.find((row) => row.segment === '尾部（51-100）');
const genimoJuneHeadSalesShare = insight.genimoJune2026.sales ? genimoJuneHead.sales / insight.genimoJune2026.sales * 100 : null;
const ppSalesShare2025 = insight.ppSalesShare2026;

// 六、父体进退 Cohort
md.push('', '## 六、父体进退（Cohort）', '',
  'BSR前100按父ASIN（优先）或ASIN统计的留存、退出、新进入及头/中/尾迁移（SPEC 7.6/验收23）。比较周期：2026.01 vs 2026.06（核心分析首尾月；主源历史 ASIN 为富文本恢复值，2026 使用父体替换快照）。',
  '',
  `- **整体市场**：前100父体池从 ${categories.overall.cohort.fromParents} 变为 ${categories.overall.cohort.toParents}；留存 ${categories.overall.cohort.retained}、退出 ${categories.overall.cohort.exited}、新进入 ${categories.overall.cohort.entered}。层间迁移：${Object.entries(categories.overall.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **PP塑料**：前100父体池从 ${categories.pp.cohort.fromParents} 变为 ${categories.pp.cohort.toParents}；留存 ${categories.pp.cohort.retained}、退出 ${categories.pp.cohort.exited}、新进入 ${categories.pp.cohort.entered}。层间迁移：${Object.entries(categories.pp.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **高客单非PP**：前100父体池从 ${categories.high.cohort.fromParents} 变为 ${categories.high.cohort.toParents}；留存 ${categories.high.cohort.retained}、退出 ${categories.high.cohort.exited}、新进入 ${categories.high.cohort.entered}。层间迁移：${Object.entries(categories.high.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  `- **GENIMO**：前100父体池从 ${categories.genimo.cohort.fromParents} 变为 ${categories.genimo.cohort.toParents}；留存 ${categories.genimo.cohort.retained}、退出 ${categories.genimo.cohort.exited}、新进入 ${categories.genimo.cohort.entered}。层间迁移：${Object.entries(categories.genimo.cohort.migration).map(([k, v]) => k + '=' + v).join('、')}。`,
  '', '> 注：2025年主源 ASIN/父ASIN已恢复，但其行级/变体口径与2026父体替换快照不同；跨年父体进退不作为严格同比结论。',
  '');

// 七、GENIMO 2026累计Top父体
md.push('', '## 七、GENIMO 2026.01-06累计Top父体', '', '| 排名 | 父体/Listing | 代表ASIN | 累计销量 | 累计销售额($) | 覆盖月数 | 最新价($) | 标题 |',
  '|---:|---|---|---:|---:|---:|---:|---|');
data.genimoTopProducts.forEach((r, i) => md.push(`| ${i + 1} | ${r.parent || r.listingKey || '-'} | ${r.asin || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${r.months} | ${fmt(r.latestPrice, 2)} | ${String(r.title || '').replace(/\|/g, '\\|')} |`));

// 八、趋势结论与GENIMO建议
const leadershipIndustry = leadershipBenchmark.industry;
const leadershipBsr = leadershipBenchmark.bsrTop100;
md.push('', '## 八、趋势结论与GENIMO建议（基于2026.01-06实绩）', '',
  leadershipIndustry && leadershipIndustry.available
    ? `- **整体市场（领导验收主口径）**：计划部 BI 全类目 2026.01-06 销量 ${fmt(leadershipIndustry.currentSales)}，2025.01-06 ${fmt(leadershipIndustry.baselineSales)}；独立重算方向变化 **${fmtPct(leadershipIndustry.growthPct)}**。该基准只提供销量，不推导销售额/均价。${leadershipBsr && leadershipBsr.available ? ` BSR Top100 同期销量方向变化 ${fmtPct(leadershipBsr.growthPct)}。` : ''}`
    : `- **整体市场（领导验收主口径）**：计划部 BI 参考表读取失败（${leadershipBenchmark.error || '未知错误'}），暂不能给出正向验收值）。`,
  `- **整体市场（market.db 明细参考）**：2026.01-06销量 ${fmt(insight.overall2026.sales)}、销售额 $${fmt(insight.overall2026.revenue)}；相对2025同期方向变化为 ${fmtPct(insight.overall2026.yoySales)} / ${fmtPct(insight.overall2026.yoyRevenue)}，加权成交均价方向变化 ${fmtPct(insight.overall2026.yoyWeightedPrice)}。该负向值来自 2025 行级含变体与 2026 父体去重的混合统计单元，仅作不可比明细参考，不作为领导验收主结论。`,
  `- **PP市场**：2026.01-06销量 ${fmt(insight.pp2026.sales)}、销售额 $${fmt(insight.pp2026.revenue)}；相对2025同期的方向变化为 ${fmtPct(insight.pp2026.yoySales)} / ${fmtPct(insight.pp2026.yoyRevenue)}，销量占整体 ${fmt(insight.ppSalesShare2026, 1)}%，核心峰值为 ${fmtMonth(insight.ppPeak2026.month)} 的 ${fmt(insight.ppPeak2026.sales)} 件。`,
  `- **高客单非PP**：2026.01-06销量 ${fmt(insight.high2026.sales)}、销售额 $${fmt(insight.high2026.revenue)}；相对2025同期的方向变化为 ${fmtPct(insight.high2026.yoySales)} / ${fmtPct(insight.high2026.yoyRevenue)}，加权成交均价方向变化 ${fmtPct(insight.high2026.yoyWeightedPrice)}。`,
  `- **GENIMO**：2026.01-06品牌销量 ${fmt(insight.genimo2026.sales)}、销售额 $${fmt(insight.genimo2026.revenue)}；相对2025同期的方向变化为 ${fmtPct(insight.genimo2026.yoySales)} / ${fmtPct(insight.genimo2026.yoyRevenue)}，GENIMO在PP中的销量/销售额份额为 ${fmt(insight.genimoPpShare2026, 2)}% / ${fmt(insight.genimoPpRevenueShare2026, 2)}%。`, '',
  '### 建议', '',
  `1. 2026.06整体销量 ${fmt(categories.overall.monthly.find((row) => row.month === '202606').sales)}，为核心期峰值；2027旺季补货和广告应在峰值前4-8周完成。`,
  `2. PP在2026.01-06贡献整体销量 ${fmt(insight.ppSalesShare2026, 1)}%、销售额 ${fmt(insight.ppRevenueShare2026, 1)}%；保持PP流量盘，同时用非PP高客单产品修复销售额与价格结构。`,
  `3. GENIMO在2026.06的BSR头部仅 ${fmt(genimoJuneHead.skuCount)} 个Listing，却贡献品牌当月销量 ${fmt(genimoJuneHeadSalesShare, 1)}%；应控制头部集中风险，并补强中部 ${fmt(genimoJuneMiddle.skuCount)} 个、尾部 ${fmt(genimoJuneTail.skuCount)} 个在榜Listing。`,
  `4. GENIMO 2026.06相对2025同月的月度MOM/环比销量、销售额方向变化为 ${fmtPct(insight.genimoJune2026.momSales)} / ${fmtPct(insight.genimoJune2026.momRevenue)}；2027扩量必须同步约束毛利、TACOS和库存覆盖。`,
  '5. 2026.07已并入整体市场趋势表，但只有94个父体，不得把其总量变化解读为完整市场同比或环比。',
  '6. GENIMO 2027规划采用“1个头部锚点 + 3-5个中部利润层 + 4-8个尾部测试池”，并执行下述晋级/退出门槛。');
md.push('', '### GENIMO 2027产品规划（来自参考 workbook，SPEC 1.3/7.7/验收22）', '',
  `- **2026实绩基线**：2026.01-06销量 ${fmt(insight.genimo2026.sales)}、销售额 $${fmt(insight.genimo2026.revenue)}；2026.06头部/中部/尾部在榜Listing分别为 ${fmt(genimoJuneHead.skuCount)} / ${fmt(genimoJuneMiddle.skuCount)} / ${fmt(genimoJuneTail.skuCount)}。`,
  `- **2026主力父体**：${insight.genimoTopProduct2026 ? (insight.genimoTopProduct2026.parent || insight.genimoTopProduct2026.listingKey) + '，累计销量 ' + fmt(insight.genimoTopProduct2026.sales) : '无可用数据'}。`,
  '- **链接组合（Link Portfolio）**：1 个头部锚点 + 3-5 个中部利润层 + 4-8 个尾部测试池。',
  '- **头部锚点**：仅保留 1 个 BSR 1-20 核心锚点，参考款为 5x8 Black Gray；承担流量、类目权重和品牌防守，结算毛利率达标后才扩量。',
  '- **中部利润层**：BSR 21-50，重点扩张 8x10、9x12、10x14 及 Black Beige/Blue Grey 差异化组合，贡献主要销售额与利润。',
  '- **尾部测试池**：BSR 51-100，以低库存和精准长尾投放测试新花型、特殊尺寸与场景款，验证通过才加量。',
  '- **工艺/包装小范围验证**：取消包边、包装袋或地钉等改动，先在低风险颜色或 4x6/5x8 做小批量测试；新旧工艺保留批次标记，持续跟踪散边、卷边、破损和退货原因，累计 300-500 单后再决定是否扩大。8x10、9x12及超大尺寸暂时保留更稳定的边缘处理。',
  '- **决策门槛（晋级/退出）**：',
  '  - 尾部→中部：连续 4 周 BSR ≤ 100、CVR 达到类目基准、TACOS ≤ 15%、库存覆盖 ≤ 90 天。',
  '  - 中部→头部：连续 6 周 BSR ≤ 50、贡献毛利 ≥ 15%、自然单占比提升，且可支撑 60 天补货周期。',
  '  - 头部继续扩量：结算毛利率 ≥ 5%、TACOS ≤ 12%；若亏损连续 14 天，降低广告并将销量占比收缩 2-3pp。',
  '  - 退出：连续 90 天无法进入前100，或库存/广告占用明显高于增量利润。',
  '- **尺寸角色分工**：5x8 Black Gray 为头部锚点；8x10、9x12、10x14 为中部利润层；新花型、特殊尺寸和场景款进入尾部测试池。',
  '- **2027 决策建议**：3-5月旺季前完成头部锚点补货与中部新品上架；Q3 复盘尾部测试池，Q4 确定 2028 组合。',
  );


// 九、2027规划与预测 (SPEC 1.3, 7.7, 验收22)
md.push('', '## 九、2027规划与预测（预测/假设，非历史实绩）', '',
  '> 以下数据来自参考 workbook 的预测基准和程序综合研判，均标注为“预测/假设”，不作为历史实绩使用。',
  '',
  '### 2026年9—12月市场趋势预测', '',
  '| 月份 | 销量基准预测 | 销量可能区间 | 销售额基准预测 | 市场阶段 |',
  '|---|---:|---:|---:|---|');
for (const fm of FORECAST_2026_Q4) {
  md.push(`| ${fmtMonth(fm.month)} | 约${fmt(fm.sales)} | ${fm.range} | 约${fmt(fm.rev,0)}美元 | ${fm.stage} |`);
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
md.push('', '### 预测可调整参数', '',
  '> 预测值均为参考 workbook 基准，不是历史实绩。更新情景时只调整下列参数，不回写历史数据。', '',
  '| 参数 | 默认值 | 调整方式 |', '|---|---|---|');
for (const fp of FORECAST_PARAMETERS) md.push(`| ${fp.parameter} | ${fp.defaultValue} | ${fp.effect} |`);

// 十、参考材料核对 (SPEC 验收24)
md.push('', '## 十、参考材料核对', '',
  '- 两份参考 workbook 由领导提供并确认通过（PP管数据、BSR年度分层与2027规划）；已核对工作表结构、筛选公式和BSR解析规则。',
  '- PP workbook 采用独立Listing键（父ASIN优先）去重，2025.1-2026.7 含父ASIN；市场DB 2025年 ASIN/父ASIN 已从源表富文本超链接恢复，当前 PP 前100可按父体复核，但主源行级/变体口径与参考 workbook 的筛选边界仍需分别披露。',
  '- 2026.01-06 已更新为全市场父体级快照（1038-1993父体/月，替代原64-94父体口径），与领导参考 workbook 的2026年PP数据源不同，月度总量不可直接横向比较；2026.07 仍为94父体快照。',
  '- 计划部核对 workbook（实际路径：新增参考的材料和内容/销量预测计划部底表-户外地垫.xlsx）：行业大盘 BI 全类目 2026H1 同比独立重算为 +2.8106%；BSR Top100 按 Q=1..100、AX:BC 对 BJ:BO 独立重算为 +5.0910%。原表 P 列误用 BJ:BN+BP，且 M:P 未覆盖 Q=79..100，以上公式缺陷仅作核对记录，不替代当前市场 DB。',
  '- 参考 workbook Cohort 进退层：2025 Top100 parents=160、2026 Top100 parents=144、Retained=53、Exited=107、Entered=91。该核实使用参考 workbook 自身数据源（含父ASIN）；市场DB 2025虽已恢复父ASIN，但源数据为行级/变体展开，跨年父体进退仅作方向性参考。',
  '- 参考 workbook GENIMO 2027规划已整合进本报告第八节建议；建议中的历史实绩基线统一使用2026.01-06数据，参考 workbook 只提供规划结构、门槛和预测假设。',
  '');

function esc(s) {
  return String(s).replace(/\s+/g, ' ').trim()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function htmlTable(headers, rows) {
  return '<div class="table-wrap"><table><thead><tr>' + headers.map((h) => '<th>' + esc(h) + '</th>').join('')
    + '</tr></thead><tbody>' + rows.map((row) => '<tr>' + row.map((v) => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('')
    + '</tbody></table></div>';
}

function svgLineChart(title, rows, field, valueLabel) {
  const width = 760;
  const height = 230;
  const left = 28;
  const right = 18;
  const top = 22;
  const bottom = 42;
  const values = rows.map((row) => Number(row[field] || 0));
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const span = Math.max(max - min, 1);
  const x = (index) => left + (width - left - right) * index / Math.max(rows.length - 1, 1);
  const y = (value) => top + (height - top - bottom) * (1 - (value - min) / span);
  const points = values.map((value, index) => `${x(index).toFixed(1)},${y(value).toFixed(1)}`).join(' ');
  const monthLabels = rows.map((row, index) => `<text x="${x(index).toFixed(1)}" y="${height - 13}" text-anchor="middle">${esc(fmtMonth(row.month).slice(5))}</text>`).join('');
  const circles = values.map((value, index) => `<circle cx="${x(index).toFixed(1)}" cy="${y(value).toFixed(1)}" r="4"><title>${esc(fmtMonth(rows[index].month) + ' ' + valueLabel + ' ' + fmt(value))}</title></circle>`).join('');
  return `<figure class="chart-card"><figcaption><b>${esc(title)}</b><span>2026.01-06 · 同一父体统计口径</span></figcaption><svg viewBox="0 0 ${width} ${height}" role="img" aria-label="${esc(title)}"><line x1="${left}" y1="${height - bottom}" x2="${width - right}" y2="${height - bottom}" class="chart-axis"/><polyline points="${points}" class="chart-line"/>${circles}${monthLabels}</svg></figure>`;
}

function monthlyHtml(rows) {
  return htmlTable(['月份', 'MOM/环比基准月份', 'SKU数', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'MOM/环比销量', 'MOM/环比销售额', 'MOM/环比标价', 'MOM/环比成交均价'],
    rows.map((r) => [fmtMonth(r.month), r.momBasis ? fmtMonth(r.momBasis) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.momAvgListPrice), fmtPct(r.momWeightedPrice)]));
}
function overallMarketTrendHtml(rows) {
  return htmlTable(['月份', '口径状态', 'MOM/环比基准月份', 'Listing数', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'MOM/环比销量', 'MOM/环比销售额'],
    rows.map((row) => [fmtMonth(row.month), row.scopeStatus, row.momBasis ? fmtMonth(row.momBasis) : '-', fmt(row.skuCount), fmt(row.sales), fmt(row.revenue), fmt(row.avgListPrice, 2), fmt(row.weightedPrice, 2), mergedTrendPct(row, 'momSales'), mergedTrendPct(row, 'momRevenue')]));
}
function annualHtml(rows) {
  return htmlTable(['年份/数据周期', 'YoY比较周期', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'YOY销量', 'YOY销售额', 'YOY标价', 'YOY成交均价', '同比范围一致性'],
    rows.map((r) => [`${r.year} (${fmtPeriod(r.period)})`, r.comparison ? fmtPeriod(r.comparison) : '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyAvgListPrice), fmtPct(r.yoyWeightedPrice), scopeLabel(r)]));
}

function segmentHtml(rows) {
  const hasGap = rows.some((r) => r.momGapReason);
  const table = htmlTable(['月份', '小类BSR分层', 'MOM/环比基准月份', 'SKU数', '销量', '销售额($)', '加权成交均价', 'MOM/环比销量', 'MOM/环比销售额', 'MOM/环比成交均价'],
    rows.map((r) => [fmtMonth(r.month), r.segment, r.momBasis ? fmtMonth(r.momBasis) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), segPct(r, 'momSales', 'momGapReason'), segPct(r, 'momRevenue', 'momGapReason'), segPct(r, 'momWeightedPrice', 'momGapReason')]));
  return table + (hasGap
    ? '<p class="note">*无对应数据：跨年同月基准分层为空（0条Listing），MOM/环比无法计算。整体市场与高客单的 2026.05 中部/尾部缺失源于 2025.05 源数据小类BSR同值重复（详见一、口径说明）；GENIMO 部分月份分层无在榜商品属正常稀疏。</p>'
    : '');
}
function annualSegmentsHtml(rows) {
  return htmlTable(['年份/数据周期', 'BSR分层', 'YoY比较周期', 'SKU数', '销量', '销售额($)', '加权成交均价($)', 'YoY销量', 'YoY销售额', 'YoY成交均价', '同比范围一致性'],
    rows.map((r) => [`${r.year} (${fmtPeriod(r.period)})`, r.segment, r.comparison ? fmtPeriod(r.comparison) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyWeightedPrice), scopeLabel(r)]));
}

function trendHtml(c, category, label) {
  return trendAnalysis(c, category, label).split('\n').filter(Boolean).map((line) => {
    if (line.startsWith('### ')) return '<h3>' + esc(line.slice(4)) + '</h3>';
    return '<p>' + esc(line.replace(/^- /, '')) + '</p>';
  }).join('');
}

const htmlSections = ['overall', 'pp', 'high', 'genimo'].map((category, index) => {
  const c = categories[category];
  const mergedOverallTrend = category === 'overall'
    ? `<details open class="trend-details"><summary><b>2026.01-07整体市场趋势（合并展示）</b></summary>${overallMarketTrendHtml(overallMarketTrend2026)}<p class="note">2026.07只有94个父体，是小范围样本；已按领导反馈并入同一趋势表，但不与1-6月直接计算同比、环比或累计。</p></details>`
    : '';
  const replacementAudit = category === 'overall'
    ? `<details><summary><b>2026数据替换审计记录</b>（逐月来源、行数、SHA-256）</summary>${htmlTable(['月份','主表替换前行数','快照原始行数','当前父体行数','源库SHA-256','应用时间'], replacementMetadata.map((row) => [fmtMonth(row.month), fmt(row.base_imported_rows), fmt(row.source_raw_rows), fmt(row.replacement_rows), row.source_sha256 || '-', row.applied_at || '-']))}<p class="note">记录来自 market.db.analysis_replacements；用于证明替换来源和当前行数可复跑。</p></details>`
    : '';
  const ppListingDetail = category === 'pp'
    ? (() => {
      const rows = data.ppListingDetails.filter((row) => row.month === REPORT_CUTOFF);
      return `<details><summary><b>PP独立Listing明细（${fmtMonth(REPORT_CUTOFF)}，${fmt(rows.length)}条）</b></summary><p class="note">完整JSON保留 ${fmtPeriod(data.ppListingDetailsPeriod)} 共 ${fmt(data.ppListingDetails.length)} 条月度Listing记录；下表为核心截止月独立Listing。</p>${htmlTable(['月份','Listing键','父ASIN','代表ASIN','小类BSR','多值解析','销量','销售额($)','价格($)','商品标题'], rows.map((row) => [fmtMonth(row.month), row.listingKey || '-', row.parent || '-', row.asin || '-', row.rank === null ? '-' : fmt(row.rank), row.bsrMulti ? '是' : '否', fmt(row.sales), fmt(row.revenue), fmt(row.price, 2), row.title || '-']))}</details>`;
    })()
    : '';
  return `<section id="${category}"><h2>${index + 2}、${esc(labels[category])}</h2>${mergedOverallTrend}${replacementAudit}<details open class="trend-details"><summary><b>${esc(labels[category])}趋势分析</b></summary><div class="trend-body">${trendHtml(c, category, labels[category])}</div></details><details open><summary>月度指标、用户定义MOM与月度环比</summary>${monthlyHtml(c.monthly)}</details><details><summary>年度/同周期汇总</summary>${annualHtml(c.annual)}</details><details open class="bsr-details"><summary><b>小类BSR Top100（前100）分析</b>（1-100名汇总 + 年度 + 头中尾 + 五档）</summary><h4>BSR Top100月度汇总（1-100名）</h4>${monthlyHtml(c.bsrTop100.monthly)}<h4>BSR Top100年度/同周期汇总</h4>${annualHtml(c.bsrTop100.annual)}<h4>BSR头部/中部/尾部（月度）</h4>${segmentHtml(c.bsrGroups.monthly)}<h4>BSR头部/中部/尾部（年度）</h4>${annualSegmentsHtml(c.bsrGroups.annual)}<h4>BSR五档分层（月度）</h4>${segmentHtml(c.bsrSegments.monthly)}<h4>BSR五档分层（年度）</h4>${annualSegmentsHtml(c.bsrSegments.annual)}</details>${ppListingDetail}</section>`;
}).join('\n');

const insightHtml = `<section id="insights"><h2>八、趋势结论与GENIMO建议（基于2026.01-06实绩）</h2><p class="note">2025为行级、2026为父体级，以下跨年百分比均为方向变化，不是严格同口径同比；2026月度环比保持父体口径一致。</p><div class="insight-grid"><article><h3>整体市场</h3><p>2026.01-06销量 ${fmt(insight.overall2026.sales)}、销售额 $${fmt(insight.overall2026.revenue)}；相对2025同期方向变化 ${fmtPct(insight.overall2026.yoySales)} / ${fmtPct(insight.overall2026.yoyRevenue)}，加权成交均价方向变化 ${fmtPct(insight.overall2026.yoyWeightedPrice)}。</p></article><article><h3>PP塑料地垫</h3><p>2026.01-06销量 ${fmt(insight.pp2026.sales)}、销售额 $${fmt(insight.pp2026.revenue)}；相对2025同期方向变化 ${fmtPct(insight.pp2026.yoySales)} / ${fmtPct(insight.pp2026.yoyRevenue)}，销量占整体 ${fmt(insight.ppSalesShare2026, 1)}%，核心峰值为 ${fmtMonth(insight.ppPeak2026.month)} 的 ${fmt(insight.ppPeak2026.sales)} 件。</p></article><article><h3>高客单非PP</h3><p>2026.01-06销量 ${fmt(insight.high2026.sales)}、销售额 $${fmt(insight.high2026.revenue)}；相对2025同期方向变化 ${fmtPct(insight.high2026.yoySales)} / ${fmtPct(insight.high2026.yoyRevenue)}，加权成交均价方向变化 ${fmtPct(insight.high2026.yoyWeightedPrice)}。</p></article><article><h3>GENIMO</h3><p>2026.01-06品牌销量 ${fmt(insight.genimo2026.sales)}、销售额 $${fmt(insight.genimo2026.revenue)}；相对2025同期方向变化 ${fmtPct(insight.genimo2026.yoySales)} / ${fmtPct(insight.genimo2026.yoyRevenue)}，GENIMO在PP中的销量/销售额份额为 ${fmt(insight.genimoPpShare2026, 2)}% / ${fmt(insight.genimoPpRevenueShare2026, 2)}%。</p></article></div><h3>GENIMO 2027产品规划（领导参考 workbook）</h3><ul><li>2026实绩基线：2026.01-06销量 ${fmt(insight.genimo2026.sales)}、销售额 $${fmt(insight.genimo2026.revenue)}；2026.06头部/中部/尾部在榜Listing分别为 ${fmt(genimoJuneHead.skuCount)} / ${fmt(genimoJuneMiddle.skuCount)} / ${fmt(genimoJuneTail.skuCount)}。</li><li>链接组合：1个头部锚点 + 3-5个中部利润层 + 4-8个尾部测试池。</li><li>头部锚点：1个BSR 1-20核心款；结算毛利率≥5%、TACOS≤12%才继续扩量。</li><li>中部利润层：BSR 21-50，重点扩张8x10、9x12、10x14及Black Beige/Blue Grey差异化组合。</li><li>尾部测试池：BSR 51-100，以低库存测试新花型、特殊尺寸和场景款。</li><li>工艺/包装小范围验证：取消包边、包装袋或地钉等改动先在低风险颜色或4x6/5x8小批量测试；新旧工艺保留批次标记，跟踪散边、卷边、破损和退货原因，累计300-500单后再决定扩大；8x10、9x12及超大尺寸暂时保留更稳定的边缘处理。</li><li>尾部→中部：连续4周BSR≤100、CVR达到类目基准、TACOS≤15%、库存覆盖≤90天。</li><li>中部→头部：连续6周BSR≤50、贡献毛利≥15%、自然单占比提升且可支撑60天补货周期。</li><li>头部继续扩量：结算毛利率≥5%、TACOS≤12%；若亏损连续14天，降低广告并收缩销量占比2-3pp。</li><li>退出：连续90天无法进入前100，或库存/广告占用明显高于增量利润。</li></ul><h3>行动建议</h3><ol><li>2026.06整体销量 ${fmt(categories.overall.monthly.find((row) => row.month === '202606').sales)}、环比 ${fmtPct(categories.overall.monthly.find((row) => row.month === '202606').chainSales)}，为核心期峰值；2027旺季补货和广告应在峰值前4-8周完成。</li><li>PP在2026.01-06贡献整体销量 ${fmt(insight.ppSalesShare2026, 1)}%、销售额 ${fmt(insight.ppRevenueShare2026, 1)}%；保持PP流量盘，同时用非PP高客单产品修复销售额与价格结构。</li><li>GENIMO在2026.06的BSR头部 ${fmt(genimoJuneHead.skuCount)} 个Listing贡献品牌当月销量 ${fmt(genimoJuneHeadSalesShare, 1)}%；应控制头部集中风险，并补强中部与尾部在榜Listing。</li><li>GENIMO 2026.06相对2025同月的销量/销售额方向变化 ${fmtPct(insight.genimoJune2026.momSales)} / ${fmtPct(insight.genimoJune2026.momRevenue)}，月度环比 ${fmtPct(insight.genimoJune2026.chainSales)} / ${fmtPct(insight.genimoJune2026.chainRevenue)}；扩量必须同步约束毛利、TACOS和库存覆盖。</li><li>2026.07已并入整体市场趋势表，但只有94个父体，不得把其总量变化解读为完整市场同比或环比。</li></ol></section>`;
const renderedInsightHtml = leadershipIndustry && leadershipIndustry.available
  ? insightHtml.replace(/<article><h3>整体市场<\/h3><p>[\s\S]*?<\/p><\/article>/,
    `<article><h3>整体市场（领导验收主口径）</h3><p>计划部 BI 全类目 2026.01-06 销量 ${fmt(leadershipIndustry.currentSales)} vs 2025.01-06 ${fmt(leadershipIndustry.baselineSales)}，方向变化 <b>${fmtPct(leadershipIndustry.growthPct)}</b>；该基准只提供销量，不推导销售额/均价。当前 market.db 明细快照销量 ${fmt(insight.overall2026.sales)}、销售额 $${fmt(insight.overall2026.revenue)}，销量/销售额方向 ${fmtPct(insight.overall2026.yoySales)} / ${fmtPct(insight.overall2026.yoyRevenue)}，加权成交均价方向变化 ${fmtPct(insight.overall2026.yoyWeightedPrice)}，因统计单元不同仅作不可比参考。</p></article>`)
  : insightHtml;

const genimoProductsHtml = `<section id="genimo-products"><h2>七、GENIMO 2026.01-06累计Top父体</h2><p class="note">仅累计2026.01-06核心实绩；按父ASIN优先的独立Listing键聚合，用于识别2026主力父体。</p>${htmlTable(['排名', '父体/Listing', '代表ASIN', '累计销量', '累计销售额($)', '覆盖月数', '最新价($)', '商品标题'], data.genimoTopProducts.map((row, index) => [index + 1, row.parent || row.listingKey || '-', row.asin || '-', fmt(row.sales), fmt(row.revenue), row.months, fmt(row.latestPrice, 2), row.title || '-']))}</section>`;
const cohortHtml = '<section id="cohort"><h2>六、父体进退（Cohort）</h2><p>BSR前100按父ASIN（优先）或ASIN统计的留存、退出、新进入及头/中/尾迁移（SPEC 7.6/验收23）。比较周期：2026.01 vs 2026.06（核心分析首尾月；主源历史 ASIN 为富文本恢复值，2026 使用父体替换快照）。</p>' + ['overall','pp','high','genimo'].map((key) => { const co = categories[key].cohort; if (!co) return ''; return '<p><b>' + labels[key] + '</b>：前100父体池从 ' + co.fromParents + ' 变为 ' + co.toParents + '；留存 ' + co.retained + '、退出 ' + co.exited + '、新进入 ' + co.entered + '。层间迁移：' + Object.entries(co.migration).map(([k, v]) => k + '=' + v).join('、') + '。</p>'; }).join('') + '<p class="note">2025年主源 ASIN/父ASIN 已恢复，但其行级/变体口径与2026父体替换快照不同；跨年父体进退不作为严格同比结论。</p></section>';
const forecastHtml = '<section id="forecast"><h2>九、2027规划与预测（预测/假设，非历史实绩）</h2><p class="note">以下数据来自领导提供并确认通过的参考 workbook 预测基准和程序综合研判，均标注为"预测/假设"，不作为历史实绩使用。</p><h3>2026年9—12月市场趋势预测</h3>' + htmlTable(['月份','销量基准预测','销量可能区间','销售额基准预测','市场阶段'], FORECAST_2026_Q4.map((fm) => [fmtMonth(fm.month), '约' + fmt(fm.sales), fm.range, '约' + fmt(fm.rev,0) + '美元', fm.stage])) + '<h3>2027年销量和销售额趋势预测</h3>' + htmlTable(['月份','2027年销量基准预测','销售额基准预测','趋势'], FORECAST_2027_MONTHLY.map((fm) => [fmtMonth(fm.month), fmt(fm.sales), '约' + fmt(fm.rev,0) + '美元', fm.note])) + '<h3>2027情景</h3>' + htmlTable(['情景','年销量','年销售额','触发条件'], FORECAST_2027_SCENARIOS.map((fs) => [fs.scenario, fs.sales, fs.rev, fs.trigger])) + '<h3>可调整参数</h3><p class="note">更新情景时只调整下列参数，不回写历史实绩。</p>' + htmlTable(['参数','默认值','调整方式'], FORECAST_PARAMETERS.map((fp) => [fp.parameter, fp.defaultValue, fp.effect])) + '</section>';
const referenceHtml = '<section id="reference"><h2>十、参考材料核对</h2><ul><li>两份参考 workbook 由领导提供并确认通过（PP管数据、BSR年度分层与2027规划）；已核对工作表结构、筛选公式和BSR解析规则。</li><li>PP workbook 采用独立Listing键（父ASIN优先）去重，2025.1-2026.7 含父ASIN；市场DB 2025年 ASIN/父ASIN 已从源表富文本超链接恢复，当前 PP 前100可按父体复核，但主源行级/变体口径与参考 workbook 的筛选边界仍需分别披露。</li><li>2026.01-06为全市场父体级快照（1038-1993父体/月），2026.07为94父体小样本；领导参考 workbook 的2026年PP数据与当前市场DB数据源不同，月度总量不可直接横向比较。</li><li>参考 workbook Cohort 进退层：2025 Top100 parents=160、2026 Top100 parents=144、Retained=53、Exited=107、Entered=91。该核实使用参考 workbook 自身数据源（含父ASIN）；市场DB 2025虽已恢复父ASIN，但源数据为行级/变体展开，跨年父体进退仅作方向性参考。</li><li>参考 workbook GENIMO 2027规划结构已整合进第八节；所有历史实绩与量化建议统一采用2026.01-06核心数据。</li></ul></section>';
const planReferenceHtml = '<li>计划部核对 workbook（实际路径：新增参考的材料和内容/销量预测计划部底表-户外地垫.xlsx）：行业大盘 BI 全类目 2026H1 同比独立重算为 +2.8106%；BSR Top100 按 Q=1..100、AX:BC 对 BJ:BO 独立重算为 +5.0910%。原表 P 列误用 BJ:BN+BP，且 M:P 未覆盖 Q=79..100；以上公式缺陷仅作核对记录，不替代当前市场 DB。</li>';

const coverageHtml = `<section id="coverage"><h2>十一、需求覆盖核对</h2>${htmlTable(['要求项', '交付位置', '状态'], [['整体市场销量/销售额/均价', '整体市场月度与年度表', '已覆盖'], ['2026.01-07整体市场趋势合并', '整体市场首张合并趋势表；2026.07跨口径指标标记不适用', '已覆盖'], ['整体月度MOM与环比', '所有月度表分别显示MOM/环比基准月份', '已覆盖'], ['小类前100同比/环比', '各分类 BSR Top100月度、年度表', '已覆盖'], ['BSR多值解析标记', '口径说明中的多值解析审计表及JSON', '已覆盖'], ['头部/中部/尾部及五档', '各分类 BSR分层月度、年度表', '已覆盖'], ['PP独立Listing明细', 'PP章节核心截止月明细；JSON保留2026.01-06全量', '已覆盖'], ['高客单非PP（排除PP后全部产品）', '高客单分类（排除PP后全部产品）', '已覆盖'], ['数据替换逐月可追溯', '整体市场的数据替换审计记录', '已覆盖'], ['趋势与GENIMO建议使用2026实绩', '第八节、GENIMO 2026累计Top父体与数据JSON', '已覆盖'], ['GENIMO工艺/包装验证', '第八节300-500单小批量验证规则', '已覆盖'], ['预测可调整参数', '第九节需求/价格/时间/情景参数', '已覆盖']])}</section>`;

const html = `<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场洞察 · Outdoor Rug Intelligence</title><style>
:root{color-scheme:dark;font-family:Inter,"Segoe UI Variable","Segoe UI","Microsoft YaHei",system-ui,sans-serif;--bg:#050505;--surface:#0d0d0d;--surface-soft:rgba(255,255,255,.018);--surface-strong:rgba(255,255,255,.05);--text:#f2f2f0;--muted:#8b8b88;--faint:#5f5f5c;--line:rgba(255,255,255,.095);--line-strong:rgba(255,255,255,.22);--brand:#f4f4f1;--brand-soft:rgba(255,255,255,.05);--warning:#aaa9a2;--warning-soft:rgba(255,255,255,.04);--sidebar:#060606;--cta:#eeeeeb;--cta-text:#080808;--grid-line:rgba(255,255,255,.012);--card-bg:linear-gradient(145deg,rgba(18,18,18,.72),rgba(8,8,8,.8));--shadow:0 28px 80px rgba(0,0,0,.3),inset 0 1px rgba(255,255,255,.025);--display:"Times New Roman","Songti SC","STSong",serif;--mono:"Cascadia Code","SFMono-Regular",Consolas,monospace}
:root[data-theme="light"]{color-scheme:light;--bg:#f4f3ef;--surface:#fafaf7;--surface-soft:rgba(24,24,20,.03);--surface-strong:rgba(24,24,20,.06);--text:#191916;--muted:#6f6e68;--faint:#9a9992;--line:rgba(24,24,20,.12);--line-strong:rgba(24,24,20,.28);--brand:#1b1b18;--brand-soft:rgba(24,24,20,.05);--warning:#65635b;--warning-soft:rgba(24,24,20,.04);--sidebar:#efede7;--cta:#191916;--cta-text:#f8f7f3;--grid-line:rgba(24,24,20,.025);--card-bg:linear-gradient(145deg,rgba(255,255,252,.82),rgba(245,244,239,.9));--shadow:0 18px 48px rgba(24,24,20,.08),inset 0 1px rgba(255,255,255,.6)}
*{box-sizing:border-box}html{min-width:320px;scroll-behavior:smooth;background:var(--bg)}body{position:relative;margin:0;min-height:100vh;color:var(--text);font-size:18px;background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(255,255,255,.105),transparent 68%),radial-gradient(ellipse 48% 38% at 12% 92%,rgba(255,255,255,.035),transparent 72%),linear-gradient(145deg,#020202 0%,#080808 48%,#030303 100%);line-height:1.65}:root[data-theme="light"] body{background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(20,20,18,.08),transparent 68%),linear-gradient(145deg,#faf9f5,#f0efea 52%,#f7f6f2)}body::before{content:"";position:fixed;inset:0;pointer-events:none;opacity:.22;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:72px 72px;mask-image:linear-gradient(to bottom,#000,transparent 82%)}a{color:inherit}.app-shell{position:relative;z-index:1;min-height:100vh}.sidebar{position:fixed;inset:0 auto 0 0;z-index:30;display:flex;width:232px;flex-direction:column;overflow-y:auto;padding:34px 18px 24px;border-right:1px solid var(--line);background:linear-gradient(180deg,rgba(4,4,4,.97),rgba(8,8,8,.9));box-shadow:18px 0 80px rgba(0,0,0,.45)}:root[data-theme="light"] .sidebar{background:linear-gradient(180deg,#f2f0ea,#eae8e1);box-shadow:12px 0 40px rgba(24,24,20,.06)}.brand{display:flex;align-items:center;gap:12px;padding:0 8px;margin-bottom:34px}.brand-mark{display:grid;width:34px;height:34px;place-items:center;border:1px solid var(--line-strong);border-radius:50%;background:radial-gradient(circle at 35% 28%,#f0f0ed,#696966 38%,#080808 70%);box-shadow:0 0 34px rgba(255,255,255,.12)}.brand strong{display:block;font-family:var(--display);font-size:18px;font-weight:400;letter-spacing:.07em}.brand small{display:block;margin-top:3px;color:var(--muted);font-size:13px}.local-badge{display:flex;align-items:center;gap:7px;width:max-content;margin:10px 8px 24px;padding:6px 9px;border:1px solid var(--line);font-family:var(--mono);font-size:13px;letter-spacing:.14em;color:var(--muted)}.local-badge i,.privacy-chip i{width:6px;height:6px;border-radius:50%;background:var(--brand);box-shadow:0 0 12px rgba(255,255,255,.55)}.nav-label{display:block;padding:0 10px 7px;color:var(--muted);font-family:var(--mono);font-size:14px;letter-spacing:.22em}.sidebar nav{display:grid;gap:2px}.sidebar nav a{position:relative;display:flex;align-items:center;gap:11px;padding:11px 10px;color:var(--muted);font-size:15px;letter-spacing:.04em;text-decoration:none}.sidebar nav a:hover,.sidebar nav a.is-active{color:var(--text);background:var(--surface-strong)}.nav-icon{display:grid;width:23px;height:23px;place-items:center;border-radius:50%;font-family:var(--mono);font-size:14px}.sidebar-footer{margin-top:auto;padding:16px 10px 3px;border-top:1px solid var(--line)}.sidebar-footer span,.sidebar-footer strong,.sidebar-footer small{display:block}.sidebar-footer span,.sidebar-footer small{color:var(--muted);font-size:13px}.sidebar-footer strong{margin:7px 0 3px;font-size:15px}.workspace{min-height:100vh;margin-left:232px}.topbar{display:flex;min-height:154px;align-items:center;gap:18px;padding:26px clamp(28px,4vw,64px);border-bottom:1px solid var(--line);background:color-mix(in srgb,var(--bg),var(--surface) 25%)}.eyebrow{display:block;margin-bottom:7px;color:var(--faint);font-family:var(--mono);font-size:13px;letter-spacing:.28em}.topbar h1{margin:0 0 8px;font-family:var(--display);font-size:clamp(44px,5vw,70px);font-weight:400;letter-spacing:-.04em;background:linear-gradient(110deg,#fff 8%,#adada8 55%,#535350 100%);background-clip:text;color:transparent}:root[data-theme="light"] .topbar h1{background:linear-gradient(110deg,#191916,#55544e 55%,#9a9992);background-clip:text;color:transparent}.topbar p{margin:0;color:var(--muted);font-size:16px}.top-actions{display:flex;align-items:center;gap:10px;margin-left:auto}.privacy-chip,.theme-button{border:1px solid var(--line);color:var(--muted);background:var(--surface)}.privacy-chip{display:flex;align-items:center;gap:8px;padding:8px 11px;font-size:14px}.theme-button{display:grid;width:38px;height:38px;place-items:center;cursor:pointer}.content{width:min(1500px,100%);margin:0 auto;padding:32px clamp(28px,4vw,64px) 72px}.scope-notice{display:flex;gap:12px;padding:13px 16px;border:1px solid var(--line-strong);color:var(--muted);background:var(--brand-soft);font-size:16px}.scope-notice b{color:var(--text)}.metrics-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:22px 0}.metric-card{position:relative;min-height:142px;padding:20px;border:1px solid var(--line);background:var(--card-bg);box-shadow:var(--shadow)}.metric-card::before,section::before{content:"";position:absolute;inset:0 0 auto;height:1px;background:linear-gradient(90deg,rgba(255,255,255,.7),transparent)}.metric-label,.metric-value,.metric-note{display:block}.metric-label{color:var(--muted);font-size:15px}.metric-value{margin:15px 0 9px;font-family:var(--display);font-size:clamp(30px,2.6vw,40px);font-weight:400}.metric-note{color:var(--faint);font-size:14px;line-height:1.5}section{position:relative;min-width:0;margin-bottom:22px;padding:23px;border:1px solid var(--line);background:var(--surface);box-shadow:var(--shadow);scroll-margin-top:18px}section h2{margin:0 0 6px;font-family:var(--display);font-size:28px;font-weight:400;letter-spacing:-.01em}section h3{font-size:20px;letter-spacing:.05em}section li{color:var(--muted);font-size:16px;line-height:1.8}section li b{color:var(--text)}details{padding:15px 0;border-top:1px solid var(--line)}summary{cursor:pointer;color:var(--text);font-size:16px;font-weight:700;letter-spacing:.04em}.table-wrap{max-height:620px;margin-top:13px;overflow:auto;border:1px solid var(--line)}table{width:100%;min-width:980px;border-collapse:collapse;font-family:var(--mono);font-size:15px}th,td{padding:9px 11px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th{position:sticky;top:0;z-index:1;color:var(--muted);background:var(--surface);font-size:15px;letter-spacing:.06em}th:first-child,td:first-child{text-align:left}tbody tr:hover{background:var(--surface-soft)}.note{margin-top:18px;padding:13px 16px;border-left:2px solid var(--warning);color:var(--muted);background:var(--warning-soft);font-size:15px}code{padding:2px 5px;color:var(--text);background:var(--surface-strong);font-family:var(--mono)}
@media(max-width:980px){.sidebar{transform:translateX(-100%)}.workspace{margin-left:0}.metrics-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:560px){.topbar{padding:24px 18px}.privacy-chip{display:none}.content{padding:20px 12px 50px}.metrics-grid{grid-template-columns:1fr}.metric-card{min-height:116px}section{padding:17px}.topbar h1{font-size:44px}}
</style></head><body><div class="app-shell"><aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫市场分析系统</small></div></div><div class="local-badge"><i></i> DATA · VERIFIED</div><span class="nav-label">市场分析</span><nav><a href="#dashboard"><span class="nav-icon">总</span>市场总览</a><a href="#definitions"><span class="nav-icon">径</span>数据口径</a><a href="#overall"><span class="nav-icon">整</span>整体市场</a><a href="#pp"><span class="nav-icon">PP</span>塑料地垫</a><a href="#high"><span class="nav-icon">高</span>高客单非PP</a><a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#anomaly"><span class="nav-icon">!</span>参考附录</a><a href="#insights"><span class="nav-icon">策</span>趋势与建议</a><a href="#coverage"><span class="nav-icon">核</span>需求覆盖</a></nav><div class="sidebar-footer"><span>可比数据范围</span><strong>${analysisMonths[0]} — ${analysisMonths[analysisMonths.length - 1]}</strong><small>${analysisMonths.length}个月 · 444万单元格核验</small></div></aside><div class="workspace"><header class="topbar"><div><span class="eyebrow">OUTDOOR RUG MARKET INTELLIGENCE</span><h1>市场总览</h1><p>销量、销售额、双均价 · 小类BSR前100 · 月度MoM与年度YoY</p></div><div class="top-actions"><span class="privacy-chip"><i></i> 源数据只读</span><button class="theme-button" id="theme-toggle" aria-label="切换主题">☀</button></div></header><main class="content"><div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>核心 49 个月（202206-202606）含竞品父ASIN去重口径；2026.07 仅作附录/参考。2026 全年已统一为竞品父ASIN去重口径，与 2025 全市场口径不同。</div></div><div class="metrics-grid"><article class="metric-card"><span class="metric-label">2025 整体销量</span><strong class="metric-value">${fmt(insight.overall2025.sales)}</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoySales)} · 全市场</span></article><article class="metric-card"><span class="metric-label">2025 整体销售额</span><strong class="metric-value">$${fmt(insight.overall2025.revenue / 1000000, 1)}M</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoyRevenue)} · USD</span></article><article class="metric-card"><span class="metric-label">PP 销量贡献</span><strong class="metric-value">${fmt(ppSalesShare2025,1)}%</strong><span class="metric-note">2025 PP销量 / 整体销量</span></article><article class="metric-card"><span class="metric-label">GENIMO PP份额</span><strong class="metric-value">${fmt(insight.genimoPpShare2025,2)}%</strong><span class="metric-note">2025销量份额 · 品牌领先</span></article></div><section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>，不是按月销量重新排名。</li><li>PP标题包含 plastic；高客单排除PP后按材质关键词或价格≥$40。</li><li>同时展示SKU平均标价与销量加权均价（销售额/销量）。</li><li>年度YoY严格使用同周期；2023仅以6-12月对比2022年6-12月。</li><li>月度MOM = 今年X月 vs 去年X月（跨年同月，如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）。</li></ul><p class="note">2026全年已统一为竞品快照（按父ASIN去重，随机保留一条）口径，每月64-94个父商品；与2025全市场口径（1700-2000 SKU）跨年对比存在范围差异，请以可比口径列为准。</p></section>${htmlSections}<section id="anomaly"><h2>六、2026.03-07竞品替换数据附录（父ASIN去重）</h2>${htmlTable(['月份','纳入可比报告','销量','销售额($)','SKU平均标价','销量加权均价'],sourceDiagnostics.filter((r)=>r.month>='202603').map((r)=>[r.month,r.includedInComparableReport?'是':'否',fmt(r.sales),fmt(r.revenue),fmt(r.avgListPrice,2),fmt(r.weightedPrice,2)]))}</section>${insightHtml}</main></div></div><script>(function(){var root=document.documentElement,button=document.getElementById('theme-toggle');var saved='dark';try{saved=localStorage.getItem('market-report-theme')||'dark'}catch(e){}root.dataset.theme=saved;button.textContent=saved==='dark'?'☀':'☾';button.addEventListener('click',function(){var next=root.dataset.theme==='dark'?'light':'dark';root.dataset.theme=next;button.textContent=next==='dark'?'☀':'☾';try{localStorage.setItem('market-report-theme',next)}catch(e){}});var links=[].slice.call(document.querySelectorAll('.sidebar nav a'));var observer=new IntersectionObserver(function(entries){entries.forEach(function(entry){if(entry.isIntersecting){links.forEach(function(a){a.classList.toggle('is-active',a.getAttribute('href')==='#'+entry.target.id)})}})},{rootMargin:'-20% 0px -70% 0px'});document.querySelectorAll('[id]').forEach(function(el){observer.observe(el)})})();</script></body></html>`;

const coverageScopeText = `核心明细覆盖 ${analysisMonths.length} 个月（${fmtMonth(analysisMonths[0])}-${fmtMonth(analysisMonths[analysisMonths.length - 1])}）；当前库 ${sourceMeta.effective_sheets || 54} 张有效业务表、${fmt(currentDataRowCount)} 条实际记录、${fmt(verifiedDataCellCount)} 个实际数据单元格（原始Excel导入元数据为 ${fmt(sourceMeta.total_rows || 0)} 行，2026.01-07 后续替换为竞品父体口径）。整体市场趋势合并展示2026.01-07；2026.07只有94父体，不参与同比、环比、累计和核心结论。`;
const dashboardMetricsHtml = `<div class="metrics-grid"><article class="metric-card"><span class="metric-label">领导 BI H1 销量</span><strong class="metric-value">${leadershipIndustry && leadershipIndustry.available ? fmt(leadershipIndustry.currentSales) : '-'}</strong><span class="metric-note">相对2025方向 ${leadershipIndustry && leadershipIndustry.available ? fmtPct(leadershipIndustry.growthPct) : '-'} · 主验收</span></article><article class="metric-card"><span class="metric-label">领导 BSR Top100 H1</span><strong class="metric-value">${leadershipBsr && leadershipBsr.available ? fmt(leadershipBsr.currentSales) : '-'}</strong><span class="metric-note">相对2025方向 ${leadershipBsr && leadershipBsr.available ? fmtPct(leadershipBsr.growthPct) : '-'} · 辅助验收</span></article><article class="metric-card"><span class="metric-label">market.db 明细销量</span><strong class="metric-value">${fmt(insight.overall2026.sales)}</strong><span class="metric-note">不可比快照方向 ${fmtPct(insight.overall2026.yoySales)}</span></article><article class="metric-card"><span class="metric-label">market.db 明细销售额</span><strong class="metric-value">$${fmt(insight.overall2026.revenue / 1000000, 1)}M</strong><span class="metric-note">不可比快照方向 ${fmtPct(insight.overall2026.yoyRevenue)}</span></article></div>`;
const leadershipBenchmarkHtml = leadershipIndustry && leadershipIndustry.available
  ? `<section id="leadership-benchmark"><h2>领导验收主基准</h2><p><b>计划部 BI 全类目 Outdoor Rugs：</b>2026.01-06 销量 ${fmt(leadershipIndustry.currentSales)} vs 2025.01-06 ${fmt(leadershipIndustry.baselineSales)}，按 <code>${leadershipIndustry.formula}</code> 独立重算为 <strong>${fmtPct(leadershipIndustry.growthPct)}</strong>。</p>${leadershipBsr && leadershipBsr.available ? `<p><b>计划部 BSR Top100：</b>2026.01-06 销量 ${fmt(leadershipBsr.currentSales)} vs 2025.01-06 ${fmt(leadershipBsr.baselineSales)}，按 ${leadershipBsr.rankRange} 原始月度列独立重算为 <strong>${fmtPct(leadershipBsr.growthPct)}</strong>。</p>` : ''}<p class="note">该参考 workbook 只有销量字段，不能推导销售额或均价；market.db 的负向全量快照保留在下方并标注为不可比明细，不作为领导验收主结论。</p></section>`
  : `<section id="leadership-benchmark"><h2>领导验收主基准</h2><p class="note">计划部参考 workbook 读取失败：${esc(leadershipBenchmark.error || '未知错误')}。</p></section>`;
const chartMonths = categories.overall.monthly.filter((row) => row.month >= '202601' && row.month <= REPORT_CUTOFF);
const trendChartsHtml = `<section id="visuals"><h2>2026核心趋势可视化</h2><p class="note">三张图均只使用2026.01-06父体级快照，适合观察同口径连续月趋势；跨年变化请以方向性参考口径解读。</p><div class="chart-grid">${svgLineChart('整体市场月销量', chartMonths, 'sales', '销量')}${svgLineChart('整体市场月销售额', chartMonths, 'revenue', '销售额')}${svgLineChart('GENIMO月销量', categories.genimo.monthly.filter((row) => row.month >= '202601' && row.month <= REPORT_CUTOFF), 'sales', '销量')}</div></section>`;
const bsrQualityRows = data.dataQuality.historicalBsrTop100Quality;
const bsrQualityHtml = `<details><summary><b>2022-2025 BSR Top100逐月质量诊断</b></summary><p class="note">历史源表的 ASIN/父ASIN 已从富文本超链接恢复；以下“重复Listing行”和“重复名次行”分别揭示同一父体变体重复及同名次并列，不把行数直接当作独立Listing。2025.05异常100行头部口径基准为${overall202505HeadEvidence}。</p>${htmlTable(['月份','BSR 1-100候选行','截取行数','标识覆盖率','独立Listing数','重复Listing行','不同名次数','重复名次行','统计单元'], bsrQualityRows.map((row) => [fmtMonth(row.month), fmt(row.eligibleRows), fmt(row.selectedRows), fmt(row.identifierCoveragePct, 1) + '%', fmt(row.distinctListingKeys), fmt(row.duplicateListingRows), fmt(row.distinctRanks), fmt(row.repeatedRankRows), row.statisticalUnit]))}</details>`;
const bsrMultiAuditHtml = `<details><summary><b>BSR多值解析审计</b>（${fmt(data.dataQuality.bsrMultiValueAudit.length)}条Top100入选记录）</summary><p class="note">源小类BSR含多个数值时采用可解析最小名次；保留源字符串和多值标记，便于逐条回溯。</p>${htmlTable(['类别','月份','Listing键','父ASIN','ASIN','采用名次','源小类BSR','多值标记'], data.dataQuality.bsrMultiValueAudit.map((row) => [row.category, fmtMonth(row.month), row.listingKey || '-', row.parent || '-', row.asin || '-', fmt(row.rank), row.sourceBsr || '-', '是']))}</details>`;
const quickMd = [
  '# 户外地垫市场分析报告（极速版）', '',
  '> 用于管理层快速阅读；完整数据表、口径、质量诊断与预测区间见优化版HTML/Markdown。', '',
  '## 结论摘要', '',
  leadershipIndustry && leadershipIndustry.available
    ? `- 领导验收主口径（计划部 BI 全类目）：2026.01-06销量 ${fmt(leadershipIndustry.currentSales)} vs 2025.01-06 ${fmt(leadershipIndustry.baselineSales)}，方向变化 **${fmtPct(leadershipIndustry.growthPct)}**；该表仅提供销量。`
    : `- 领导验收主口径读取失败：${leadershipBenchmark.error || '未知错误'}。`,
  `- market.db 父体级快照明细：2026.01-06销量 ${fmt(insight.overall2026.sales)}、销售额 $${fmt(insight.overall2026.revenue)}；相对2025行级导出的方向变化为 ${fmtPct(insight.overall2026.yoySales)} / ${fmtPct(insight.overall2026.yoyRevenue)}，因统计单元不同仅作不可比明细参考。`,
  `- PP贡献2026.01-06整体销量 ${fmt(insight.ppSalesShare2026, 1)}%、销售额 ${fmt(insight.ppRevenueShare2026, 1)}%；GENIMO占PP销量/销售额 ${fmt(insight.genimoPpShare2026, 2)}% / ${fmt(insight.genimoPpRevenueShare2026, 2)}%。`,
  `- 2026.06整体销量 ${fmt(categories.overall.monthly.find((row) => row.month === '202606').sales)}，为核心期峰值；该峰值用于安排2027旺季前4-8周的补货、广告与新品测试。`,
  '- 2026.07仅94父体，是小范围样本；已合并展示，但不参与同比、环比、累计或核心结论。', '',
  '## 必须同时阅读的限制', '',
  '- 2025为含ASIN/父ASIN的行级导出（含变体行），2026为父ASIN去重快照；跨年百分比只表示方向，不是严格同口径同比。',
  '- 2022-2025 BSR Top100是源表行代理，不等于100个可验证的独立Listing；2025.05存在严重重复名次异常。',
  '- 2026代表行规则：同父体先取最小可解析小类BSR；同名次优先销量/销售额字段完整行，再按源表顺序稳定决胜；不合计子体。',
  '- SKU平均标价排除空值和不可解析值；销量加权成交均价=销售额/销量。', '',
  '## 建议', '',
  '1. 以2026同口径月度趋势安排旺季前4-8周补货与投放，不用跨口径百分比反推精确市场增长率。',
  '2. 保持PP流量盘，同时用非PP高客单产品修复销售额和价格结构。',
  '3. GENIMO采用“1个头部锚点 + 3-5个中部利润层 + 4-8个尾部测试池”，按完整报告中的毛利、TACOS、BSR和库存门槛晋级或退出。', '',
  '4. 工艺或包装改动（如包边、包装袋或地钉）先在低风险颜色或4x6/5x8小批量验证；保留批次标记，跟踪散边、卷边、破损和退货原因，累计300-500单后再决定扩大，8x10、9x12及超大尺寸暂保留稳定边缘处理。', '',
  `生成时间：${data.generatedAt}`,
].join('\n');
let htmlOutput = html
  .replace(/<div id="dashboard" class="scope-notice"><span>◎<\/span><div>[\s\S]*?<\/div><\/div>/, `<div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>${esc(coverageScopeText)}</div></div>`)
  .replace(/<div class="metrics-grid">[\s\S]*?<\/div><section id="definitions">/, dashboardMetricsHtml + leadershipBenchmarkHtml + trendChartsHtml + '<section id="definitions">')
  .replace(/<section id="definitions">[\s\S]*?<\/section>/, `<section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>。2026同父体优先取最小可解析名次；同名次优先销量/销售额字段完整行，再按源表顺序稳定决胜，代表行指标不合计子体；每分类每月最多100条，头中尾和五档复用同一Top100集合。</li><li>2022-2025源表没有ASIN/父ASIN，历史BSR Top100只能定义为按BSR排序后截取的100条源表行代理，不能证明是100个独立Listing；2026为父体/ASIN Listing池，跨年BSR变化仅作方向性参考。</li><li>PP塑料地垫：标题按不区分大小写的完整单词 <code>plastic</code>（单词边界）筛选，空标题按空字符串；2026同父体任一变体命中即归PP。</li><li>高客单非PP：排除PP父体后的全部商品（SPEC 7.5），不再叠加材质关键词或价格门槛。</li><li>SKU平均标价只统计非空、可解析的价格，缺失值不按0计入；销量加权成交均价=销售额/销量。月度MOM（用户口径）=今年X月 vs 去年X月同月；月度环比=本月 vs 上月。两种比较在全部月度表分别显示基准月份。2025是无ASIN的行级导出，2026是父ASIN去重快照；量级接近不等于统计单元一致，跨年变化仅作方向性参考。</li><li>2025.05（主表导出日 2025-06-19，无ASIN列）源数据存在小类BSR同值重复：BSR=17 重复112行、BSR=23 重复125行、BSR=58 重复156行，按小类BSR取前100后全部落入1-20。因此相关中部/尾部变化显示"无对应数据"；2026.05头部相对该异常基准的变化仅供参考。</li></ul>${bsrQualityHtml}${bsrMultiAuditHtml}<p class="note">${esc(coverageScopeText)}</p></section>`)
  .replace(/<section id="anomaly">[\s\S]*?<\/section>/, '')
  .replace('销量、销售额、双均价 · 小类BSR前100 · 月度MoM与年度YoY', '销量、销售额、双均价 · 小类BSR Top100 · 用户定义MOM与月度环比')
  .replace('<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a>', '<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#genimo-products"><span class="nav-icon">Top</span>GENIMO主力ASIN</a>')
  .replace('<a href="#definitions"><span class="nav-icon">径</span>数据口径</a>', '<a href="#visuals"><span class="nav-icon">图</span>核心趋势</a><a href="#definitions"><span class="nav-icon">径</span>数据口径</a>')
  .replace('<a href="#overall"><span class="nav-icon">整</span>整体市场</a>', '<a href="#overall"><span class="nav-icon">整</span>整体市场</a><a href="#leadership-benchmark"><span class="nav-icon">验</span>领导验收基准</a>')
  .replace('<a href="#anomaly"><span class="nav-icon">!</span>参考附录</a>', '')
  .replace('<a href="#coverage"><span class="nav-icon">核</span>需求覆盖</a>', '<a href="#coverage"><span class="nav-icon">核</span>需求覆盖核对</a>')
  .replace(/<small>[^<]*444万单元格核验<\/small>/, `<small>${analysisMonths.length}个月 · ${fmt(verifiedDataCellCount)}数据单元格</small>`)
  .replace(/<strong>\d{6} — \d{6}<\/strong>/, `<strong>${fmtMonth(analysisMonths[0])} — ${fmtMonth(analysisMonths[analysisMonths.length - 1])}</strong>`)
  .replace('</main></div></div><script>', coverageHtml + '</main></div></div><script>')
  .replace(/<section id="insights">[\s\S]*?<\/section>/, cohortHtml + genimoProductsHtml + renderedInsightHtml + forecastHtml + referenceHtml.replace('</ul></section>', planReferenceHtml + '</ul></section>'))
  // insightHtml 是历史长模板，最终渲染前再做一次口径清理，确保连续环比不进入交付 HTML。
  .replace('；2026月度环比保持父体口径一致。', '；月度MOM/环比统一采用跨年同月基准。')
  .replace('月度MOM = 今年X月 vs 去年X月（跨年同月，如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）。', '月度MOM/环比（用户口径）= 今年X月 vs 去年X月同月；本次不计算或展示本月 vs 上月的连续环比。')
  .replace('2022-2025源表没有ASIN/父ASIN', '2022-2025源表的ASIN/父ASIN以富文本超链接保存，已恢复为显示值')
  .replace('2025是无ASIN的行级导出', '2025是含ASIN/父ASIN的行级导出')
  .replace('2025.05（主表导出日 2025-06-19，无ASIN列）', '2025.05（主表导出日 2025-06-19，ASIN/父ASIN为富文本超链接）')
  .replace(/、环比 -，为核心期峰值；/g, '、为核心期峰值；')
  .replace(/，环比 -，为核心期峰值；/g, '，为核心期峰值；')
  .replace(/，月度环比 - \/ -；/g, '；')
  .replace(/月度MOM（用户口径）=今年X月 vs 去年X月同月；月度环比=本月 vs 上月。两种比较在全部月度表分别显示基准月份。/g, '月度MOM/环比（用户口径）=今年X月 vs 去年X月同月；本次不计算或展示本月 vs 上月的连续环比。')
  .replace(/月度MOM（用户口径）= 今年X月 vs 去年X月同月（如 2025\.01 vs 2024\.01）；月度环比 = 本月 vs 上月；年度YOY = 年度同周期对比。/g, '月度MOM/环比（用户口径）= 今年X月 vs 去年X月同月（如 2025.01 vs 2024.01）；本次不计算或展示本月 vs 上月的连续环比；年度YOY = 年度同周期对比。')
  .replace(/用户定义MOM与月度环比/g, '用户定义MOM/环比')
  .replace('整体月度MOM与环比', '整体月度MOM/环比')
  .replace('所有月度表分别显示MOM/环比基准月份', '所有月度表显示跨年同月MOM/环比基准月份')
  .replace('各分类 BSR Top100月度、年度表', '各分类 BSR Top100月度、年度表（仅跨年同月MOM/环比）')
  .replace('</style></head>', '<style>.trend-body{padding:10px 0 2px;color:var(--muted)}.trend-body p{margin:7px 0;font-size:14px}.insight-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.insight-grid article{padding:15px;border:1px solid var(--line);background:var(--surface-soft)}.insight-grid h3{margin:0 0 7px}.insight-grid p{margin:0;color:var(--muted);font-size:14px}.chart-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.chart-card{min-width:0;margin:0;padding:14px;border:1px solid var(--line);background:var(--surface-soft)}.chart-card figcaption{display:flex;flex-direction:column}.chart-card figcaption span{color:var(--muted);font-size:12px}.chart-card svg{display:block;width:100%;height:auto;margin-top:10px;overflow:visible}.chart-card text{fill:var(--muted);font:12px var(--mono)}.chart-axis{stroke:var(--line-strong)}.chart-line{fill:none;stroke:var(--text);stroke-width:3;stroke-linecap:round;stroke-linejoin:round}.chart-card circle{fill:var(--surface);stroke:var(--text);stroke-width:2}@media(max-width:1000px){.chart-grid{grid-template-columns:1fr}}@media(max-width:700px){.insight-grid{grid-template-columns:1fr}}</style></head>');

fs.mkdirSync(path.dirname(HTML_PATH), { recursive: true });
fs.writeFileSync(JSON_PATH, JSON.stringify(data, null, 2), 'utf8');
fs.writeFileSync(MD_PATH, md.join('\n'), 'utf8');
fs.writeFileSync(QUICK_PATH, quickMd, 'utf8');
fs.writeFileSync(HTML_PATH, htmlOutput, 'utf8');
if (competitorDb) competitorDb.close();
db.close();
console.log('Generated: ' + path.relative(ROOT, JSON_PATH));
console.log('Generated: ' + path.relative(ROOT, MD_PATH));
console.log('Generated: ' + path.relative(ROOT, QUICK_PATH));
console.log('Generated: ' + path.relative(ROOT, HTML_PATH));
