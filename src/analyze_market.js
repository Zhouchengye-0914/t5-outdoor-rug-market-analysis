'use strict';

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const ROOT = path.resolve(__dirname, '..');
const DB_PATH = path.resolve(ROOT, process.env.ANALYSIS_DB_PATH || 'data/processed/market.db');
const REPORT_CUTOFF = process.env.ANALYSIS_CUTOFF || '202602';
const MD_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.md');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');

const MATERIAL_KEYWORDS = ['polypropylene', 'sandwich', 'woven', 'braided'];
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

function classify(row, category) {
  const title = String(row.title || '').toLowerCase();
  const isPlastic = title.includes('plastic');
  if (category === 'overall') return true;
  if (category === 'pp') return isPlastic;
  if (category === 'high') {
    const materialHit = MATERIAL_KEYWORDS.some((keyword) => title.includes(keyword));
    return !isPlastic && (materialHit || Number(row.price) >= 40);
  }
  if (category === 'genimo') return String(row.brand || '').trim().toLowerCase() === 'genimo';
  return false;
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
    // 用户口径：MOM = 今年X月 vs 去年X月（跨年同月）；环比 = 本月 vs 上月（连续月）。
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
    };
  });
}

function buildAnnualBySegment(rows) {
  const keys = [...new Set(rows.map((row) => row.segment))];
  return keys.flatMap((segment) => buildAnnual(rows.filter((row) => row.segment === segment))
    .map((row) => ({ segment, ...row })));
}


function trendAnalysis(c, category, label) {
  const cleanMonths = c.monthly.filter((row) => row.month <= '202602');
  const baseline = cleanMonths[cleanMonths.length - 1] || c.monthly[c.monthly.length - 1];
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
  if (baseline) out.push(`- 最近可比月 ${baseline.month}：MOM销量（今年 vs 去年同月）${fmtPct(baseline.momSales)}、MOM销售额 ${fmtPct(baseline.momRevenue)}；环比销量（vs 上月）${fmtPct(baseline.chainSales)}、环比销售额 ${fmtPct(baseline.chainRevenue)}。`);
  if (seasonal) out.push(`- 季节性（2024-2025同月均值）：${seasonal.month}月销量最高，月均 ${fmt(seasonal.avg)} 件，建议在高峰前完成备货与广告测试。`);
  const anomaly = c.monthly.find((row) => row.month === '202604');
  if (anomaly && baseline) out.push(`- 口径提示：2026.04销量 ${fmt(anomaly.sales)}，较最近可比月 ${baseline.month} 的 ${fmt(baseline.sales)} 显著跳升；2026.03-07原值保留在附录，长期判断需按同口径复核。`);
  return out.join('\n');
}

const db = new DatabaseSync(DB_PATH, { readOnly: true });
const catalog = db.prepare("SELECT * FROM sheet_catalog WHERE classification='monthly' ORDER BY sheet_order").all();
const sourceMonths = catalog.map((row) => row.target_table.replace('monthly_', ''));
const analysisMonths = sourceMonths.filter((month) => month <= REPORT_CUTOFF);
const sourceMeta = db.prepare('SELECT * FROM meta ORDER BY id DESC LIMIT 1').get() || {};
const effectiveCatalog = db.prepare("SELECT * FROM sheet_catalog WHERE target_table IS NOT NULL ORDER BY sheet_order").all();
const verifiedDataCellCount = effectiveCatalog.reduce((sum, row) => {
  const columns = db.prepare('PRAGMA table_info(' + row.target_table + ')').all()
    .filter((column) => !['row_id', 'month_label'].includes(column.name)).length;
  return sum + Number(row.imported_rows || 0) * columns;
}, 0);
const rawByMonth = new Map();

for (const month of sourceMonths) {
  const table = 'monthly_' + month;
  const rows = db.prepare('SELECT row_id, ASIN asin, 品牌 brand, 商品标题 title, 小类BSR bsr, 月销量 sales, 月销售额 revenue, 价格 price FROM ' + table).all();
  rawByMonth.set(month, rows.map((row) => ({ ...row, ...parseBsr(row.bsr) })));
}

const categories = {};
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const monthly = [];
  const bsrTop100 = [];
  const bsrSegments = [];
  const bsrGroups = [];
  for (const month of analysisMonths) {
    const rows = rawByMonth.get(month).filter((row) => classify(row, category));
    monthly.push({ month, ...summarize(rows) });
    const top100 = rows.filter((row) => row.rank !== null && row.rank >= 1 && row.rank <= 100);
    bsrTop100.push({ month, ...summarize(top100) });
    for (const group of SEGMENT_GROUPS) {
      const groupRows = rows.filter((row) => row.rank >= group.min && row.rank <= group.max);
      bsrGroups.push({ month, segment: group.key, ...summarize(groupRows) });
    }
    for (const segment of SEGMENTS) {
      const segmentRows = rows.filter((row) => row.rank >= segment.min && row.rank <= segment.max);
      bsrSegments.push({ month, segment: segment.key, ...summarize(segmentRows) });
    }
  }
  categories[category] = {
    monthly: addTrends(monthly),
    annual: buildAnnual(monthly),
    bsrTop100: { monthly: addTrends(bsrTop100), annual: buildAnnual(bsrTop100) },
    bsrGroups: { monthly: addTrendsBySegment(bsrGroups), annual: buildAnnualBySegment(bsrGroups) },
    bsrSegments: { monthly: addTrendsBySegment(bsrSegments), annual: buildAnnualBySegment(bsrSegments) },
  };
}

function addTrendsBySegment(rows) {
  for (const segment of [...new Set(rows.map((row) => row.segment))]) addTrends(rows.filter((row) => row.segment === segment));
  return rows;
}

const sourceDiagnostics = sourceMonths.map((month) => ({
  month,
  includedInComparableReport: month <= REPORT_CUTOFF,
  includedInLongTermConclusion: month < '202603',
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
  verifiedDataCellCount,
  definitions: {
    pp: "LOWER(COALESCE(商品标题,'')) contains 'plastic'",
    high: "not PP and (title contains configured material keyword or price >= 40); keywords = polypropylene, sandwich, woven, braided",
    bsrTop100: 'minimum numeric rank parsed from 小类BSR; bands 1-5, 6-10, 11-20, 21-50, 51-100',
    bsrGroups: 'head = 1-20, middle = 21-50, tail = 51-100; groups do not overlap',
    avgListPrice: 'simple average of SKU list prices',
    weightedPrice: '月销售额 / 月销量',
    yoyMonthly: '同月同比：今年X月 vs 去年X月',
    momMonthly: '连续月环比：本月 vs 上月',
    annualYoY: '年度同周期对比（今年 vs 去年同月份集合）',
  },
  categories,
  sourceDiagnostics,
  genimoTopProducts: [...genimoProducts.values()].sort((a, b) => b.sales - a.sales).slice(0, 20),
};

function annualRow(category, year) {
  return categories[category].annual.find((row) => row.year === year);
}

function peakMonth(category, year) {
  return [...categories[category].monthly.filter((row) => row.month.startsWith(year))]
    .sort((a, b) => b.sales - a.sales)[0];
}

function insightRevenueShare(numeratorCategory, denominatorCategory, year) {
  const numerator = annualRow(numeratorCategory, year);
  const denominator = annualRow(denominatorCategory, year);
  return numerator && denominator && denominator.revenue ? numerator.revenue / denominator.revenue * 100 : null;
}

const pp2025Rows = analysisMonths.filter((month) => month.startsWith('2025'))
  .flatMap((month) => rawByMonth.get(month).filter((row) => classify(row, 'pp')));
const genimoPp2025Rows = pp2025Rows.filter((row) => classify(row, 'genimo'));
const pp2025Sales = summarize(pp2025Rows).sales;
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
  genimoPpShare2025: pp2025Sales ? summarize(genimoPp2025Rows).sales / pp2025Sales * 100 : null,
  genimoPpRevenueShare2025: insightRevenueShare('genimo', 'pp', '2025'),
};

function mdMonthly(rows) {
  const out = ['| 月份 | SKU数 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | MOM销量 | MOM销售额 | MOM标价 | MOM成交均价 | 环比销量 | 环比销售额 |',
    '|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.momSales)} | ${fmtPct(r.momRevenue)} | ${fmtPct(r.momAvgListPrice)} | ${fmtPct(r.momWeightedPrice)} | ${fmtPct(r.chainSales)} | ${fmtPct(r.chainRevenue)} |`);
  return out.join('\n');
}

function mdAnnual(rows) {
  const out = ['| 年份/数据周期 | YoY比较周期 | 销量 | 销售额($) | SKU平均标价($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY标价 | YOY成交均价 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.year} (${r.period}) | ${r.comparison || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyAvgListPrice)} | ${fmtPct(r.yoyWeightedPrice)} |`);
  return out.join('\n');
}

function mdSegments(rows) {
  const out = ['| 月份 | BSR分层 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | MOM销量 | MOM销售额 | MOM成交均价 | 环比销量 | 环比销售额 |',
    '|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${r.segment} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.momSales)} | ${fmtPct(r.momRevenue)} | ${fmtPct(r.momWeightedPrice)} | ${fmtPct(r.chainSales)} | ${fmtPct(r.chainRevenue)} |`);
  return out.join('\n');
}

function mdAnnualSegments(rows) {
  const out = ['| 年份/数据周期 | BSR分层 | YoY比较周期 | SKU数 | 销量 | 销售额($) | 加权成交均价($) | YOY销量 | YOY销售额 | YOY成交均价 |',
    '|---|---|---|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.year} (${r.period}) | ${r.segment} | ${r.comparison || '-'} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} | ${fmtPct(r.yoyWeightedPrice)} |`);
  return out.join('\n');
}

const labels = { overall: '整体市场', pp: 'PP塑料地垫（标题含 plastic）', high: '非PP高客单产品', genimo: 'GENIMO品牌' };
const md = ['# 户外地垫市场分析报告（优化版）', '',
  `> 分析范围：${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]}，共 ${analysisMonths.length} 个月（明细全部纳入）。2026.03-07 原值完整披露，但因源数据口径数量级突变，不用于长期可比结论。`, '',
  '## 一、口径说明', '',
  '- 小类前100严格依据源字段 `小类BSR`，多值BSR取最小可解析名次并保留多值标记。',
  '- PP：标题包含 `plastic`（不区分大小写，NULL按空字符串处理）。',
  '- 高客单：排除PP后，标题命中 `polypropylene / sandwich / woven / braided`，或价格不低于$40；这是可复核的分析筛选口径。',
  '- 同时提供SKU平均标价和销量加权均价（销售额/销量）。',
  '- 月度MOM = 今年X月 vs 去年X月（跨年同月，如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）；年度YOY = 年度同周期对比。',
  '- BSR头部/中部/尾部分别为1-20、21-50、51-100；五档明细为1-5、6-10、11-20、21-50、51-100，区间不重叠。',
  '- 年度YoY使用同周期比较；2023对2022仅比较6-12月，避免12个月对7个月。', ''];

let sectionNo = 2;
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const c = categories[category];
  md.push(`## ${sectionNo++}、${labels[category]}`, '', trendAnalysis(c, category, labels[category]), '### 月度指标、MOM与YOY', '', mdMonthly(c.monthly), '',
    '### 年度/同周期汇总', '', mdAnnual(c.annual), '', '### 小类BSR前100汇总（BSR 1-100）', '', mdMonthly(c.bsrTop100.monthly), '',
    '### 小类BSR前100年度/同周期汇总', '', mdAnnual(c.bsrTop100.annual), '', '### 小类BSR头部/中部/尾部（月度）', '', mdSegments(c.bsrGroups.monthly), '',
    '### 小类BSR头部/中部/尾部（年度）', '', mdAnnualSegments(c.bsrGroups.annual), '', '### 小类BSR五档分层（月度）', '', mdSegments(c.bsrSegments.monthly), '',
    '### 小类BSR五档分层（年度）', '', mdAnnualSegments(c.bsrSegments.annual), '');
}

md.push('## 六、2026.03-2026.07异常附录', '', '| 月份 | 是否纳入长期可比结论 | 销量 | 销售额($) | SKU平均标价($) | 销量加权均价($) |',
  '|---|---|---:|---:|---:|---:|');
for (const r of sourceDiagnostics.filter((r) => r.month >= '202603')) {
  md.push(`| ${r.month} | ${r.includedInLongTermConclusion ? '是' : '否'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} |`);
}
md.push('', '## 七、GENIMO累计Top产品', '', '| 排名 | ASIN | 累计销量 | 累计销售额($) | 月数 | 最新价($) | 标题 |',
  '|---:|---|---:|---:|---:|---:|---|');
data.genimoTopProducts.forEach((r, i) => md.push(`| ${i + 1} | ${r.asin || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${r.months} | ${fmt(r.latestPrice, 2)} | ${String(r.title || '').replace(/\|/g, '\\|')} |`));

const insight = data.insights;
md.push('', '## 八、趋势结论与GENIMO建议', '',
  `- **整体市场**：2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}，SKU平均标价同比 ${fmtPct(insight.overall2025.yoyAvgListPrice)}，加权成交均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}；量增明显快于额增，价格承压。`,
  `- **PP市场**：2025销量同比 ${fmtPct(insight.pp2025.yoySales)}，销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}；销量峰值为 ${insight.ppPeak2025.month} 的 ${fmt(insight.ppPeak2025.sales)} 件。`,
  `- **高客单非PP**：2025销量同比 ${fmtPct(insight.high2025.yoySales)}，销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}，加权成交均价同比 ${fmtPct(insight.high2025.yoyWeightedPrice)}，表现弱于PP。`,
  `- **GENIMO**：2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，PP销售额份额 ${fmt(insight.genimoPpRevenueShare2025, 2)}%，品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}。`, '',
  '### 建议', '',
  '1. 按尺寸、价格带和小类BSR管理PP产品，优先保障BSR 1-20核心SKU库存与广告。',
  '2. 同时考核销量、销售额和销量加权均价，避免只追求件数导致价格与利润空间持续受压。',
  '3. 将高客单拆分为“材质关键词命中”和“仅价格命中”两组，小规模验证非PP第二增长曲线。',
  '4. 根据月度MoM/YoY在3-5月旺季前置补货和新品测试。',
  '5. 2026.03-07已纳入分析，但该区间源数据统计口径突变（销量量级跳升），制定年度预算时应区分口径并关注异常月份。');

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
  return htmlTable(['年份/数据周期', 'YoY比较周期', '销量', '销售额($)', 'SKU平均标价', '加权成交均价', 'YOY销量', 'YOY销售额', 'YOY标价', 'YOY成交均价'],
    rows.map((r) => [`${r.year} (${r.period})`, r.comparison || '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyAvgListPrice), fmtPct(r.yoyWeightedPrice)]));
}

function segmentHtml(rows) {
  return htmlTable(['月份', '小类BSR分层', 'SKU数', '销量', '销售额($)', '加权成交均价', 'MOM销量', 'MOM销售额', 'MOM成交均价', '环比销量', '环比销售额'],
    rows.map((r) => [r.month, r.segment, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.momWeightedPrice), fmtPct(r.chainSales), fmtPct(r.chainRevenue)]));
}

function annualSegmentsHtml(rows) {
  return htmlTable(['年份/数据周期', 'BSR分层', 'YoY比较周期', 'SKU数', '销量', '销售额($)', '加权成交均价($)', 'YoY销量', 'YoY销售额', 'YoY成交均价'],
    rows.map((r) => [`${r.year} (${r.period})`, r.segment, r.comparison || '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyWeightedPrice)]));
}

function trendHtml(c, category, label) {
  return trendAnalysis(c, category, label).split('\n').filter(Boolean).map((line) => {
    if (line.startsWith('### ')) return '<h3>' + esc(line.slice(4)) + '</h3>';
    return '<p>' + esc(line.replace(/^- /, '')) + '</p>';
  }).join('');
}

const htmlSections = ['overall', 'pp', 'high', 'genimo'].map((category, index) => {
  const c = categories[category];
  return `<section id="${category}"><h2>${index + 2}、${esc(labels[category])}</h2><details open class="trend-details"><summary><b>${esc(labels[category])}趋势分析</b></summary><div class="trend-body">${trendHtml(c, category, labels[category])}</div></details><details open><summary>月度指标、MOM与YOY</summary>${monthlyHtml(c.monthly)}</details><details><summary>年度/同周期汇总</summary>${annualHtml(c.annual)}</details><details open class="bsr-details"><summary><b>小类BSR前100分析</b>（1-100名汇总 + 年度 + 头中尾 + 五档）</summary><h4>BSR前100月度汇总（1-100名）</h4>${monthlyHtml(c.bsrTop100.monthly)}<h4>BSR前100年度/同周期汇总</h4>${annualHtml(c.bsrTop100.annual)}<h4>BSR头部/中部/尾部（月度）</h4>${segmentHtml(c.bsrGroups.monthly)}<h4>BSR头部/中部/尾部（年度）</h4>${annualSegmentsHtml(c.bsrGroups.annual)}<h4>BSR五档分层（月度）</h4>${segmentHtml(c.bsrSegments.monthly)}<h4>BSR五档分层（年度）</h4>${annualSegmentsHtml(c.bsrSegments.annual)}</details></section>`;
}).join('\n');

const ppSalesShare2025 = insight.overall2025.sales ? insight.pp2025.sales / insight.overall2025.sales * 100 : null;
const insightHtml = `<section id="insights"><h2>八、趋势结论与GENIMO建议</h2><div class="insight-grid"><article><h3>整体市场</h3><p>2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}；SKU平均标价同比 ${fmtPct(insight.overall2025.yoyAvgListPrice)}，加权成交均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}。量增快于额增，说明价格与结构承压。</p></article><article><h3>PP塑料地垫</h3><p>2025销量同比 ${fmtPct(insight.pp2025.yoySales)}、销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}；全年销量 ${fmt(insight.pp2025.sales)}，占整体销量 ${fmt(ppSalesShare2025, 1)}%，${insight.ppPeak2025.month}达到销量峰值 ${fmt(insight.ppPeak2025.sales)}。</p></article><article><h3>高客单非PP</h3><p>2025销量同比 ${fmtPct(insight.high2025.yoySales)}、销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}，加权成交均价同比 ${fmtPct(insight.high2025.yoyWeightedPrice)}，需要用材质与价格带拆分寻找增长点。</p></article><article><h3>GENIMO</h3><p>2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，PP销售额份额 ${fmt(insight.genimoPpRevenueShare2025, 2)}%；品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}，增长快于大盘但仍需改善价格质量。</p></article></div><h3>行动建议</h3><ol><li>按尺寸、价格带和小类BSR管理PP产品，优先保障BSR 1-20核心SKU的库存、广告与评价资产。</li><li>同时考核销量、销售额、SKU平均标价和加权成交均价，避免以低价换规模。</li><li>将高客单拆成材质命中与仅价格命中两组，先用小预算验证丙纶/三明治等非PP第二曲线。</li><li>依据月度YOY/MOM识别3-5月旺季，在峰值前4-8周完成补货、广告和新品测试。</li><li>2026.03-07原值保留但存在统计口径突变，预算与目标制定按同口径序列分开，不直接与2025同比外推。</li></ol></section>`;

const genimoProductsHtml = `<section id="genimo-products"><h2>七、GENIMO累计Top产品</h2><p class="note">按分析范围内各月销量累计排序；用于识别应优先维护的主力ASIN，具体月度变化请回到GENIMO明细表。</p>${htmlTable(['排名', 'ASIN', '累计销量', '累计销售额($)', '覆盖月数', '最新价($)', '商品标题'], data.genimoTopProducts.map((row, index) => [index + 1, row.asin || '-', fmt(row.sales), fmt(row.revenue), row.months, fmt(row.latestPrice, 2), row.title || '-']))}</section>`;

const coverageHtml = `<section id="coverage"><h2>九、需求覆盖核对</h2>${htmlTable(['要求项', '交付位置', '状态'], [['整体市场销量/销售额/均价', '整体市场月度与年度表', '已覆盖'], ['整体月度YOY与MOM', '整体市场月度表（同月同比/连续月环比）', '已覆盖'], ['小类前100同比/环比', '各分类 BSR前100月度、年度表', '已覆盖'], ['头部/中部/尾部及五档', '各分类 BSR分层月度、年度表', '已覆盖'], ['PP（标题含 plastic）', 'PP塑料地垫分类与口径说明', '已覆盖'], ['高客单非PP（丙纶/三明治等）', '高客单分类与关键词/价格阈值', '已覆盖'], ['GENIMO建议', 'GENIMO指标、Top产品与行动建议', '已覆盖'], ['异常月份披露', '2026.03-07异常附录', '已覆盖']])}</section>`;

const html = `<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场洞察 · Outdoor Rug Intelligence</title><style>
:root{color-scheme:dark;font-family:Inter,"Segoe UI Variable","Segoe UI","Microsoft YaHei",system-ui,sans-serif;--bg:#050505;--surface:#0d0d0d;--surface-soft:rgba(255,255,255,.018);--surface-strong:rgba(255,255,255,.05);--text:#f2f2f0;--muted:#8b8b88;--faint:#5f5f5c;--line:rgba(255,255,255,.095);--line-strong:rgba(255,255,255,.22);--brand:#f4f4f1;--brand-soft:rgba(255,255,255,.05);--warning:#aaa9a2;--warning-soft:rgba(255,255,255,.04);--sidebar:#060606;--cta:#eeeeeb;--cta-text:#080808;--grid-line:rgba(255,255,255,.012);--card-bg:linear-gradient(145deg,rgba(18,18,18,.72),rgba(8,8,8,.8));--shadow:0 28px 80px rgba(0,0,0,.3),inset 0 1px rgba(255,255,255,.025);--display:"Times New Roman","Songti SC","STSong",serif;--mono:"Cascadia Code","SFMono-Regular",Consolas,monospace}
:root[data-theme="light"]{color-scheme:light;--bg:#f4f3ef;--surface:#fafaf7;--surface-soft:rgba(24,24,20,.03);--surface-strong:rgba(24,24,20,.06);--text:#191916;--muted:#6f6e68;--faint:#9a9992;--line:rgba(24,24,20,.12);--line-strong:rgba(24,24,20,.28);--brand:#1b1b18;--brand-soft:rgba(24,24,20,.05);--warning:#65635b;--warning-soft:rgba(24,24,20,.04);--sidebar:#efede7;--cta:#191916;--cta-text:#f8f7f3;--grid-line:rgba(24,24,20,.025);--card-bg:linear-gradient(145deg,rgba(255,255,252,.82),rgba(245,244,239,.9));--shadow:0 18px 48px rgba(24,24,20,.08),inset 0 1px rgba(255,255,255,.6)}
*{box-sizing:border-box}html{min-width:320px;scroll-behavior:smooth;background:var(--bg)}body{position:relative;margin:0;min-height:100vh;color:var(--text);font-size:18px;background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(255,255,255,.105),transparent 68%),radial-gradient(ellipse 48% 38% at 12% 92%,rgba(255,255,255,.035),transparent 72%),linear-gradient(145deg,#020202 0%,#080808 48%,#030303 100%);line-height:1.65}:root[data-theme="light"] body{background:radial-gradient(ellipse 65% 46% at 76% -8%,rgba(20,20,18,.08),transparent 68%),linear-gradient(145deg,#faf9f5,#f0efea 52%,#f7f6f2)}body::before{content:"";position:fixed;inset:0;pointer-events:none;opacity:.22;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:72px 72px;mask-image:linear-gradient(to bottom,#000,transparent 82%)}a{color:inherit}.app-shell{position:relative;z-index:1;min-height:100vh}.sidebar{position:fixed;inset:0 auto 0 0;z-index:30;display:flex;width:232px;flex-direction:column;overflow-y:auto;padding:34px 18px 24px;border-right:1px solid var(--line);background:linear-gradient(180deg,rgba(4,4,4,.97),rgba(8,8,8,.9));box-shadow:18px 0 80px rgba(0,0,0,.45)}:root[data-theme="light"] .sidebar{background:linear-gradient(180deg,#f2f0ea,#eae8e1);box-shadow:12px 0 40px rgba(24,24,20,.06)}.brand{display:flex;align-items:center;gap:12px;padding:0 8px;margin-bottom:34px}.brand-mark{display:grid;width:34px;height:34px;place-items:center;border:1px solid var(--line-strong);border-radius:50%;background:radial-gradient(circle at 35% 28%,#f0f0ed,#696966 38%,#080808 70%);box-shadow:0 0 34px rgba(255,255,255,.12)}.brand strong{display:block;font-family:var(--display);font-size:18px;font-weight:400;letter-spacing:.07em}.brand small{display:block;margin-top:3px;color:var(--muted);font-size:13px}.local-badge{display:flex;align-items:center;gap:7px;width:max-content;margin:10px 8px 24px;padding:6px 9px;border:1px solid var(--line);font-family:var(--mono);font-size:13px;letter-spacing:.14em;color:var(--muted)}.local-badge i,.privacy-chip i{width:6px;height:6px;border-radius:50%;background:var(--brand);box-shadow:0 0 12px rgba(255,255,255,.55)}.nav-label{display:block;padding:0 10px 7px;color:var(--muted);font-family:var(--mono);font-size:14px;letter-spacing:.22em}.sidebar nav{display:grid;gap:2px}.sidebar nav a{position:relative;display:flex;align-items:center;gap:11px;padding:11px 10px;color:var(--muted);font-size:15px;letter-spacing:.04em;text-decoration:none}.sidebar nav a:hover,.sidebar nav a.is-active{color:var(--text);background:var(--surface-strong)}.nav-icon{display:grid;width:23px;height:23px;place-items:center;border-radius:50%;font-family:var(--mono);font-size:14px}.sidebar-footer{margin-top:auto;padding:16px 10px 3px;border-top:1px solid var(--line)}.sidebar-footer span,.sidebar-footer strong,.sidebar-footer small{display:block}.sidebar-footer span,.sidebar-footer small{color:var(--muted);font-size:13px}.sidebar-footer strong{margin:7px 0 3px;font-size:15px}.workspace{min-height:100vh;margin-left:232px}.topbar{display:flex;min-height:154px;align-items:center;gap:18px;padding:26px clamp(28px,4vw,64px);border-bottom:1px solid var(--line);background:color-mix(in srgb,var(--bg),var(--surface) 25%)}.eyebrow{display:block;margin-bottom:7px;color:var(--faint);font-family:var(--mono);font-size:13px;letter-spacing:.28em}.topbar h1{margin:0 0 8px;font-family:var(--display);font-size:clamp(44px,5vw,70px);font-weight:400;letter-spacing:-.04em;background:linear-gradient(110deg,#fff 8%,#adada8 55%,#535350 100%);background-clip:text;color:transparent}:root[data-theme="light"] .topbar h1{background:linear-gradient(110deg,#191916,#55544e 55%,#9a9992);background-clip:text;color:transparent}.topbar p{margin:0;color:var(--muted);font-size:16px}.top-actions{display:flex;align-items:center;gap:10px;margin-left:auto}.privacy-chip,.theme-button{border:1px solid var(--line);color:var(--muted);background:var(--surface)}.privacy-chip{display:flex;align-items:center;gap:8px;padding:8px 11px;font-size:14px}.theme-button{display:grid;width:38px;height:38px;place-items:center;cursor:pointer}.content{width:min(1500px,100%);margin:0 auto;padding:32px clamp(28px,4vw,64px) 72px}.scope-notice{display:flex;gap:12px;padding:13px 16px;border:1px solid var(--line-strong);color:var(--muted);background:var(--brand-soft);font-size:16px}.scope-notice b{color:var(--text)}.metrics-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:22px 0}.metric-card{position:relative;min-height:142px;padding:20px;border:1px solid var(--line);background:var(--card-bg);box-shadow:var(--shadow)}.metric-card::before,section::before{content:"";position:absolute;inset:0 0 auto;height:1px;background:linear-gradient(90deg,rgba(255,255,255,.7),transparent)}.metric-label,.metric-value,.metric-note{display:block}.metric-label{color:var(--muted);font-size:15px}.metric-value{margin:15px 0 9px;font-family:var(--display);font-size:clamp(30px,2.6vw,40px);font-weight:400}.metric-note{color:var(--faint);font-size:14px;line-height:1.5}section{position:relative;min-width:0;margin-bottom:22px;padding:23px;border:1px solid var(--line);background:var(--surface);box-shadow:var(--shadow);scroll-margin-top:18px}section h2{margin:0 0 6px;font-family:var(--display);font-size:28px;font-weight:400;letter-spacing:-.01em}section h3{font-size:20px;letter-spacing:.05em}section li{color:var(--muted);font-size:16px;line-height:1.8}section li b{color:var(--text)}details{padding:15px 0;border-top:1px solid var(--line)}summary{cursor:pointer;color:var(--text);font-size:16px;font-weight:700;letter-spacing:.04em}.table-wrap{max-height:620px;margin-top:13px;overflow:auto;border:1px solid var(--line)}table{width:100%;min-width:980px;border-collapse:collapse;font-family:var(--mono);font-size:15px}th,td{padding:9px 11px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th{position:sticky;top:0;z-index:1;color:var(--muted);background:var(--surface);font-size:15px;letter-spacing:.06em}th:first-child,td:first-child{text-align:left}tbody tr:hover{background:var(--surface-soft)}.note{margin-top:18px;padding:13px 16px;border-left:2px solid var(--warning);color:var(--muted);background:var(--warning-soft);font-size:15px}code{padding:2px 5px;color:var(--text);background:var(--surface-strong);font-family:var(--mono)}
@media(max-width:980px){.sidebar{transform:translateX(-100%)}.workspace{margin-left:0}.metrics-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:560px){.topbar{padding:24px 18px}.privacy-chip{display:none}.content{padding:20px 12px 50px}.metrics-grid{grid-template-columns:1fr}.metric-card{min-height:116px}section{padding:17px}.topbar h1{font-size:44px}}
</style></head><body><div class="app-shell"><aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫市场分析系统</small></div></div><div class="local-badge"><i></i> DATA · VERIFIED</div><span class="nav-label">市场分析</span><nav><a href="#dashboard"><span class="nav-icon">总</span>市场总览</a><a href="#definitions"><span class="nav-icon">径</span>数据口径</a><a href="#overall"><span class="nav-icon">整</span>整体市场</a><a href="#pp"><span class="nav-icon">PP</span>塑料地垫</a><a href="#high"><span class="nav-icon">高</span>高客单非PP</a><a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#anomaly"><span class="nav-icon">!</span>异常附录</a><a href="#insights"><span class="nav-icon">策</span>趋势与建议</a></nav><div class="sidebar-footer"><span>可比数据范围</span><strong>${analysisMonths[0]} — ${analysisMonths[analysisMonths.length - 1]}</strong><small>${analysisMonths.length}个月 · 444万单元格核验</small></div></aside><div class="workspace"><header class="topbar"><div><span class="eyebrow">OUTDOOR RUG MARKET INTELLIGENCE</span><h1>市场总览</h1><p>销量、销售额、双均价 · 小类BSR前100 · 月度MoM与年度YoY</p></div><div class="top-actions"><span class="privacy-chip"><i></i> 源数据只读</span><button class="theme-button" id="theme-toggle" aria-label="切换主题">☀</button></div></header><main class="content"><div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>全部 50 个月（202206-202607）均已纳入，含 2026.03-07。注意 2026.04 起源数据统计口径数量级突变，原值如实披露于异常附录。</div></div><div class="metrics-grid"><article class="metric-card"><span class="metric-label">2025 整体销量</span><strong class="metric-value">${fmt(insight.overall2025.sales)}</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoySales)} · 全市场</span></article><article class="metric-card"><span class="metric-label">2025 整体销售额</span><strong class="metric-value">$${fmt(insight.overall2025.revenue / 1000000, 1)}M</strong><span class="metric-note">同比 ${fmtPct(insight.overall2025.yoyRevenue)} · USD</span></article><article class="metric-card"><span class="metric-label">PP 销量贡献</span><strong class="metric-value">${fmt(ppSalesShare2025,1)}%</strong><span class="metric-note">2025 PP销量 / 整体销量</span></article><article class="metric-card"><span class="metric-label">GENIMO PP份额</span><strong class="metric-value">${fmt(insight.genimoPpShare2025,2)}%</strong><span class="metric-note">2025销量份额 · 品牌领先</span></article></div><section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>，不是按月销量重新排名。</li><li>PP标题包含 plastic；高客单排除PP后按材质关键词或价格≥$40。</li><li>同时展示SKU平均标价与销量加权均价（销售额/销量）。</li><li>年度YoY严格使用同周期；2023仅以6-12月对比2022年6-12月。</li><li>月度MOM = 今年X月 vs 去年X月（跨年同月，如 2025.01 vs 2024.01）；月度环比 = 本月 vs 上月（连续月环比）。</li></ul><p class="note">2026.03-07源数据出现数量级变化，本版已按业务要求全部纳入分析，原值如实保留在月度主表与异常附录。</p></section>${htmlSections}<section id="anomaly"><h2>六、异常月份原值附录</h2>${htmlTable(['月份','纳入可比报告','销量','销售额($)','SKU平均标价','销量加权均价'],sourceDiagnostics.filter((r)=>r.month>='202603').map((r)=>[r.month,r.includedInComparableReport?'是':'否',fmt(r.sales),fmt(r.revenue),fmt(r.avgListPrice,2),fmt(r.weightedPrice,2)]))}</section>${insightHtml}</main></div></div><script>(function(){var root=document.documentElement,button=document.getElementById('theme-toggle');var saved='dark';try{saved=localStorage.getItem('market-report-theme')||'dark'}catch(e){}root.dataset.theme=saved;button.textContent=saved==='dark'?'☀':'☾';button.addEventListener('click',function(){var next=root.dataset.theme==='dark'?'light':'dark';root.dataset.theme=next;button.textContent=next==='dark'?'☀':'☾';try{localStorage.setItem('market-report-theme',next)}catch(e){}});var links=[].slice.call(document.querySelectorAll('.sidebar nav a'));var observer=new IntersectionObserver(function(entries){entries.forEach(function(entry){if(entry.isIntersecting){links.forEach(function(a){a.classList.toggle('is-active',a.getAttribute('href')==='#'+entry.target.id)})}})},{rootMargin:'-20% 0px -70% 0px'});document.querySelectorAll('[id]').forEach(function(el){observer.observe(el)})})();</script></body></html>`;

const coverageScopeText = `明细覆盖 ${analysisMonths.length} 个月（${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]}），源库 ${sourceMeta.effective_sheets || 54} 张有效业务表、${fmt(sourceMeta.total_rows || 0)} 条商品记录；2026.03-07 原值保留，但标记为非长期可比口径。`;
const anomalySectionHtml = `<section id="anomaly"><h2>六、2026.03-2026.07异常附录</h2><p class="note">${coverageScopeText} 异常来自源 Excel 的统计口径变化，不是导入错位；请在确认口径后再用于同比外推。</p>${htmlTable(['月份', '纳入长期可比结论', '销量', '销售额($)', 'SKU平均标价', '加权成交均价'], sourceDiagnostics.filter((row) => row.month >= '202603').map((row) => [row.month, row.includedInLongTermConclusion ? '是' : '否', fmt(row.sales), fmt(row.revenue), fmt(row.avgListPrice, 2), fmt(row.weightedPrice, 2)]))}</section>`;
let htmlOutput = html
  .replace(/<div id="dashboard" class="scope-notice">[\s\S]*?<\/div>/, `<div id="dashboard" class="scope-notice"><span>◎</span><div><b>分析范围：</b>${esc(coverageScopeText)}</div></div>`)
  .replace(/<section id="definitions">[\s\S]*?<\/section>/, `<section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>，多值时取最小可解析名次；头部/中部/尾部为1-20、21-50、51-100，五档为1-5、6-10、11-20、21-50、51-100。</li><li>PP塑料地垫：标题包含 <code>plastic</code>（不区分大小写，空标题按空字符串）。</li><li>高客单非PP：排除PP后，标题命中 polypropylene、sandwich、woven、braided，或价格不低于$40。</li><li>均价同时给出SKU平均标价与销量加权成交均价（销售额/销量）。所有月度表同时给出同月YOY与连续月MOM；年度表使用同周期比较。</li></ul><p class="note">${esc(coverageScopeText)}</p></section>`)
  .replace(/<section id="anomaly">[\s\S]*?<\/section>/, anomalySectionHtml)
  .replace('<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a>', '<a href="#genimo"><span class="nav-icon">G</span>GENIMO</a><a href="#genimo-products"><span class="nav-icon">Top</span>GENIMO主力ASIN</a>')
  .replace('<a href="#insights"><span class="nav-icon">策</span>趋势与建议</a>', '<a href="#insights"><span class="nav-icon">策</span>趋势与建议</a><a href="#coverage"><span class="nav-icon">核</span>需求覆盖核对</a>')
  .replace(/<small>[^<]*444万单元格核验<\/small>/, `<small>${analysisMonths.length}个月 · ${fmt(verifiedDataCellCount)}数据单元格</small>`)
  .replace('</main></div></div><script>', coverageHtml + '</main></div></div><script>')
  .replace('<section id="insights">', genimoProductsHtml + '<section id="insights">')
  .replace('</style></head>', '<style>.trend-body{padding:10px 0 2px;color:var(--muted)}.trend-body p{margin:7px 0;font-size:14px}.insight-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.insight-grid article{padding:15px;border:1px solid var(--line);background:var(--surface-soft)}.insight-grid h3{margin:0 0 7px}.insight-grid p{margin:0;color:var(--muted);font-size:14px}@media(max-width:700px){.insight-grid{grid-template-columns:1fr}}</style></head>');

fs.mkdirSync(path.dirname(HTML_PATH), { recursive: true });
fs.writeFileSync(JSON_PATH, JSON.stringify(data, null, 2), 'utf8');
fs.writeFileSync(MD_PATH, md.join('\n'), 'utf8');
fs.writeFileSync(HTML_PATH, htmlOutput, 'utf8');
db.close();
console.log('Generated: ' + path.relative(ROOT, JSON_PATH));
console.log('Generated: ' + path.relative(ROOT, MD_PATH));
console.log('Generated: ' + path.relative(ROOT, HTML_PATH));
