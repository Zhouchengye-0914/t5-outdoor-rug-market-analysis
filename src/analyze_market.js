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
    row.momSales = previous ? pct(row.sales, previous.sales) : null;
    row.momRevenue = previous ? pct(row.revenue, previous.revenue) : null;
    row.yoySales = lastYear ? pct(row.sales, lastYear.sales) : null;
    row.yoyRevenue = lastYear ? pct(row.revenue, lastYear.revenue) : null;
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
      yoyWeightedPrice: comparable ? pct(comparableCurrent.weightedPrice, prior.weightedPrice) : null,
    };
  });
}

const db = new DatabaseSync(DB_PATH, { readOnly: true });
const catalog = db.prepare("SELECT * FROM sheet_catalog WHERE classification='monthly' ORDER BY sheet_order").all();
const sourceMonths = catalog.map((row) => row.target_table.replace('monthly_', ''));
const analysisMonths = sourceMonths.filter((month) => month <= REPORT_CUTOFF);
const rawByMonth = new Map();

for (const month of sourceMonths) {
  const table = 'monthly_' + month;
  const rows = db.prepare('SELECT row_id, ASIN asin, 品牌 brand, 商品标题 title, 小类BSR bsr, 月销量 sales, 月销售额 revenue, 价格 price FROM ' + table).all();
  rawByMonth.set(month, rows.map((row) => ({ ...row, ...parseBsr(row.bsr) })));
}

const categories = {};
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const monthly = [];
  const bsrSegments = [];
  for (const month of analysisMonths) {
    const rows = rawByMonth.get(month).filter((row) => classify(row, category));
    monthly.push({ month, ...summarize(rows) });
    for (const segment of SEGMENTS) {
      const segmentRows = rows.filter((row) => row.rank >= segment.min && row.rank <= segment.max);
      bsrSegments.push({ month, segment: segment.key, ...summarize(segmentRows) });
    }
  }
  categories[category] = {
    monthly: addTrends(monthly),
    annual: buildAnnual(monthly),
    bsrSegments: addTrendsBySegment(bsrSegments),
  };
}

function addTrendsBySegment(rows) {
  for (const segment of SEGMENTS) {
    addTrends(rows.filter((row) => row.segment === segment.key));
  }
  return rows;
}

const sourceDiagnostics = sourceMonths.map((month) => ({
  month,
  includedInComparableReport: month <= REPORT_CUTOFF,
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
  sourceMonths,
  analysisMonths,
  excludedFromComparableReport: sourceMonths.filter((month) => month > REPORT_CUTOFF),
  definitions: {
    pp: "LOWER(COALESCE(商品标题,'')) contains 'plastic'",
    high: "not PP and (title contains configured material keyword or price >= 40)",
    bsrTop100: 'minimum numeric rank parsed from 小类BSR; bands 1-5, 6-10, 11-20, 21-50, 51-100',
    avgListPrice: 'simple average of SKU list prices',
    weightedPrice: '月销售额 / 月销量',
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

const pp2025Rows = analysisMonths.filter((month) => month.startsWith('2025'))
  .flatMap((month) => rawByMonth.get(month).filter((row) => classify(row, 'pp')));
const genimoPp2025Rows = pp2025Rows.filter((row) => classify(row, 'genimo'));
const pp2025Sales = summarize(pp2025Rows).sales;
data.insights = {
  overall2025: annualRow('overall', '2025'),
  pp2025: annualRow('pp', '2025'),
  high2025: annualRow('high', '2025'),
  genimo2025: annualRow('genimo', '2025'),
  ppPeak2025: peakMonth('pp', '2025'),
  genimoPpShare2025: pp2025Sales ? summarize(genimoPp2025Rows).sales / pp2025Sales * 100 : null,
};

function mdMonthly(rows) {
  const out = ['| 月份 | SKU数 | 销量 | 销售额($) | SKU平均标价($) | 销量加权均价($) | MoM销量 | MoM销售额 | YoY销量 | YoY销售额 |',
    '|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.momSales)} | ${fmtPct(r.momRevenue)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} |`);
  return out.join('\n');
}

function mdAnnual(rows) {
  const out = ['| 年份/数据周期 | YoY比较周期 | 销量 | 销售额($) | SKU平均标价($) | 销量加权均价($) | 同周期YoY销量 | 同周期YoY销售额 |',
    '|---|---|---:|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.year} (${r.period}) | ${r.comparison || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} | ${fmtPct(r.yoySales)} | ${fmtPct(r.yoyRevenue)} |`);
  return out.join('\n');
}

function mdSegments(rows) {
  const out = ['| 月份 | BSR分层 | SKU数 | 销量 | 销售额($) | MoM销量 | YoY销量 |',
    '|---|---|---:|---:|---:|---:|---:|'];
  for (const r of rows) out.push(`| ${r.month} | ${r.segment} | ${fmt(r.skuCount)} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmtPct(r.momSales)} | ${fmtPct(r.yoySales)} |`);
  return out.join('\n');
}

const labels = { overall: '整体市场', pp: 'PP塑料地垫', high: '非PP高客单产品', genimo: 'GENIMO品牌' };
const md = ['# 户外地垫市场分析报告（优化版）', '',
  `> 可比分析范围：${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]}，共 ${analysisMonths.length} 个月。源数据另含 ${data.excludedFromComparableReport.join('、')}，因统计口径数量级突变，仅在异常附录披露原值。`, '',
  '## 一、口径说明', '',
  '- 小类前100严格依据源字段 `小类BSR`，多值BSR取最小可解析名次并保留多值标记。',
  '- PP：标题包含 `plastic`（不区分大小写，NULL按空字符串处理）。',
  '- 高客单：排除PP后，标题命中材质关键词或价格不低于$40。',
  '- 同时提供SKU平均标价和销量加权均价（销售额/销量）。',
  '- 年度YoY使用同周期比较；2023对2022仅比较6-12月，避免12个月对7个月。', ''];

let sectionNo = 2;
for (const category of ['overall', 'pp', 'high', 'genimo']) {
  const c = categories[category];
  md.push(`## ${sectionNo++}、${labels[category]}`, '', '### 月度指标、MoM与YoY', '', mdMonthly(c.monthly), '',
    '### 年度/同周期汇总', '', mdAnnual(c.annual), '', '### 小类BSR前100分层', '', mdSegments(c.bsrSegments), '');
}

md.push('## 六、2026.03-2026.07异常附录', '', '| 月份 | 是否纳入可比报告 | 销量 | 销售额($) | SKU平均标价($) | 销量加权均价($) |',
  '|---|---|---:|---:|---:|---:|');
for (const r of sourceDiagnostics.filter((r) => r.month >= '202603')) {
  md.push(`| ${r.month} | ${r.includedInComparableReport ? '是' : '否'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${fmt(r.avgListPrice, 2)} | ${fmt(r.weightedPrice, 2)} |`);
}
md.push('', '## 七、GENIMO累计Top产品', '', '| 排名 | ASIN | 累计销量 | 累计销售额($) | 月数 | 最新价($) | 标题 |',
  '|---:|---|---:|---:|---:|---:|---|');
data.genimoTopProducts.forEach((r, i) => md.push(`| ${i + 1} | ${r.asin || '-'} | ${fmt(r.sales)} | ${fmt(r.revenue)} | ${r.months} | ${fmt(r.latestPrice, 2)} | ${String(r.title || '').replace(/\|/g, '\\|')} |`));

const insight = data.insights;
md.push('', '## 八、趋势结论与GENIMO建议', '',
  `- **整体市场**：2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}，销量加权均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}。销量增长快于销售额，市场存在价格下压。`,
  `- **PP市场**：2025销量同比 ${fmtPct(insight.pp2025.yoySales)}，销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}；销量峰值为 ${insight.ppPeak2025.month} 的 ${fmt(insight.ppPeak2025.sales)} 件。`,
  `- **高客单非PP**：2025销量同比 ${fmtPct(insight.high2025.yoySales)}，销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}，表现弱于PP。`,
  `- **GENIMO**：2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}。`, '',
  '### 建议', '',
  '1. 按尺寸、价格带和小类BSR管理PP产品，优先保障BSR 1-20核心SKU库存与广告。',
  '2. 同时考核销量、销售额和销量加权均价，避免只追求件数导致价格与利润空间持续受压。',
  '3. 将高客单拆分为“材质关键词命中”和“仅价格命中”两组，小规模验证非PP第二增长曲线。',
  '4. 根据月度MoM/YoY在3-5月旺季前置补货和新品测试。',
  '5. 在确认2026.03-07统计口径前，不使用异常月份制定年度预算或宣称市场暴增。');

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function htmlTable(headers, rows) {
  return '<div class="table-wrap"><table><thead><tr>' + headers.map((h) => '<th>' + esc(h) + '</th>').join('')
    + '</tr></thead><tbody>' + rows.map((row) => '<tr>' + row.map((v) => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('')
    + '</tbody></table></div>';
}

function monthlyHtml(rows) {
  return htmlTable(['月份', 'SKU数', '销量', '销售额($)', 'SKU平均标价', '销量加权均价', 'MoM销量', 'MoM销售额', 'YoY销量', 'YoY销售额'],
    rows.map((r) => [r.month, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.yoySales), fmtPct(r.yoyRevenue)]));
}

function annualHtml(rows) {
  return htmlTable(['年份/数据周期', 'YoY比较周期', '销量', '销售额($)', 'SKU平均标价', '销量加权均价', '同周期YoY销量', '同周期YoY销售额'],
    rows.map((r) => [`${r.year} (${r.period})`, r.comparison || '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2), fmtPct(r.yoySales), fmtPct(r.yoyRevenue)]));
}

function segmentHtml(rows) {
  return htmlTable(['月份', '小类BSR分层', 'SKU数', '销量', '销售额($)', 'MoM销量', 'YoY销量'],
    rows.map((r) => [r.month, r.segment, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmtPct(r.momSales), fmtPct(r.yoySales)]));
}

const htmlSections = ['overall', 'pp', 'high', 'genimo'].map((category, index) => {
  const c = categories[category];
  return `<section id="${category}"><h2>${index + 2}、${labels[category]}</h2><details open><summary>月度指标、MoM与YoY</summary>${monthlyHtml(c.monthly)}</details><details><summary>年度/同周期汇总</summary>${annualHtml(c.annual)}</details><details><summary>小类BSR前100五档分层</summary>${segmentHtml(c.bsrSegments)}</details></section>`;
}).join('\n');

const insightHtml = `<section id="insights"><h2>七、趋势结论与GENIMO建议</h2><ul><li><b>整体：</b>2025销量同比 ${fmtPct(insight.overall2025.yoySales)}，销售额同比 ${fmtPct(insight.overall2025.yoyRevenue)}，销量加权均价同比 ${fmtPct(insight.overall2025.yoyWeightedPrice)}。</li><li><b>PP：</b>2025销量同比 ${fmtPct(insight.pp2025.yoySales)}，销售额同比 ${fmtPct(insight.pp2025.yoyRevenue)}。</li><li><b>高客单非PP：</b>2025销量同比 ${fmtPct(insight.high2025.yoySales)}，销售额同比 ${fmtPct(insight.high2025.yoyRevenue)}。</li><li><b>GENIMO：</b>2025年PP销量份额 ${fmt(insight.genimoPpShare2025, 2)}%，品牌销量同比 ${fmtPct(insight.genimo2025.yoySales)}。</li></ul><h3>行动建议</h3><ol><li>优先保障小类BSR 1-20核心PP SKU的库存与广告。</li><li>同时考核销量、销售额与销量加权均价，控制价格下压风险。</li><li>把高客单拆成材质命中与仅价格命中两组，小规模验证非PP第二曲线。</li><li>依据3-5月旺季规律前置补货与新品测试。</li><li>2026.03-07口径确认前不用于年度预算和增长宣称。</li></ol></section>`;

const html = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场分析报告（优化版）</title><style>
:root{--ink:#172033;--muted:#667085;--brand:#155eef;--line:#d0d5dd;--bg:#f6f8fc;--card:#fff}*{box-sizing:border-box}body{margin:0;font-family:Inter,"Microsoft YaHei",sans-serif;color:var(--ink);background:var(--bg);line-height:1.55}header{background:linear-gradient(135deg,#102a56,#155eef);color:#fff;padding:40px max(24px,calc((100vw - 1320px)/2))}header h1{margin:0 0 10px;font-size:32px}header p{margin:0;opacity:.9}nav{position:sticky;top:0;z-index:2;background:#fff;border-bottom:1px solid var(--line);padding:10px max(16px,calc((100vw - 1320px)/2));display:flex;gap:14px;flex-wrap:wrap}nav a{color:var(--brand);text-decoration:none;font-weight:600}main{max-width:1320px;margin:24px auto;padding:0 20px}section{background:var(--card);border:1px solid #e4e7ec;border-radius:14px;padding:24px;margin-bottom:20px;box-shadow:0 4px 18px rgba(16,24,40,.05)}h2{margin-top:0}details{border-top:1px solid #eaecf0;padding:14px 0}summary{cursor:pointer;font-weight:700;color:#1849a9}.table-wrap{overflow:auto;max-height:620px;margin-top:12px}table{border-collapse:collapse;width:100%;min-width:980px;font-size:13px}th,td{padding:8px 10px;border-bottom:1px solid #eaecf0;text-align:right;white-space:nowrap}th{position:sticky;top:0;background:#eff4ff;color:#1849a9}th:first-child,td:first-child{text-align:left}.note{padding:14px 16px;background:#fffaeb;border-left:4px solid #f79009;border-radius:8px}code{background:#f2f4f7;padding:2px 5px;border-radius:4px}@media(max-width:700px){header h1{font-size:24px}main{padding:0 10px}section{padding:16px}}
</style></head><body><header><h1>户外地垫市场分析报告（优化版）</h1><p>可比范围 ${analysisMonths[0]}-${analysisMonths[analysisMonths.length - 1]} · 小类BSR前100 · 月/年 MoM & YoY · 双均价口径</p></header><nav><a href="#definitions">口径</a><a href="#overall">整体</a><a href="#pp">PP</a><a href="#high">高客单</a><a href="#genimo">GENIMO</a><a href="#anomaly">异常附录</a><a href="#insights">结论建议</a></nav><main><section id="definitions"><h2>一、口径说明</h2><ul><li>小类前100依据源字段 <code>小类BSR</code>，不是按月销量重新排名。</li><li>PP标题包含 plastic；高客单排除PP后按材质关键词或价格≥$40。</li><li>同时展示SKU平均标价与销量加权均价（销售额/销量）。</li><li>年度YoY严格使用同周期；2023仅以6-12月对比2022年6-12月。</li></ul><p class="note">2026.03-07源数据出现数量级变化，本版保留原值附录但不并入长期可比结论，等待业务确认统计口径。</p></section>${htmlSections}<section id="anomaly"><h2>六、异常月份原值附录</h2>${htmlTable(['月份','纳入可比报告','销量','销售额($)','SKU平均标价','销量加权均价'],sourceDiagnostics.filter((r)=>r.month>='202603').map((r)=>[r.month,r.includedInComparableReport?'是':'否',fmt(r.sales),fmt(r.revenue),fmt(r.avgListPrice,2),fmt(r.weightedPrice,2)]))}</section>${insightHtml}</main></body></html>`;

fs.mkdirSync(path.dirname(HTML_PATH), { recursive: true });
fs.writeFileSync(JSON_PATH, JSON.stringify(data, null, 2), 'utf8');
fs.writeFileSync(MD_PATH, md.join('\n'), 'utf8');
fs.writeFileSync(HTML_PATH, html, 'utf8');
db.close();
console.log('Generated: ' + path.relative(ROOT, JSON_PATH));
console.log('Generated: ' + path.relative(ROOT, MD_PATH));
console.log('Generated: ' + path.relative(ROOT, HTML_PATH));
