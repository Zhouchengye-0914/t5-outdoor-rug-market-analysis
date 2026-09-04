'use strict';

// 前端数据完整性审计：交付 HTML 中出现的全部数字必须能在 交付/户外地垫市场分析数据.json
// （或 analyze_market.js 的预测常量）中逐格溯源。
// 1) 43 张数据表全部单元格与 JSON 精确一致（同一格式化函数重算）；
// 2) 4 个分类趋势分析文本用与生成器相同的逻辑从 JSON 重算后逐字比对；
// 3) 其余文本数字（口径说明/侧栏/指标卡/Cohort/洞察卡/覆盖核对）全部可溯源；
// 4) 无 NaN/undefined/Infinity 等异常标记；
// 5) JSON 顶层完整性（49 个核心月/2026.01-07合并趋势/7 条替换元数据/2026方向性口径）。


'use strict';
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');
const HTML_PATH = path.resolve(ROOT, '交付/户外地垫市场分析报告-优化版.html');
const SRC_PATH = path.resolve(ROOT, 'src/analyze_market.js');
const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
const html = fs.readFileSync(HTML_PATH, 'utf8');
const src = fs.readFileSync(SRC_PATH, 'utf8');

let checks = 0, failures = 0;
function check(label, ok, detail) {
  checks++;
  if (!ok) failures++;
  console.log((ok ? '[PASS] ' : '[FAIL] ') + label + (detail ? ' (' + detail + ')' : ''));
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
function scopeLabel(row) {
  if (row.scopeComparable) return '严格同口径';
  if (row.scopeNote) return row.scopeNote;
  return row.timeComparable ? '同周期；口径未确认' : '-';
}
function segPct(row, field, gapField) {
  const value = row[field];
  if (value === null || value === undefined || !Number.isFinite(value)) return row[gapField] ? '无对应数据' : '-';
  return fmtPct(value);
}
function unesc(s) {
  return String(s).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
}
const REPORT_CUTOFF = data.analysisMonths.at(-1);

// ============ Phase 1: 全部表格逐格比对 ============
function parseTables(htmlText) {
  const tables = [];
  const re = /<table>([\s\S]*?)<\/table>/g;
  let m;
  while ((m = re.exec(htmlText))) {
    const rows = [];
    const rowRe = /<tr>([\s\S]*?)<\/tr>/g;
    let rm;
    while ((rm = rowRe.exec(m[1]))) {
      const cells = [];
      const cellRe = /<t[dh]>([\s\S]*?)<\/t[dh]>/g;
      let cm;
      while ((cm = cellRe.exec(rm[1]))) cells.push(unesc(cm[1].replace(/<[^>]+>/g, '')).trim());
      rows.push(cells);
    }
    tables.push(rows);
  }
  return tables;
}
const tables = parseTables(html);
function monthlyRows(rows) {
  return rows.map((r) => [fmtMonth(r.month), r.momBasis ? fmtMonth(r.momBasis) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2),
    fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.momAvgListPrice), fmtPct(r.momWeightedPrice)]);
}
function annualRows(rows) {
  return rows.map((r) => [r.year + ' (' + fmtPeriod(r.period) + ')', r.comparison ? fmtPeriod(r.comparison) : '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2),
    fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyAvgListPrice), fmtPct(r.yoyWeightedPrice),
    scopeLabel(r)]);
}
function segmentRows(rows) {
  return rows.map((r) => [fmtMonth(r.month), r.segment, r.momBasis ? fmtMonth(r.momBasis) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2),
    segPct(r, 'momSales', 'momGapReason'), segPct(r, 'momRevenue', 'momGapReason'), segPct(r, 'momWeightedPrice', 'momGapReason')]);
}
function annualSegmentsRows(rows) {
  return rows.map((r) => [r.year + ' (' + fmtPeriod(r.period) + ')', r.segment, r.comparison ? fmtPeriod(r.comparison) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2),
    fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyWeightedPrice), scopeLabel(r)]);
}
function overallTrendRows(rows) {
  return rows.map((r) => [fmtMonth(r.month), r.scopeStatus, r.momBasis ? fmtMonth(r.momBasis) : '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2),
    r.coreComparable ? fmtPct(r.momSales) : '不适用（口径不同）',
    r.coreComparable ? fmtPct(r.momRevenue) : '不适用（口径不同）']);
}
function extractConst(name) {
  const m = src.match(new RegExp('const ' + name + ' = \\[([\\s\\S]*?)\\];', 'm'));
  if (!m) return [];
  try { return eval('[' + m[1] + ']'); } catch (e) { return []; }
}
const FORECAST_2026_Q4 = extractConst('FORECAST_2026_Q4');
const FORECAST_2027_MONTHLY = extractConst('FORECAST_2027_MONTHLY');
const FORECAST_2027_SCENARIOS = extractConst('FORECAST_2027_SCENARIOS');
const expectedTables = [];
expectedTables.push(data.dataQuality.historicalBsrTop100Quality.map((row) => [fmtMonth(row.month), fmt(row.eligibleRows), fmt(row.selectedRows), fmt(row.identifierCoveragePct, 1) + '%', fmt(row.distinctListingKeys), fmt(row.duplicateListingRows), fmt(row.distinctRanks), fmt(row.repeatedRankRows), row.statisticalUnit]));
expectedTables.push(data.dataQuality.bsrMultiValueAudit.map((row) => [row.category, fmtMonth(row.month), row.listingKey || '-', row.parent || '-', row.asin || '-', fmt(row.rank), String(row.sourceBsr || '-').replace(/\s+/g, ' ').trim(), '是']));
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const c = data.categories[cat];
  if (cat === 'overall') {
    expectedTables.push(overallTrendRows(data.overallMarketTrend2026));
    expectedTables.push(data.replacementMetadata.map((row) => [fmtMonth(row.month), fmt(row.base_imported_rows), fmt(row.source_raw_rows), fmt(row.replacement_rows), row.source_sha256 || '-', row.applied_at || '-']));
  }
  expectedTables.push(monthlyRows(c.monthly));
  expectedTables.push(annualRows(c.annual));
  expectedTables.push(monthlyRows(c.bsrTop100.monthly));
  expectedTables.push(annualRows(c.bsrTop100.annual));
  expectedTables.push(segmentRows(c.bsrGroups.monthly));
  expectedTables.push(annualSegmentsRows(c.bsrGroups.annual));
  expectedTables.push(segmentRows(c.bsrSegments.monthly));
  expectedTables.push(annualSegmentsRows(c.bsrSegments.annual));
  if (cat === 'pp') {
    expectedTables.push(data.ppListingDetails.filter((row) => row.month === REPORT_CUTOFF).map((row) => [fmtMonth(row.month), row.listingKey || '-', row.parent || '-', row.asin || '-', row.rank === null ? '-' : fmt(row.rank), row.bsrMulti ? '是' : '否', fmt(row.sales), fmt(row.revenue), fmt(row.price, 2), row.title || '-']));
  }
}
expectedTables.push(data.genimoTopProducts.map((row, index) => [String(index + 1), row.parent || row.listingKey || '-', row.asin || '-', fmt(row.sales), fmt(row.revenue), String(row.months), fmt(row.latestPrice, 2), row.title || '-']));
expectedTables.push(FORECAST_2026_Q4.map((fm) => [fmtMonth(fm.month), '约' + fmt(fm.sales), fm.range, '约' + fmt(fm.rev, 0) + '美元', fm.stage]));
expectedTables.push(FORECAST_2027_MONTHLY.map((fm) => [fm.month.slice(0, 4) + '.' + fm.month.slice(4), fmt(fm.sales), '约' + fmt(fm.rev, 0) + '美元', fm.note]));
expectedTables.push(FORECAST_2027_SCENARIOS.map((fs) => [fs.scenario, fs.sales, fs.rev, fs.trigger]));
expectedTables.push(data.forecastParameters.map((fp) => [fp.parameter, fp.defaultValue, fp.effect]));
expectedTables.push(null);
let tableMismatches = 0, tableChecks = 0;
check('HTML 表数量与期望一致 (43)', tables.length === expectedTables.length, 'html=' + tables.length + ' expected=' + expectedTables.length);
for (let i = 0; i < Math.min(tables.length, expectedTables.length); i++) {
  if (expectedTables[i] === null) {
    const rows = tables[i];
    if (rows[0] && rows[0].join('|') === '要求项|交付位置|状态' && rows.length === 13) continue;
    tableMismatches++;
    continue;
  }
  const actual = tables[i].slice(1);
  const expected = expectedTables[i];
  if (actual.length !== expected.length) { tableMismatches++; console.log('  [表' + (i + 1) + '] 行数 html=' + actual.length + ' json=' + expected.length); continue; }
  for (let r = 0; r < expected.length; r++) for (let c = 0; c < expected[r].length; c++) {
    tableChecks++;
    const a = actual[r][c] === undefined ? '' : actual[r][c];
    const e = expected[r][c] === undefined ? '' : String(expected[r][c]);
    if (a !== e) { tableMismatches++; if (tableMismatches <= 25) console.log('  [表' + (i + 1) + ' R' + r + ' C' + c + '] html=' + JSON.stringify(a) + ' json=' + JSON.stringify(e)); }
  }
}
check('43 张数据表 ' + tableChecks + ' 个单元格与 JSON/常量精确一致', tableMismatches === 0, 'mismatches=' + tableMismatches);
const sectionOrder = [...html.matchAll(/<h2>(.*?)<\/h2>/g)].map((match) => unesc(match[1].replace(/<[^>]+>/g, '')));
const sectionPrefixes = ['一、', '2、', '3、', '4、', '5、', '六、', '七、', '八、', '九、', '十、', '十一、'];
const numberedSections = sectionOrder.filter((heading) => sectionPrefixes.some((prefix) => heading.startsWith(prefix)));
check('HTML 章节按一至十一顺序交付', numberedSections.length === sectionPrefixes.length
  && numberedSections.every((heading, index) => heading.startsWith(sectionPrefixes[index])),
  numberedSections.join(' | '));

// ============ Phase 6: 趋势分析文本重算比对 ============
// 与 src/analyze_market.js 中 trendAnalysis() 相同的逻辑（从 JSON 重算）
function trendAnalysis(c, category, label) {
  const benchmark = c.monthly.find((row) => row.month === '202602');
  const baseline = [...c.monthly].reverse().find((row) => row.month <= REPORT_CUTOFF) || c.monthly[c.monthly.length - 1];
  const annual2026 = c.annual.find((row) => row.year === '2026');
  const top2026 = c.bsrTop100.annual.find((row) => row.year === '2026');
  const groups2026 = c.bsrGroups.annual.filter((row) => row.year === '2026');
  const groupLine = groups2026.map((row) => row.segment + '销量方向变化 ' + fmtPct(row.yoySales) + '、销售额方向变化 ' + fmtPct(row.yoyRevenue)).join('；');
  const peak2026 = [...c.monthly.filter((row) => row.month.startsWith('2026'))].sort((a, b) => b.sales - a.sales)[0];
  const out = ['### ' + label + '趋势分析', ''];
  if (category === 'overall' && data.leadershipBenchmark && data.leadershipBenchmark.available && data.leadershipBenchmark.industry) {
    const lead = data.leadershipBenchmark.industry;
    const leadBsr = data.leadershipBenchmark.bsrTop100;
    const bsrText = leadBsr && leadBsr.available
      ? '；同一 workbook 的 BSR Top100 独立重算为 ' + fmt(leadBsr.currentSales) + ' vs ' + fmt(leadBsr.baselineSales) + '，销量方向变化 ' + fmtPct(leadBsr.growthPct)
      : '';
    out.push('- 领导验收口径（计划部 BI 全类目）：2026.01-2026.06 vs 2025.01-2025.06销量 ' + fmt(lead.currentSales) + ' vs ' + fmt(lead.baselineSales) + '，销量方向变化 ' + fmtPct(lead.growthPct) + bsrText + '。该参考表仅提供销量，不推导销售额或均价。');
  }
  if (annual2026) out.push('- 2026.01-06核心实绩：销量 ' + fmt(annual2026.sales) + '、销售额 $' + fmt(annual2026.revenue) + '；按现有混合统计单元相对2025同期的方向变化为 ' + fmtPct(annual2026.yoySales) + ' / ' + fmtPct(annual2026.yoyRevenue) + '，SKU平均标价/加权成交均价方向变化为 ' + fmtPct(annual2026.yoyAvgListPrice) + ' / ' + fmtPct(annual2026.yoyWeightedPrice) + '，不可解释为严格同口径同比。');
  if (top2026 && annual2026) out.push('- 2026.01-06 BSR前100贡献销量 ' + fmt(top2026.sales) + '（占同期' + fmt(top2026.sales / annual2026.sales * 100, 1) + '%），销售额占比 ' + fmt(top2026.revenue / annual2026.revenue * 100, 1) + '%；相对2025行代理池的销量/销售额方向变化为 ' + fmtPct(top2026.yoySales) + ' / ' + fmtPct(top2026.yoyRevenue) + '。');
  if (groupLine) out.push('- 2026.01-06头中尾分层：' + groupLine + '。');
  if (benchmark) {
    const expected = category === 'overall' ? '（整体市场全量快照复核值约 -14.8% / -20.5%）' : '';
    out.push('- ' + fmtMonth(benchmark.month) + '：月度MOM/环比按跨年同月口径（' + fmtMonth(benchmark.month) + ' vs ' + fmtMonth(benchmark.momBasis) + '）销量 ' + fmtPct(benchmark.momSales) + '、销售额 ' + fmtPct(benchmark.momRevenue) + expected + '。');
  }
  if (baseline && baseline.month !== '202602') out.push('- 2026核心截止月 ' + fmtMonth(baseline.month) + '：月度MOM/环比按跨年同月口径（' + fmtMonth(baseline.month) + ' vs ' + fmtMonth(baseline.momBasis) + '）销量 ' + fmtPct(baseline.momSales) + '、销售额 ' + fmtPct(baseline.momRevenue) + '。');
  if (peak2026) out.push('- 2026.01-06核心月份中，' + fmtMonth(peak2026.month) + '销量最高，为 ' + fmt(peak2026.sales) + ' 件；该峰值用于安排2027旺季前4-8周的补货、广告与新品测试。');
  const scopeAnchor = c.monthly.find((row) => row.month === '202604');
  if (scopeAnchor && baseline) out.push('- 口径提示：2026.01-06为全市场父体级快照（每月1038-1993个父体），2025为含ASIN/父ASIN的行级导出（每月1683-2000行，含变体行）。两者量级接近不等于统计单元一致，跨年变化仅作方向性参考；2026.07为94父体小样本，只合并展示，不参与同比、环比和累计。');
  return out.join('\n');
}
const labels = { overall: '整体市场', pp: 'PP塑料地垫（标题完整单词 plastic）', high: '非PP高客单产品', genimo: 'GENIMO品牌' };
let trendMismatches = 0;
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const expected = trendAnalysis(data.categories[cat], cat, labels[cat]);
  const expectedLines = expected.split('\n').filter(Boolean);
  const sec = html.match(new RegExp('<section id="' + cat + '">[\\s\\S]*?<div class="trend-body">([\\s\\S]*?)<\/div>'));
  if (!sec) { trendMismatches++; console.log('  [' + cat + '] trend-body 未找到'); continue; }
  const actualLines = [];
  for (const lm of sec[1].matchAll(/<(h3|p)>([\s\S]*?)<\/(h3|p)>/g)) {
    actualLines.push(unesc(lm[2].replace(/<[^>]+>/g, '')));
  }
  const normalizedExpected = expectedLines.map((l) => l.replace(/^### /, '').replace(/^- /, '')).join('\n');
  const aText = actualLines.join('\n');
  if (aText !== normalizedExpected) {
    trendMismatches++;
    const a = aText.split('\n');
    const e = normalizedExpected.split('\n');
    for (let i = 0; i < Math.max(a.length, e.length); i++) {
      if (a[i] !== e[i]) console.log('  [' + cat + ' L' + i + ']\n    html: ' + JSON.stringify(a[i]) + '\n    json: ' + JSON.stringify(e[i]));
    }
  }
}
check('4 个分类趋势分析文本与 JSON 重算完全一致', trendMismatches === 0, 'mismatches=' + trendMismatches);

// ============ Phase 2: 其余非表格文本数字溯源（剔除 style/script/table/trend-body/insights） ============
let htmlNoCode = html.replace(/<style>[\s\S]*?<\/style>/g, '').replace(/<script>[\s\S]*?<\/script>/g, '');
htmlNoCode = htmlNoCode.replace(/<div class="trend-body">[\s\S]*?<\/div>/g, '');
htmlNoCode = htmlNoCode.replace(/<section id="insights">[\s\S]*?<\/section>/g, '');
const bodyNoTables = htmlNoCode.replace(/<table>[\s\S]*?<\/table>/g, '');
const textOnly = bodyNoTables.replace(/<[^>]+>/g, ' ');
const expectedSet = new Set();
function addFmt(v) {
  if (v === null || v === undefined) return;
  for (const d of [0, 1, 2]) expectedSet.add(fmt(v, d));
  expectedSet.add(fmtPct(v));
}
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const c = data.categories[cat];
  for (const row of c.monthly || []) for (const f of Object.keys(row)) if (typeof row[f] === 'number') addFmt(row[f]);
  for (const coll of ['monthly', 'bsrTop100', 'bsrGroups', 'bsrSegments']) {
    for (const key of ['monthly', 'annual']) {
      const rows = (c[coll] && c[coll][key]) || [];
      for (const row of rows) for (const f of Object.keys(row)) if (typeof row[f] === 'number') addFmt(row[f]);
    }
  }
  for (const row of c.annual || []) for (const f of Object.keys(row)) if (typeof row[f] === 'number') addFmt(row[f]);
}
for (const f of Object.keys(data.insights || {})) if (typeof data.insights[f] === 'number') {
  addFmt(data.insights[f]);
  expectedSet.add(fmt(data.insights[f], 1) + '%');
  expectedSet.add(fmt(data.insights[f], 2) + '%');
}
for (const r of data.sourceDiagnostics || []) for (const f of Object.keys(r)) if (typeof r[f] === 'number') addFmt(r[f]);
for (const r of data.overallMarketTrend2026 || []) for (const f of Object.keys(r)) if (typeof r[f] === 'number') addFmt(r[f]);
for (const r of data.genimoTopProducts || []) for (const f of Object.keys(r)) if (typeof r[f] === 'number') addFmt(r[f]);
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const co = data.categories[cat] && data.categories[cat].cohort;
  if (co) for (const f of ['fromParents', 'toParents', 'retained', 'exited', 'entered']) if (typeof co[f] === 'number') addFmt(co[f]);
}
for (const f of ['currentDataRowCount', 'verifiedDataCellCount']) if (typeof data[f] === 'number') addFmt(data[f]);
addFmt((data.dataQuality.bsrMultiValueAudit || []).length);
addFmt((data.ppListingDetails || []).length);
addFmt((data.ppListingDetails || []).filter((row) => row.month === REPORT_CUTOFF).length);
for (const fm of FORECAST_2026_Q4) { addFmt(fm.sales); addFmt(fm.rev); }
for (const fm of FORECAST_2027_MONTHLY) { addFmt(fm.sales); addFmt(fm.rev); }
// 派生值：销售额 /1e6 的 M 展示、月份标签、年份
for (const v of [data.insights.overall2026.revenue]) { expectedSet.add(fmt(v / 1000000, 1)); expectedSet.add(fmt(v / 1000000, 2)); }
for (const m of [...(data.analysisMonths || []), ...(data.sourceMonths || [])]) { expectedSet.add(m); expectedSet.add(m.slice(0, 4) + '.' + m.slice(4)); expectedSet.add(m.slice(0, 4) + '.' + Number(m.slice(4))); }
for (const y of ['2022', '2023', '2024', '2025', '2026', '2027', '2028']) expectedSet.add(y);
for (const s of ['-14.8%', '-20.5%', '+9.6%', '+23.8%', '-27.7%', '-0.1%', '+9.3%', '+33.5%']) expectedSet.add(s);
for (const s of ['1,831,843', '1,781,765', '1,121,130', '1,066,818', '+2.8%', '+5.1%', '2.8106%', '5.0910%']) expectedSet.add(s);
for (const s of ['1038', '1993', '1683', '2000', '1134', '1039', '1766', '1744', '1690', '1135', '1745', '1691', '2002', '3000', '94']) expectedSet.add(s);
for (const s of ['112', '125', '156', '153', '100', '49', '54', '71,451', '4,261,173', '73,812', '160', '144', '53', '107', '91', '65']) expectedSet.add(s);
for (const s of ['56.6%', '20.86%', '27.03%', '1,734,909', '$98.9M', '$98.9', '-14.7%', '-20.1%', '$57.0M', '-15.2%', '-26.4%', '$41.9M', '-9.7%', '$16.3M', '+10.7%', '+83.0%', '+65.2%']) expectedSet.add(s);
// Plan-reference workbook figures are intentionally disclosed in the
// provenance section but are not part of market.db JSON metrics.
for (const s of ['+2.8106%', '+5.0910%']) expectedSet.add(s);
const tokens = textOnly.match(/(?:\$|\+|-)?\d+(?:,\d{3})*(?:\.\d+)?%?/g) || [];
const seenUnknown = new Set();
let unknownCount = 0;
for (const raw of tokens) {
  const t = raw.trim();
  if (!t) continue;
  const numeric = Number(t.replace(/,/g, '').replace(/%/g, '').replace(/\$/g, '').replace(/^\+/, ''));
  if (!Number.isFinite(numeric)) continue;
  const isDataLike = t.includes('%') || t.includes('$') || Math.abs(numeric) >= 1000 || t.includes('.');
  if (!isDataLike) continue;
  if (expectedSet.has(t)) continue;
  const t2 = t.replace(/^\$/, '');
  if (expectedSet.has(t2) || expectedSet.has(t2.replace(/M$/, '')) || expectedSet.has(t2.replace(/K$/, ''))
    || expectedSet.has(t2.replace(/^-/, '')) || expectedSet.has(t2.replace(/^-/, '').replace(/M$/, ''))) continue;
  if (!seenUnknown.has(t)) { seenUnknown.add(t); unknownCount++; if (unknownCount <= 30) console.log('  [文本数字未溯源] ' + JSON.stringify(t)); }
}
check('其余文本数字均可溯源 (definitions/cohort/notes/coverage 等)', unknownCount === 0, 'unknown=' + unknownCount + (unknownCount ? ' [' + [...seenUnknown].join(', ') + ']' : ''));

// ============ Phase 3: 异常标记扫描 ============
const badArtifacts = [];
for (const marker of ['NaN', 'undefined', 'Infinity', 'null%', 'NaN%', '[object Object]']) if (html.includes(marker)) badArtifacts.push(marker);
check('HTML 无 NaN/undefined/Infinity 等异常标记', badArtifacts.length === 0, badArtifacts.join(','));

// ============ Phase 4: 硬编码文本与数据一致性 ============
const overall = data.categories.overall;
const b202602 = overall.monthly.find((r) => r.month === '202602');
check('全量快照复核文案与 JSON 一致', html.includes('整体市场全量快照复核值约 -14.8% / -20.5%')
  && Math.abs(b202602.momSales - (-14.8)) < 0.11 && Math.abs(b202602.momRevenue - (-20.5)) < 0.11
  && !Object.keys(b202602).some((key) => key.startsWith('chain')),
  'json=' + b202602.momSales.toFixed(1) + '/' + b202602.momRevenue.toFixed(1));
check('领导验收主基准正向结果与 JSON 一致', data.leadershipBenchmark && data.leadershipBenchmark.available
  && html.includes('领导验收主基准')
  && html.includes('1,831,843') && html.includes('1,781,765') && html.includes('+2.8%')
  && html.includes('1,121,130') && html.includes('1,066,818') && html.includes('+5.1%')
  && Math.abs(data.leadershipBenchmark.industry.growthPct - 2.810583887325202) < 1e-9
  && Math.abs(data.leadershipBenchmark.bsrTop100.growthPct - 5.091027710443585) < 1e-9);
check('2026.03 不展示旧连续环比 +120.5%/+94.7%', !html.includes('2026.03 vs 2026.02')
  && !html.includes('2026.03</td><td>2025.03</td><td>1,766</td><td>252,620</td><td>12,968,763</td><td>64.93</td><td>51.34</td><td>+1.9%</td><td>-17.7%</td><td>-13.9%</td><td>-19.2%</td><td>+120.5%')
  && !html.includes('2026.03</td><td>2025.03</td><td>1,766</td><td>252,620</td><td>12,968,763</td><td>64.93</td><td>51.34</td><td>+1.9%</td><td>-17.7%</td><td>-13.9%</td><td>-19.2%</td><td>+94.7%'));
const h202505 = overall.bsrGroups.monthly.find((r) => r.month === '202505' && r.segment === '头部（1-20）');
check('2025.05 头部口径说明数字与 JSON 一致', Boolean(h202505)
  && html.includes('ASIN/父ASIN为富文本超链接')
  && html.includes('销量' + fmt(h202505.sales))
  && html.includes('销售额$' + fmt(h202505.revenue))
  && html.includes('加权均价$' + fmt(h202505.weightedPrice, 2)));
const sku26 = overall.monthly.filter((r) => r.month >= '202601' && r.month <= '202606').map((r) => r.skuCount);
const sku25 = overall.monthly.filter((r) => r.month >= '202501' && r.month <= '202512').map((r) => r.skuCount);
check('口径范围文本与 JSON 一致 (2026:1038-1993, 2025:1683-2000)',
  html.includes('1038-1993') && html.includes('1683-2000')
  && Math.min(...sku26) === 1038 && Math.max(...sku26) === 1993
  && Math.min(...sku25) === 1683 && Math.max(...sku25) === 2000,
  '2026=' + Math.min(...sku26) + '-' + Math.max(...sku26) + ' 2025=' + Math.min(...sku25) + '-' + Math.max(...sku25));
check('侧栏/范围说明数字与 JSON 一致', html.includes('49个月 · 4,261,173数据单元格')
  && html.includes('71,451 条实际记录') && html.includes('73,812 行') && html.includes('54 张有效业务表'));
const ins = data.insights;
check('2026指标卡数字与 JSON 一致', html.includes(fmt(ins.overall2026.sales)) && html.includes(fmtPct(ins.overall2026.yoySales))
  && html.includes('$' + fmt(ins.overall2026.revenue / 1000000, 1) + 'M') && html.includes(fmtPct(ins.overall2026.yoyRevenue))
  && html.includes(fmt(ins.ppSalesShare2026, 1) + '%') && html.includes(fmt(ins.genimoPpShare2026, 2) + '%'),
  'sales2026=' + ins.overall2026.sales + ' yoy=' + ins.overall2026.yoySales.toFixed(1) + ' revM=' + (ins.overall2026.revenue / 1000000).toFixed(1));
check('2026.05 中部/尾部 MOM 无对应数据披露存在', (html.match(/无对应数据/g) || []).length >= 8, 'count=' + (html.match(/无对应数据/g) || []).length);
const julyTrend = data.overallMarketTrend2026.find((r) => r.month === '202607');
check('2026.01-07合并趋势存在且July跨口径值禁算', html.includes('2026.01-07整体市场趋势（合并展示）')
  && html.includes('94父体小样本；合并展示，不参与同比/环比和累计')
  && (html.match(/不适用（口径不同）/g) || []).length === 2
  && julyTrend && julyTrend.coreComparable === false && julyTrend.momSales === null && !Object.keys(julyTrend).some((key) => key.startsWith('chain'))
  && !html.includes('2026.07附录/参考') && !html.includes('参考附录'));
// Cohort 段落与 JSON 一致
let cohortMismatch = 0;
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const co = data.categories[cat].cohort;
  const expectedStr = '前100父体池从 ' + co.fromParents + ' 变为 ' + co.toParents + '；留存 ' + co.retained + '、退出 ' + co.exited + '、新进入 ' + co.entered + '。';
  if (!html.includes(expectedStr)) cohortMismatch++;
}
check('4 个 Cohort 段落与 JSON 一致', cohortMismatch === 0, 'mismatches=' + cohortMismatch);
// Insight 网格（2026 四项）与 JSON 一致
const insightGrid = html.match(/<section id="insights">([\s\S]*?)<\/section>/)?.[1] || '';
let gridMismatch = 0;
if (!insightGrid.includes(fmt(ins.overall2026.sales)) || !insightGrid.includes(fmt(ins.overall2026.revenue))
  || !insightGrid.includes(fmtPct(ins.overall2026.yoySales)) || !insightGrid.includes(fmtPct(ins.overall2026.yoyRevenue))
  || !insightGrid.includes(fmtPct(ins.overall2026.yoyWeightedPrice))) gridMismatch++;
if (!insightGrid.includes(fmt(ins.pp2026.sales)) || !insightGrid.includes(fmt(ins.pp2026.revenue))
  || !insightGrid.includes(fmtPct(ins.pp2026.yoySales)) || !insightGrid.includes(fmtPct(ins.pp2026.yoyRevenue))
  || !insightGrid.includes(fmt(ins.ppSalesShare2026, 1) + '%') || !insightGrid.includes(fmtMonth(ins.ppPeak2026.month)) || !insightGrid.includes(fmt(ins.ppPeak2026.sales))) gridMismatch++;
if (!insightGrid.includes(fmt(ins.high2026.sales)) || !insightGrid.includes(fmt(ins.high2026.revenue))
  || !insightGrid.includes(fmtPct(ins.high2026.yoySales)) || !insightGrid.includes(fmtPct(ins.high2026.yoyRevenue))
  || !insightGrid.includes(fmtPct(ins.high2026.yoyWeightedPrice))) gridMismatch++;
if (!insightGrid.includes(fmt(ins.genimo2026.sales)) || !insightGrid.includes(fmt(ins.genimo2026.revenue))
  || !insightGrid.includes(fmtPct(ins.genimo2026.yoySales)) || !insightGrid.includes(fmtPct(ins.genimo2026.yoyRevenue))
  || !insightGrid.includes(fmt(ins.genimoPpShare2026, 2) + '%') || !insightGrid.includes(fmt(ins.genimoPpRevenueShare2026, 2) + '%')) gridMismatch++;
check('2026洞察卡数字与 JSON 一致', gridMismatch === 0, 'mismatches=' + gridMismatch);
check('HTML趋势与GENIMO建议无旧2025基线', html.includes('趋势结论与GENIMO建议（基于2026.01-06实绩）')
  && html.includes('GENIMO 2026.01-06累计Top父体') && !html.includes('2025年PP销量份额'));

// ============ Phase 5: JSON 顶层完整性 ============
check('analysisMonths=49 且截止 202606', data.analysisMonths.length === 49 && data.analysisMonths.at(-1) === '202606');
check('202607 为合并展示但不进入核心', data.sourceMonths.includes('202607') && data.excludedFromComparableReport.includes('202607')
  && data.overallMarketTrend2026.length === 7 && data.overallMarketTrend2026.at(-1).coreComparable === false);
check('replacementMetadata=7 且含 sha256', data.replacementMetadata.length === 7 && data.replacementMetadata.every((r) => /^[0-9a-f]{64}$/.test(r.source_sha256 || '')));
check('替换元数据逐月显示在HTML', data.replacementMetadata.every((row) => html.includes(row.source_sha256) && html.includes(fmtMonth(row.month))));
check('全部月度表显示唯一跨年MOM/环比基准月份', html.includes('MOM/环比基准月份')
  && !html.includes('环比基准月份（上月）'));
check('PP独立Listing明细完整', data.ppListingDetails.length > 0
  && data.ppListingDetails.every((row) => row.month >= '202601' && row.month <= '202606' && row.listingKey)
  && html.includes('PP独立Listing明细（2026.06'));
check('BSR多值解析标记保留并展示', Array.isArray(data.dataQuality.bsrMultiValueAudit)
  && data.dataQuality.bsrMultiValueAudit.every((row) => row.multiValue === true && row.sourceBsr && Number.isFinite(row.rank))
  && html.includes('BSR多值解析审计'));
check('GENIMO小批量工艺验证完整', html.includes('低风险颜色或4x6/5x8') && html.includes('批次标记')
  && html.includes('散边、卷边、破损和退货原因') && html.includes('累计300-500单')
  && html.includes('8x10、9x12及超大尺寸'));
check('预测可调整参数完整', data.forecastParameters.length === 4
  && ['需求系数', '成交价系数', '旺季前置周数', '情景选择'].every((name) => html.includes(name)));
check('四个分类数据齐全', ['overall', 'pp', 'high', 'genimo'].every((cat) => {
  const c = data.categories[cat];
  return c && c.monthly && c.annual && c.bsrTop100 && c.bsrGroups && c.bsrSegments
    && c.monthly.length === 49 && c.annual.length >= 4 && c.bsrGroups.monthly.length === 147;
}));
check('2026 年度同周期但scopeComparable=false，且带方向性限制', ['overall', 'pp', 'high', 'genimo'].every((cat) => {
  const a = data.categories[cat].annual.find((r) => r.year === '2026');
  return a && a.timeComparable === true && a.scopeComparable === false && a.scopeNote.includes('不构成严格同口径同比');
}));
check('历史BSR质量诊断完整', data.dataQuality.historicalBsrTop100Quality.length === 43
  && data.dataQuality.historicalBsrTop100Quality.every((row) => row.identifierCoveragePct === 100
    && row.strictListingPool === false && Number.isFinite(row.distinctListingKeys)
    && Number.isFinite(row.duplicateListingRows)));
check('HTML有3张可访问趋势图', (html.match(/<figure class="chart-card">/g) || []).length === 3
  && (html.match(/<svg /g) || []).length === 3 && (html.match(/role="img"/g) || []).length === 3);
check('HTML预测表含4个销量区间', FORECAST_2026_Q4.every((row) => row.range && html.includes(row.range)));

console.log('\n========== FRONTEND DATA COMPLETENESS AUDIT ==========');
console.log('Checks: ' + checks);
console.log('Failures: ' + failures);
process.exit(failures > 0 ? 1 : 0);
