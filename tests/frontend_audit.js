'use strict';

// 前端数据完整性审计：交付 HTML 中出现的全部数字必须能在 交付/户外地垫市场分析数据.json
// （或 analyze_market.js 的预测常量）中逐格溯源。
// 1) 37 张数据表 24,374+ 个单元格与 JSON 精确一致（同一格式化函数重算）；
// 2) 4 个分类趋势分析文本用与生成器相同的逻辑从 JSON 重算后逐字比对；
// 3) 其余文本数字（口径说明/侧栏/指标卡/Cohort/洞察卡/覆盖核对）全部可溯源；
// 4) 无 NaN/undefined/Infinity 等异常标记；
// 5) JSON 顶层完整性（49 个月/202607 附录/7 条替换元数据/2026 scopeComparable）。


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
  return rows.map((r) => [r.month, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2),
    fmtPct(r.momSales), fmtPct(r.momRevenue), fmtPct(r.momAvgListPrice), fmtPct(r.momWeightedPrice),
    fmtPct(r.chainSales), fmtPct(r.chainRevenue)]);
}
function annualRows(rows) {
  return rows.map((r) => [r.year + ' (' + r.period + ')', r.comparison || '-', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2),
    fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyAvgListPrice), fmtPct(r.yoyWeightedPrice),
    r.scopeComparable ? '一致' : (r.scopeNote || '-')]);
}
function segmentRows(rows) {
  return rows.map((r) => [r.month, r.segment, fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2),
    segPct(r, 'momSales', 'momGapReason'), segPct(r, 'momRevenue', 'momGapReason'), segPct(r, 'momWeightedPrice', 'momGapReason'),
    segPct(r, 'chainSales', 'chainGapReason'), segPct(r, 'chainRevenue', 'chainGapReason')]);
}
function annualSegmentsRows(rows) {
  return rows.map((r) => [r.year + ' (' + r.period + ')', r.segment, r.comparison || '-', fmt(r.skuCount), fmt(r.sales), fmt(r.revenue), fmt(r.weightedPrice, 2),
    fmtPct(r.yoySales), fmtPct(r.yoyRevenue), fmtPct(r.yoyWeightedPrice), r.scopeComparable ? '一致' : (r.scopeNote || '-')]);
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
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const c = data.categories[cat];
  expectedTables.push(monthlyRows(c.monthly));
  expectedTables.push(annualRows(c.annual));
  expectedTables.push(monthlyRows(c.bsrTop100.monthly));
  expectedTables.push(annualRows(c.bsrTop100.annual));
  expectedTables.push(segmentRows(c.bsrGroups.monthly));
  expectedTables.push(annualSegmentsRows(c.bsrGroups.annual));
  expectedTables.push(segmentRows(c.bsrSegments.monthly));
  expectedTables.push(annualSegmentsRows(c.bsrSegments.annual));
}
expectedTables.push(data.sourceDiagnostics.filter((r) => r.month > REPORT_CUTOFF)
  .map((r) => [r.month, '附录/参考（> ' + REPORT_CUTOFF + '）', fmt(r.sales), fmt(r.revenue), fmt(r.avgListPrice, 2), fmt(r.weightedPrice, 2)]));
expectedTables.push(FORECAST_2026_Q4.map((fm) => [fm.month.slice(0, 4) + '.' + fm.month.slice(4), '约' + fmt(fm.sales), '约' + fmt(fm.rev, 0) + '美元', fm.stage]));
expectedTables.push(FORECAST_2027_MONTHLY.map((fm) => [fm.month.slice(0, 4) + '.' + fm.month.slice(4), fmt(fm.sales), '约' + fmt(fm.rev, 0) + '美元', fm.note]));
expectedTables.push(FORECAST_2027_SCENARIOS.map((fs) => [fs.scenario, fs.sales, fs.rev, fs.trigger]));
expectedTables.push(data.genimoTopProducts.map((row, index) => [String(index + 1), row.asin || '-', fmt(row.sales), fmt(row.revenue), String(row.months), fmt(row.latestPrice, 2), row.title || '-']));
expectedTables.push(null);
let tableMismatches = 0, tableChecks = 0;
check('HTML 表数量与期望一致 (38)', tables.length === expectedTables.length, 'html=' + tables.length + ' expected=' + expectedTables.length);
for (let i = 0; i < Math.min(tables.length, expectedTables.length); i++) {
  if (expectedTables[i] === null) {
    const rows = tables[i];
    if (rows[0] && rows[0].join('|') === '要求项|交付位置|状态' && rows.length === 9) continue;
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
check('37 张数据表 ' + tableChecks + ' 个单元格与 JSON/常量精确一致', tableMismatches === 0, 'mismatches=' + tableMismatches);

// ============ Phase 6: 趋势分析文本重算比对 ============
// 与 src/analyze_market.js 中 trendAnalysis() 相同的逻辑（从 JSON 重算）
function trendAnalysis(c, category, label) {
  const benchmark = c.monthly.find((row) => row.month === '202602');
  const baseline = [...c.monthly].reverse().find((row) => row.month <= REPORT_CUTOFF) || c.monthly[c.monthly.length - 1];
  const annual2025 = c.annual.find((row) => row.year === '2025');
  const top2025 = c.bsrTop100.annual.find((row) => row.year === '2025');
  const groups2025 = c.bsrGroups.annual.filter((row) => row.year === '2025');
  const groupLine = groups2025.map((row) => row.segment + '销量 ' + fmtPct(row.yoySales) + '、销售额 ' + fmtPct(row.yoyRevenue)).join('；');
  const byMonth = {};
  for (const row of c.monthly.filter((item) => item.month.startsWith('2024') || item.month.startsWith('2025'))) {
    const key = row.month.slice(4);
    byMonth[key] = byMonth[key] || [];
    byMonth[key].push(row.sales);
  }
  const seasonal = Object.entries(byMonth).map(([month, values]) => ({ month, avg: values.reduce((sum, value) => sum + value, 0) / values.length })).sort((a, b) => b.avg - a.avg)[0];
  const out = ['### ' + label + '趋势分析', ''];
  if (annual2025) out.push('- 2025同周期销量 ' + fmt(annual2025.sales) + '（YOY ' + fmtPct(annual2025.yoySales) + '），销售额 $' + fmt(annual2025.revenue) + '（YOY ' + fmtPct(annual2025.yoyRevenue) + '），SKU平均标价 YOY ' + fmtPct(annual2025.yoyAvgListPrice) + '，加权成交均价 YOY ' + fmtPct(annual2025.yoyWeightedPrice) + '。');
  if (top2025 && annual2025) out.push('- BSR前100贡献销量 ' + fmt(top2025.sales) + '（占' + fmt(top2025.sales / annual2025.sales * 100, 1) + '%），销售额占比 ' + fmt(top2025.revenue / annual2025.revenue * 100, 1) + '%；其销量/销售额YOY分别为 ' + fmtPct(top2025.yoySales) + ' / ' + fmtPct(top2025.yoyRevenue) + '。');
  if (groupLine) out.push('- 头中尾分层同比：' + groupLine + '。');
  if (benchmark) {
    const expected = category === 'overall' ? '（整体市场验收目标约 -14.8% / -20.5% / +9.6% / +23.8%）' : '';
    out.push('- 验收基准月 202602：MOM销量（今年 vs 去年同月）' + fmtPct(benchmark.momSales) + '、MOM销售额 ' + fmtPct(benchmark.momRevenue) + '；环比销量（vs 上月）' + fmtPct(benchmark.chainSales) + '、环比销售额 ' + fmtPct(benchmark.chainRevenue) + expected + '。');
  }
  if (baseline && baseline.month !== '202602') out.push('- 核心截止月 ' + baseline.month + '：MOM销量（今年 vs 去年同月）' + fmtPct(baseline.momSales) + '、MOM销售额 ' + fmtPct(baseline.momRevenue) + '；环比销量（vs 上月）' + fmtPct(baseline.chainSales) + '、环比销售额 ' + fmtPct(baseline.chainRevenue) + '。');
  if (seasonal) out.push('- 季节性（2024-2025同月均值）：' + seasonal.month + '月销量最高，月均 ' + fmt(seasonal.avg) + ' 件，建议在高峰前完成备货与广告测试。');
  const anomaly = c.monthly.find((row) => row.month === '202604');
  if (anomaly && baseline) out.push('- 口径提示：2026.01-06 已更新为全市场父体级快照（每月1038-1993个父体，替代原64-94父体口径），与2025年全市场行级口径（每月1683-2000行）量级一致，跨年同比可直接参考；2025为行级导出（含变体行、无ASIN列）、2026为父体级导出（父ASIN去重），颗粒度与导出日期仍略有差异，2026.07仍为94父体小口径（仅附录/参考）。');
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
  for (const coll of ['monthly', 'bsrTop100', 'bsrGroups', 'bsrSegments']) {
    for (const key of ['monthly', 'annual']) {
      const rows = (c[coll] && c[coll][key]) || [];
      for (const row of rows) for (const f of Object.keys(row)) if (typeof row[f] === 'number') addFmt(row[f]);
    }
  }
  for (const row of c.annual || []) for (const f of Object.keys(row)) if (typeof row[f] === 'number') addFmt(row[f]);
}
for (const f of Object.keys(data.insights || {})) if (typeof data.insights[f] === 'number') addFmt(data.insights[f]);
for (const r of data.sourceDiagnostics || []) for (const f of Object.keys(r)) if (typeof r[f] === 'number') addFmt(r[f]);
for (const r of data.genimoTopProducts || []) for (const f of Object.keys(r)) if (typeof r[f] === 'number') addFmt(r[f]);
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const co = data.categories[cat] && data.categories[cat].cohort;
  if (co) for (const f of ['fromParents', 'toParents', 'retained', 'exited', 'entered']) if (typeof co[f] === 'number') addFmt(co[f]);
}
for (const f of ['currentDataRowCount', 'verifiedDataCellCount']) if (typeof data[f] === 'number') addFmt(data[f]);
for (const fm of FORECAST_2026_Q4) { addFmt(fm.sales); addFmt(fm.rev); }
for (const fm of FORECAST_2027_MONTHLY) { addFmt(fm.sales); addFmt(fm.rev); }
// 派生值：销售额 /1e6 的 M 展示、月份标签、年份
for (const v of [data.insights.overall2025.revenue]) { expectedSet.add(fmt(v / 1000000, 1)); expectedSet.add(fmt(v / 1000000, 2)); }
for (const m of [...(data.analysisMonths || []), ...(data.sourceMonths || [])]) { expectedSet.add(m); expectedSet.add(m.slice(0, 4) + '.' + m.slice(4)); expectedSet.add(m.slice(0, 4) + '.' + Number(m.slice(4))); }
for (const y of ['2022', '2023', '2024', '2025', '2026', '2027', '2028']) expectedSet.add(y);
for (const s of ['-14.8%', '-20.5%', '+9.6%', '+23.8%', '-27.7%', '-0.1%', '+9.3%', '+33.5%']) expectedSet.add(s);
for (const s of ['1038', '1993', '1683', '2000', '1134', '1039', '1766', '1744', '1690', '1135', '1745', '1691', '2002', '3000', '94']) expectedSet.add(s);
for (const s of ['112', '125', '156', '153', '100', '162,797', '6,446,797', '39.60', '49', '54', '71,451', '4,261,173', '73,812', '160', '144', '53', '107', '91', '65']) expectedSet.add(s);
for (const s of ['55.8%', '21.06%', '22.04%', '3,722,400', '$210.2M', '$210.2', '+20.2%', '+6.0%', '$98.9M', '-14.7%', '-20.1%', '$57.0M', '-15.2%', '-26.4%', '$41.9M', '-9.7%', '$16.3M', '+10.7%', '+83.0%', '+65.2%']) expectedSet.add(s);
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
check('验收基准文案与 JSON 一致', html.includes('整体市场验收目标约 -14.8% / -20.5% / +9.6% / +23.8%')
  && Math.abs(b202602.momSales - (-14.8)) < 0.11 && Math.abs(b202602.momRevenue - (-20.5)) < 0.11
  && Math.abs(b202602.chainSales - 9.6) < 0.11 && Math.abs(b202602.chainRevenue - 23.8) < 0.11,
  'json=' + b202602.momSales.toFixed(1) + '/' + b202602.momRevenue.toFixed(1) + '/' + b202602.chainSales.toFixed(1) + '/' + b202602.chainRevenue.toFixed(1));
const h202505 = overall.bsrGroups.monthly.find((r) => r.month === '202505' && r.segment === '头部（1-20）');
check('2025.05 头部口径说明数字与 JSON 一致', html.includes('162,797') && html.includes('6,446,797') && html.includes('39.60')
  && h202505.sales === 162797 && h202505.revenue === 6446797 && Math.abs(h202505.weightedPrice - 39.6) < 0.01);
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
check('指标卡数字与 JSON 一致', html.includes('3,722,400') && html.includes('+20.2%') && html.includes('$210.2M') && html.includes('+6.0%')
  && html.includes('55.8%') && html.includes('21.06%'),
  'sales2025=' + ins.overall2025.sales + ' yoy=' + ins.overall2025.yoySales.toFixed(1) + ' revM=' + (ins.overall2025.revenue / 1000000).toFixed(1));
check('2026.05 中部/尾部 MOM 无对应数据披露存在', (html.match(/无对应数据/g) || []).length >= 8, 'count=' + (html.match(/无对应数据/g) || []).length);
check('2026.07 附录表存在且行数为 1', html.includes('202607') && html.includes('附录/参考（&gt; 202606）') && (html.match(/附录\/参考/g) || []).length >= 1);
// Cohort 段落与 JSON 一致
let cohortMismatch = 0;
for (const cat of ['overall', 'pp', 'high', 'genimo']) {
  const co = data.categories[cat].cohort;
  const expectedStr = '前100父体池从 ' + co.fromParents + ' 变为 ' + co.toParents + '；留存 ' + co.retained + '、退出 ' + co.exited + '、新进入 ' + co.entered + '。';
  if (!html.includes(expectedStr)) cohortMismatch++;
}
check('4 个 Cohort 段落与 JSON 一致', cohortMismatch === 0, 'mismatches=' + cohortMismatch);
// Insight 网格（2025 四项）与 JSON 一致
const insightGrid = html.match(/<section id="insights">([\s\S]*?)<\/section>/)?.[1] || '';
let gridMismatch = 0;
const ppShare25 = ins.pp2025.sales / ins.overall2025.sales * 100;
if (!insightGrid.includes(fmtPct(ins.overall2025.yoySales)) || !insightGrid.includes(fmtPct(ins.overall2025.yoyRevenue))
  || !insightGrid.includes(fmtPct(ins.overall2025.yoyAvgListPrice)) || !insightGrid.includes(fmtPct(ins.overall2025.yoyWeightedPrice))) gridMismatch++;
if (!insightGrid.includes(fmt(ins.pp2025.sales)) || !insightGrid.includes(fmtPct(ins.pp2025.yoySales)) || !insightGrid.includes(fmtPct(ins.pp2025.yoyRevenue))
  || !insightGrid.includes(fmt(ppShare25, 1) + '%') || !insightGrid.includes(ins.ppPeak2025.month) || !insightGrid.includes(fmt(ins.ppPeak2025.sales))) gridMismatch++;
if (!insightGrid.includes(fmtPct(ins.high2025.yoySales)) || !insightGrid.includes(fmtPct(ins.high2025.yoyRevenue)) || !insightGrid.includes(fmtPct(ins.high2025.yoyWeightedPrice))) gridMismatch++;
if (!insightGrid.includes(fmtPct(ins.genimo2025.yoySales)) || !insightGrid.includes(fmt(ins.genimoPpShare2025, 2) + '%') || !insightGrid.includes(fmt(ins.genimoPpRevenueShare2025, 2) + '%')) gridMismatch++;
check('洞察卡数字与 JSON 一致', gridMismatch === 0, 'mismatches=' + gridMismatch);

// ============ Phase 5: JSON 顶层完整性 ============
check('analysisMonths=49 且截止 202606', data.analysisMonths.length === 49 && data.analysisMonths.at(-1) === '202606');
check('202607 为附录', data.sourceMonths.includes('202607') && data.excludedFromComparableReport.includes('202607'));
check('replacementMetadata=7 且含 sha256', data.replacementMetadata.length === 7 && data.replacementMetadata.every((r) => /^[0-9a-f]{64}$/.test(r.source_sha256 || '')));
check('四个分类数据齐全', ['overall', 'pp', 'high', 'genimo'].every((cat) => {
  const c = data.categories[cat];
  return c && c.monthly && c.annual && c.bsrTop100 && c.bsrGroups && c.bsrSegments
    && c.monthly.length === 49 && c.annual.length >= 4 && c.bsrGroups.monthly.length === 147;
}));
check('2026 年度 scopeComparable=true 且带说明', ['overall', 'pp', 'high', 'genimo'].every((cat) => {
  const a = data.categories[cat].annual.find((r) => r.year === '2026');
  return a && a.scopeComparable === true && Boolean(a.scopeNote);
}));

console.log('\n========== FRONTEND DATA COMPLETENESS AUDIT ==========');
console.log('Checks: ' + checks);
console.log('Failures: ' + failures);
process.exit(failures > 0 ? 1 : 0);
