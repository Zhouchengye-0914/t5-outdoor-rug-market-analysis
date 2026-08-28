# SPEC: T5 户外地垫市场数据 -> SQLite 转换工程

## 1. 外部资源与依赖清单 (Prerequisites)

### 数据源
- 文件: data/raw/地垫-卖家精灵市场数据.xlsx (44 MB)
- 来源: 飞书多维表格导出 / 卖家精灵市场数据
- 站点: 美国站 (Competitor-US-)
- 时间跨度: 2022.06 - 2026.07
- 工作簿内部共登记 55 张工作表，但 Excel 界面只显示 23 张：
  - 可见表 23 张：4 张 TOP 汇总 + 19 张月度明细（2025.01 - 2026.07）
  - 隐藏表 32 张：31 张月度历史明细（2022.06 - 2024.12）+ 1 张无有效区域的遗留表 `Sheet6`
- 有效业务数据共 54 张表：4 张 TOP 汇总 + 50 张月度明细。`Sheet6` 不属于业务数据，不应称为“1 张空业务表”。
- 原始需求要求分析“历年”数据并计算年度/月度 YoY、MoM，因此导入范围必须包含 31 张隐藏历史明细；工作表是否隐藏只属于 Excel 展示属性，不作为排除数据的条件。

### 运行环境
- Node.js >= 22.5 (已验证 v24.15.0, 内建 node:sqlite)
- 无需 Python 或数据库服务端

### 核心依赖包
- xlsx (SheetJS): 解析 .xlsx
- node:sqlite (Node.js >= 22.5 内建): 写入 SQLite (同步 API, 零原生依赖)
- dotenv: 加载 .env

## 2. 配置说明

代码中读取 .env 必须严格使用以下变量名:

| 变量名 | 类型 | 必填 | 说明 |
|---|---|---|---|
| RAW_EXCEL_PATH | string | 是 | Excel 源文件路径 |
| DATABASE_URL | string | 是 | SQLite URL, 形如 sqlite:///data/processed/market.db |
| IMPORT_OVERWRITE | bool | 否 | 是否覆盖已有 DB, 默认 true |
| LOG_LEVEL | string | 否 | debug/info/warn/error |
| ANALYSIS_DB_PATH | string | 否 | 分析生成器读取的 SQLite 路径 |
| ANALYSIS_CUTOFF | YYYYMM | 否 | 可比分析截止月份，默认 202602 |

## 3. 数据结构 (Schema)

### 表命名规则

- 月度明细表: monthly_<YYYYMM>, 例如 monthly_202206
- TOP 销量表: top_sales_volume
- TOP 销量倍率表: top_sales_volume_ratio
- TOP 总销售额表: top_total_sales
- TOP 平均单价表: top_avg_price
- 元数据表: meta (记录每次导入的元信息)
- 工作表目录: sheet_catalog (记录工作表顺序、可见性、有效区域、分类、目标表、导入行数和跳过原因)

### 工作表范围与可见性口径

- SheetJS 的 `wb.SheetNames` 会同时返回可见表和隐藏表，不能将其长度解释为 Excel 界面可见子表数。
- 可见性读取自 `wb.Workbook.Sheets[].Hidden`：`0` 为可见，`1` 为隐藏，`2` 为深度隐藏。
- 当前源文件统计：`total_sheets=55`、可见 23、隐藏 32、有效导入 54、跳过无有效区域 1。
- 转换器应保留全部有效业务表；验证和文档必须分别报告“工作簿内部表数”“界面可见表数”“隐藏表数”“有效导入表数”。
- 空表判断必须依据工作表是否存在有效区域 `!ref`，不能只依赖固定名称 `Sheet6`。

### 通用列规则

- 第一列固定为 row_id (INTEGER, 自增主键), 保证顺序
- 第二列固定为 month_label (TEXT), 规范化月份标签
- 原始 Excel 中的列名直接转 SQLite 列名, 特殊字符替换为下划线
- 所有列类型保持 SQLite 亲和类型 (TEXT/INTEGER/REAL)
- 空值统一写 NULL
- 启用 PRAGMA journal_mode=WAL 提高并发读

### 月度明细表统一字段

| 列名 | 类型 | 含义 |
|---|---|---|
| row_id | INTEGER PK | 行号 |
| month_label | TEXT | 月份标签 |
| ASIN | TEXT | 商品 ASIN |
| SKU | TEXT | SKU/规格 |
| 详细参数 | TEXT | 详细参数 |
| 品牌 | TEXT | 品牌名 |
| 品牌链接 | TEXT | 品牌搜索链接 |
| 商品标题 | TEXT | 商品标题 |
| 商品详情页链接 | TEXT | 商品链接 |
| 商品主图 | TEXT | 主图 URL |
| 父ASIN | TEXT | 父 ASIN |
| 类目路径 | TEXT | 类目路径 |
| 大类目 | TEXT | 大类目名称 |
| 标签 | TEXT | 标签 (仅 2026.3+ sheet) |
| 大类BSR | INTEGER | 大类 BSR 排名 |
| 大类BSR增长数 | INTEGER | 大类 BSR 增长数 |
| 大类BSR增长率 | REAL | 大类 BSR 增长率 |
| 小类目 | TEXT | 小类目 |
| 小类BSR | INTEGER | 小类 BSR 排名 |
| 月销量 | INTEGER | 月销量 |
| 月销量增长率 | REAL | 月销量增长率 |
| 月销售额 | REAL | 月销售额($) |
| 月销售额增长率 | TEXT | 月销售额增长率 |
| 子体销量 | INTEGER | 子体销量 |
| 子体销售额 | REAL | 子体销售额($) |
| 变体数 | INTEGER | 变体数 |
| 价格 | REAL | 价格($) |
| prime价格 | REAL | prime 价格($) |
| Coupon | TEXT | Coupon |
| Q_A | INTEGER | Q&A 数 |
| 评分数 | INTEGER | 评分数 |
| 月新增评分数 | INTEGER | 月新增评分数 |
| 评分 | REAL | 评分 |
| 留评率 | REAL | 留评率 |
| FBA | REAL | FBA($) |
| 毛利率 | REAL | 毛利率 |
| 评级 | TEXT | 评级 |
| 上架时间 | TEXT | 上架时间 |
| 上架天数 | INTEGER | 上架天数 |
| 配送方式 | TEXT | 配送方式 |
| 买家运费 | REAL | 买家运费 (仅 2025.6+) |
| LQS | INTEGER | LQS |
| 卖家数 | INTEGER | 卖家数 |
| BuyBox卖家 | TEXT | BuyBox 卖家 |
| BuyBox类型 | TEXT | BuyBox 类型 |
| 卖家所属地 | TEXT | 卖家所属地 |
| 卖家信息 | TEXT | 卖家信息 |
| 卖家首页 | TEXT | 卖家首页 |
| Best_Seller标识 | TEXT | Best Seller |
| Amazon_Choice | TEXT | Amazon's Choice |
| New_Release标识 | TEXT | New Release |
| A页面 | TEXT | A+ 页面 |
| 视频介绍 | TEXT | 视频介绍 |
| SP广告 | TEXT | SP 广告 |
| 品牌故事 | TEXT | 品牌故事 |
| 品牌广告 | TEXT | 品牌广告 |
| 7天促销 | TEXT | 7天促销 |
| AC关键词 | TEXT | AC 关键词 |
| 商品重量 | TEXT | 商品重量 |
| 商品重量_单位换算 | TEXT | 商品重量换算 |
| 商品尺寸 | TEXT | 商品尺寸 |
| 商品尺寸_单位换算 | TEXT | 商品尺寸换算 |
| 包装重量 | TEXT | 包装重量 (仅 2026.3+) |
| 包装重量_单位换算 | TEXT | 包装重量换算 |
| 包装尺寸 | TEXT | 包装尺寸 |
| 包装尺寸_单位换算 | TEXT | 包装尺寸换算 |
| 包装尺寸分段 | TEXT | 包装尺寸分段 (仅 2026.3+) |

> 旧 sheet 中 重量/体积 与新 sheet 的 商品重量/商品尺寸 含义相同, 统一列名.
> 文本中含括号/$/特殊字符的列名一律替换为下划线.
> Q&A 中的 & 转为下划线 (Q_A).
> Amazon's Choice 中的撇号转为下划线 (Amazon_Choice).

### TOP 汇总表 (4 张) 结构相同

| 列名 | 类型 | 含义 |
|---|---|---|
| row_id | INTEGER PK | 行号 |
| month_label | TEXT | sheet 名 |
| rank | INTEGER | 排名 |
| group | TEXT | 组别 (仅 TOP销量 有) |
| <YYYYMM> 系列 | INTEGER | 每月销量/销售额 |

例: top_sales_volume 表第 1 行: rank=1, group=, 202206=10329, 202207=13063, ...

### meta 元数据表

| 列名 | 类型 | 含义 |
|---|---|---|
| id | INTEGER PK AUTOINCREMENT | 导入批次 ID |
| imported_at | TEXT | ISO 时间戳 |
| source_file | TEXT | 源 Excel 文件名 |
| source_size_bytes | INTEGER | 源文件字节数 |
| total_sheets | INTEGER | 工作簿内部登记的 sheet 总数（含隐藏及无有效区域工作表） |
| monthly_tables | INTEGER | 月度表数 |
| summary_tables | INTEGER | 汇总表数 |
| total_rows | INTEGER | 总记录行数 |
| db_size_bytes | INTEGER | 生成的 DB 字节数 |
| schema_version | TEXT | schema 版本 |

### sheet_catalog 工作表目录

| 列名 | 类型 | 含义 |
|---|---|---|
| sheet_order | INTEGER PK | 源工作簿顺序（1-based） |
| sheet_name | TEXT | 源工作表名称 |
| visibility | TEXT | visible / hidden / very_hidden |
| hidden_code | INTEGER | SheetJS Hidden 原始代码 |
| sheet_ref | TEXT | Excel 有效区域 `!ref` |
| classification | TEXT | monthly / top / empty |
| target_table | TEXT | 目标业务表；跳过时为 NULL |
| imported_rows | INTEGER | 实际导入行数 |
| skip_reason | TEXT | 跳过原因 |

## 4. 技术方案与架构

### 转换流程
```
Excel (.xlsx)
  -> xlsx.readFile() [SheetJS]
  -> 自动识别 header 行 (row 0 或 row 1)
  -> 读取工作表可见性并按内容/名称分类: TOP* / 月度 / 无有效区域 / 未识别
  -> 统一列名 (去括号/$/特殊字符)
  -> 类型推断 (数值列 -> INTEGER/REAL, 其余 -> TEXT)
  -> 批量 INSERT (事务)
SQLite (market.db)
```

### 关键决策
1. 动态 schema: 脚本启动时根据 Excel sheet 动态生成表结构, 避免遗漏新列
2. 批量 INSERT: 单事务 (BEGIN/COMMIT) 包整张 sheet, 速度与稳定性平衡
3. 空值归一: 空字符串/undefined/null 全部写 NULL; 纯空白格视为 NULL
4. header 探测:
   - 月度表: 扫描 row 0 和 row 1, 找到含"品牌"/"商品标题"/"月销量"的行
   - TOP 表: header 在 row 0 (经审计验证)
5. 月份标签规范化: 2025.6 -> 202506, 202501 -> 202501
6. 覆盖策略: `IMPORT_OVERWRITE=true` 时重建 DB；为 `false` 且目标 DB 已存在时必须拒绝执行，不能静默覆盖
7. 空 header 列保留: TOP销量 (倍率) 等表有空 header 但有数据, 自动命名为 `_col_<index>` 保留
8. 隐藏历史表纳入: 31 张隐藏月度表是历年 YoY 的必要基线，不能因隐藏状态而跳过
9. 无有效区域表跳过: 当前 `Sheet6` 没有 `!ref`，不创建业务表、不计入有效导入表数

## 5. 验证过的 API 片段 (PoC)

### 5.1 xlsx 读取
```javascript
const xlsx = require('xlsx');
const wb = xlsx.readFile('data/raw/地垫-卖家精灵市场数据.xlsx', {
  cellDates: false,
  cellNF: false,
  cellFormula: false,
});
console.log(wb.SheetNames.length); // 55（内部总数，不等于界面可见表数 23）
```

### 5.2 node:sqlite 写入 (Node.js >= 22.5 内建)
```javascript
const { DatabaseSync } = require('node:sqlite');
const db = new DatabaseSync('data/processed/market.db');
db.exec('PRAGMA journal_mode = WAL;');
const stmt = db.prepare('INSERT INTO monthly_202206 (asin, brand) VALUES (?, ?)');
db.exec('BEGIN');
for (const r of rowsArray) stmt.run(r.asin, r.brand);
db.exec('COMMIT');
```

## 6. 验收标准

1. 范围完整性: 识别 23 张可见表、32 张隐藏表；导入 50 张月度明细 + 4 张 TOP 汇总；跳过 1 张无 `!ref` 的遗留表
2. 数据完整性: 50 个月度表的行数与 Excel 中 !ref 行数一致 (误差 0)
3. 列完整性: 每个有效 sheet 写入 SQLite 的源数据列数等于 Excel 有效区域列数
4. 类型一致性: 数值列实际为 REAL/INTEGER; 文本列实际为 TEXT
5. 空值归一: 空 cell 全部为 NULL, 无空字符串
6. meta 表: 每次导入写入一条元数据记录，且表数口径不得把 55 解释为可见表数
7. DB 文件: 单文件 data/processed/market.db, 可被任何 SQLite 客户端打开
8. 抽查验证: 至少覆盖隐藏旧月、可见月份、TOP 汇总和最新月份
9. 全量验证: `tests/full_audit.js` 必须存在并可复跑，对 54 张有效业务表逐行逐列比对；不得只在文档中声称通过
10. 覆盖保护: `IMPORT_OVERWRITE=false` 且 DB 已存在时，导入命令必须非零退出且原 DB 不变
11. 值与存储类型分别验收: 源数值与 DB 数值等值不代表 SQLite 存储类型一致；全量审计必须分别报告 value mismatch 与 numeric type change
12. 数据口径异常处置: 2026.03 - 2026.07 源 Excel 为子体级展开口径（销量虚增约 10 倍），已用竞品快照按父 ASIN 去重数据替换（见第 7 章）；替换表须保留列名交集并填充 month_label


## 7. 后续更新：2026 年数据口径修正与分析规格（2026-08-27）

### 7.1 新增数据源：竞品快照（每月 3000 行子体级）

| 文件 | 月份 | 导出批次 |
|---|---|---|
| Competitor-US-2026.01-809631.xlsx | 2026.01 | 809631 |
| Competitor-US-2026.02-809622.xlsx | 2026.02 | 809622 |
| Competitor-US-2026.03-809620.xlsx | 2026.03 | 809620 |
| Competitor-US-2026.04-809615.xlsx | 2026.04 | 809615 |
| Competitor-US-2026.05-809606.xlsx | 2026.05 | 809606 |
| Competitor-US-2026.06-809594.xlsx | 2026.06 | 809594 |
| Competitor-US-2026.07-809440.xlsx | 2026.07 | 809440 |

结构：单数据 sheet（Competitor-US-YYYYMM）+ Notes；65 列商品明细（ASIN/SKU/父ASIN/品牌/商品标题/大类BSR/小类BSR/月销量/月销售额/价格 等），子体级（每 ASIN 一行）。

### 7.2 新数据库：data/processed/competitor_809440.db

- `raw_YYYYMM`：原始 3000 行导入（含月销量/月销售额）
- `dedup_YYYYMM`：按父 ASIN 去重（同一父 ASIN 多子体随机保留 1 条；无父 ASIN 按自身保留）

去重结果：202601=64、202602=74、202603=82、202604=79、202605=88、202606=91、202607=94 个父商品。

### 7.3 market.db 月度表替换规则（2026 全年）

- `monthly_202601` ~ `monthly_202607` 已用 `dedup_YYYYMM` 替换（DROP + CREATE + INSERT）
- 列映射取目标表与 dedup 表列名交集；缺失列（如 月销量增长率、_7天促销）置 NULL；month_label 回填月份
- 替换后 2026 全年为竞品父 ASIN 去重口径（64-94 商品/月），与 2025 全市场口径（1700-2000 SKU）跨年对比存在范围差异，报告须标注

### 7.4 分析口径（用户确认的 B 方案）

| 指标 | 定义 |
|---|---|
| MOM（月度） | 今年 X 月 vs 去年 X 月（跨年同月，如 2025.01 vs 2024.01） |
| 环比（月度） | 本月 vs 上月（连续月环比），单独列示 |
| YoY（年度） | 年度同周期对比（今年 vs 去年同月份集合） |
| 可比口径列 | 2026 年行额外输出 `cleanYoySales/cleanYoyRevenue`（2026.01-02 vs 2025.01-02），规避口径突变失真 |

### 7.5 分析分类定义（analyze_market.js）

| 类别 | 筛选规则 |
|---|---|
| overall | 全部商品 |
| pp | 商品标题（COALESCE 空串）包含 `plastic`（不区分大小写） |
| high | 非 PP 且（标题含 polypropylene/sandwich/woven/braided 之一，或 价格 ≥ 40） |
| genimo | 品牌（COALESCE 空串）等于 `genimo`（不区分大小写） |

### 7.6 BSR 前 100 分析结构（每类别 × 50 月）

- `bsrTop100`：BSR 1-100 汇总（月度 + 年度）
- `bsrGroups`：头部（1-20）/中部（21-50）/尾部（51-100）（月度 + 年度）
- `bsrSegments`：五档 1-5/6-10/11-20/21-50/51-100（月度 + 年度）
- BSR 解析：小类BSR 取可解析最小名次；多值标记保留

### 7.7 分析输出与交付

- `src/analyze_market.js` 从 market.db 生成三份交付物（数据全部代码生成，禁止硬编码）：
  - 交付/户外地垫市场分析数据.json（全量结构化中间结果）
  - 交付/户外地垫市场分析报告-优化版.md
  - 交付/户外地垫市场分析报告-优化版.html（大字版，body 18px/表格 15px）
- 每类别含：月度指标（MOM+环比）、年度汇总（YoY+可比列）、BSR前100（汇总/头中尾/五档，月度+年度）、趋势分析（2025 同周期/BSR占比/头中尾同比/最近可比月 MOM+环比/季节性/口径提示）
- 线上部署：gh-pages 分支（index.html = 优化版 HTML），GitHub Pages 托管
- 注意：analyze_market.js 的月度 MOM/环比口径必须与 7.4 一致（曾有外部改动把 MOM 恢复成连续环比，提交前须校验 addTrends 中 mom=lastYear、chain=previous）

### 7.8 已知限制

- 2026 年为竞品监控口径（64-94 商品），2025 年为全市场口径，跨年同比存在范围差异（报告已标注）
- 2026 年商品不足 100，BSR 51-100 档数据稀疏
- 需要 2025 年竞品快照或更大范围竞品列表才能做完全同口径跨年对比
