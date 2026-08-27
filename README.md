# T5 户外地垫市场数据 -> SQLite 转换工程

> Excel (卖家精灵 / 飞书多维表格导出) -> 完整无损 -> SQLite (单文件 DB)
>
> 遵循 [`.config/单个项目开发规范手册`](../.config/单个项目开发规范手册：AI 驱动开发的工程化标准.md) 三文档闭环 (SPEC / TASKS / REVIEWS)。

## 1. 项目概述

- **目的**: 把 44 MB 的多 sheet Excel 文件（50 张月度明细 + 4 张 TOP 汇总）无损导入到 SQLite, 便于后续 BI / 二次分析。
- **数据源**: `data/raw/地垫-卖家精灵市场数据.xlsx` (卖家精灵市场数据, 美国站, 2022.06 - 2026.07)
- **输出**: `data/processed/market.db` (单文件 SQLite, ~48 MB)
- **工作表口径**: Excel 界面可见 23 张；工作簿内部另有 32 张隐藏表（31 张历史月度明细 + 1 张无有效区域的遗留 `Sheet6`）
- **规模**: 73,812 条产品记录, 54 张业务表（50 月度 + 4 TOP）+ 1 张 meta 元数据表

> 原始需求要求“历年”分析及年度/月度 YoY、MoM，因此 31 张隐藏历史月度表属于有效数据并纳入转换。55 是工作簿内部登记总数，不是 Excel 界面可见子表数；隐藏的空白 `Sheet6` 被跳过。

## 2. 目录结构

```
.
├── .env                       # 实际配置 (gitignore)
├── .env.example               # 配置模板 (提交到 git)
├── .gitignore                 # 屏蔽 data/raw, *.db, *.xlsx, node_modules
├── README.md                  # 本文件
├── package.json               # npm 元数据
├── docs/
│   ├── SPEC.md                # 唯一事实来源 (schema / 验收标准)
│   ├── TASKS.md               # 任务清单 (全部 [x])
│   └── REVIEWS.md             # 审查草稿 (默认空)
├── data/
│   ├── raw/地垫-卖家精灵市场数据.xlsx   # 原始 (gitignore)
│   └── processed/market.db             # 转换结果 (gitignore)
├── src/
│   └── import_xlsx.js         # 主转换脚本
├── sandbox/
│   └── test_xlsx.js           # PoC (验证 xlsx 能读取)
├── tests/
│   └── verify.js              # 数据完整性验证
├── tmp/                       # 临时文件 (可随时清空)
└── node_modules/              # 依赖 (gitignore)
```

## 3. 快速开始

```bash
# 1. 安装依赖
npm install

# 2. 运行 PoC (验证 xlsx 能读取)
npm run poc

# 3. 执行转换
npm run import

# 4. 验证数据完整性
npm run verify

# 5. 全量逐 cell 审计 + 覆盖保护
npm run audit
npm run test:overwrite

# 6. 生成优化版结构化数据、Markdown 和独立 HTML
npm run analyze
```

## 4. DB Schema 概览

### 4.1 命名规则

- 月度明细表: `monthly_<YYYYMM>` (例如 `monthly_202206`, `monthly_202602`)
- TOP 销量表: `top_sales_volume` / `top_sales_volume_ratio` / `top_total_sales` / `top_avg_price`
- 元数据表: `meta` (记录每次导入)
- 工作表目录: `sheet_catalog`（记录顺序、可见性、有效区域、分类、目标表、行数和跳过原因）

### 4.2 通用列

每张业务表都额外包含:

- `row_id` (INTEGER PRIMARY KEY AUTOINCREMENT)
- `month_label` (TEXT, 规范化月份标签)

详见 `docs/SPEC.md` 第 3 节。

### 4.3 示例查询

```sql
-- 查 2024 年 5 月销量最高的 10 个商品
SELECT 品牌, 商品标题, 月销量, 月销售额, 价格
FROM monthly_202405
ORDER BY 月销量 DESC
LIMIT 10;

-- 看 meta
SELECT * FROM meta;
```

### 4.4 跨月聚合 (注意: 上架时间列类型不一致, 需 CAST)

```sql
-- 跨月销量聚合 (无类型冲突)
SELECT month_label, SUM(月销量) AS sales, SUM(月销售额) AS revenue
FROM monthly_202206 GROUP BY month_label
UNION ALL
SELECT month_label, SUM(月销量), SUM(月销售额) FROM monthly_202207 GROUP BY month_label
...;

-- 上架时间过滤: 旧 TEXT vs 新 INTEGER Excel 序列号
SELECT 品牌, 商品标题, 上架时间
FROM monthly_202206 WHERE 上架时间 > '2024-01-01'
UNION ALL
SELECT 品牌, 商品标题, 上架时间
FROM monthly_202406 WHERE CAST(上架时间 AS TEXT) > '2024-01-01';

-- 上架日期统一为 ISO 格式
SELECT 品牌, 商品标题, 上架时间,
  CASE
    WHEN typeof(上架时间) = 'integer' THEN date(上架时间 - 25569, 'unixepoch')
    ELSE 上架时间
  END AS 上架日期
FROM monthly_202406;

-- 重量 / 重量_2 (旧 sheet 重复列, 第一个磅, 第二个千克)
SELECT 品牌, 商品标题, 重量 AS 重量_磅, 重量_2 AS 重量_千克
FROM monthly_202206
WHERE 重量 IS NOT NULL LIMIT 10;
```

### 4.5 数据质量注意事项

| 列 | 类型不一致 | 原因 |
|---|---|---|
| `上架时间` | TEXT (旧) vs INTEGER (新, Excel 序列号) | 早期文本, 后期 Excel 日期 |
| `子体销量` / `子体销售额` | TEXT (旧, 多空值) vs INTEGER (新) | 旧 sheet 未填充 |
| `Coupon` | TEXT (旧, 多空值) vs REAL (新) | 同上 |
| `LQS` | INTEGER (45) vs REAL (5 张 2026.3+) | 精度差异 |
| `月销量增长率` | 4 张 (2026.4-2026.7) 缺失 | 源 Excel 无此列 |
| `买家运费` | 仅 26 张 (2025.6+) | 源 Excel 新版加列 |
| `小类BSR` | 部分值含换行 | 单元格多行内容 |
| `卖家信息` | 326/494 行含换行 | 同上 |
| `重量` / `体积` | 旧 sheet 重复列名 | 磅 + 千克 |

详见 `docs/REVIEWS.md` 完整记录.

## 5. 关键决策

| 决策点 | 选型 | 理由 |
|---|---|---|
| DB 引擎 | SQLite | 零依赖、单文件、易 gitignore |
| SQLite 驱动 | Node.js 22.5+ 内建 `node:sqlite` | 避免原生编译 (better-sqlite3 需要 Python + node-gyp) |
| Excel 库 | xlsx (SheetJS) | 成熟、社区大、支持流式 |
| Schema | 动态生成 | 新增 sheet / 新增列无需改脚本 |
| 空值 | 全部 NULL | 避免空字符串歧义 |
| 月份标签 | `2025.6` -> `202506` | 统一规范, 便于 ORDER BY / GROUP BY |

## 6. 验收结果

### 6.1 标准验收 (tests/verify.js)

```
[PASS] meta has 1 row (rows=1)
[PASS] meta imported_at is string
[PASS] meta source_file = xlsx
[PASS] meta schema_version = 1.0.0

Total tables in DB: 56
Checked 54 tables, mismatches=0

[PASS] sample monthly_202206 (rows=6)
[PASS] sample monthly_202506 (rows=6)
[PASS] sample monthly_202607 (rows=6)

========== Summary ==========
PASS: 61
FAIL: 0
```

### 6.2 全表逐 cell 校验（历史记录，当前待恢复脚本）

```
========== GRAND TOTAL ==========
Sheets checked: 54
Cells checked: 4,441,989
Mismatches: 0
```

历史文档记录为 50 个月度表 + 4 张 TOP 表、4,441,989 个单元格零差异；但当前仓库缺少所述 `full_audit.js`，该结论在脚本恢复并重新运行前只能视为历史记录，不能作为当前可复验结论。

2026-08-27 独立只读复核再次检查了 54 张有效业务表、4,441,989 个单元格：结构问题 0、值不一致 0；同时发现 70,989 个 Excel 数值在混合类型列中以 SQLite TEXT 保存。该复核证明值未丢失，但一次性核验命令尚未恢复为仓库内可重复运行的 `tests/full_audit.js`。

### 6.3 自审修复历史

详见 `docs/REVIEWS.md`. 自审发现并修复了 4 个 bug:

| # | Bug | 修复 | 影响范围 |
|---|---|---|---|
| #1 | TOP 表 headerRow 错位 (1→0) | import_xlsx.js | 4 张 TOP 表共缺 1 行 / 表 + 列名错位 |
| #2 | 空 header 列数据被丢弃 | import_xlsx.js 用 `_col_N` 占位 | TOP销量(倍率) 丢失 3697 cell |
| #3 | verify.js TOP header 检测错 | verify.js | 验证脚本误报 |
| #4 | 抽样对比范围不一致 (前6 vs 前3+后3) | verify.js getDbSample | 验证脚本误报 |

## 7. 性能指标

| 阶段 | 耗时 |
|---|---|
| Excel 加载 (44 MB) | 17.5 s |
| DB 写入 (73,812 行 × 54 表) | ~3 s |
| 总耗时 | 20.2 s |
| 输出 DB 大小 | 48.75 MB |

## 8. 后续可扩展

- **增量更新**: 增量读取新月份 sheet, 跳过已存在表 (本期未实现)
- **索引优化**: 对 `month_label`、`品牌`、`小类BSR` 加索引 (本期未加, 数据量不大)
- **导出 Parquet/CSV**: 用 DuckDB / sqlite-utils 做下游分发
- **API 暴露**: 用 better-sqlite3 + Express 起 RESTful 查询服务

## 9. 规范遵循

| 规范要求 | 落地情况 |
|---|---|
| docs/SPEC.md 唯一事实来源 | ✅ 含数据源 / schema / PoC / 验收 |
| docs/TASKS.md 原子化任务 | ✅ 12 个 Task 全部 [x] |
| docs/REVIEWS.md 审查草稿 | ✅ 已建 (默认空) |
| data/raw 不入 git | ✅ .gitignore 已配置 |
| data/processed/*.db 不入 git | ✅ .gitignore 已配置 |
| .env.example 配置唯一来源 | ✅ 含 RAW_EXCEL_PATH / DATABASE_URL 等 |
| sandbox/ 跑通 PoC | ✅ sandbox/test_xlsx.js 已验证 55 sheet |
| tmp/ 可随时清空 | ✅ 已清空 |
| 验证命令 inline 在每个 Task | ✅ 每个 Task 都标了验证方式 |
