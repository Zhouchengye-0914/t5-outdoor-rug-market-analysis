# TASKS: T5 户外地垫市场数据 -> SQLite 转换工程

> AI 执行准则:
> 1. 每次仅执行一个未勾选 [ ] 的最靠前原子任务
> 2. 完成后必须运行验证命令, 确认无误后打钩 [x]
> 3. 代码实现必须严格参照 SPEC.md

## Phase 1: 基础设施与技术验证 (PoC)

- [x] Task 1.1: 创建标准目录结构 + 移动原始 Excel
  - 内容: 创建 docs/ data/raw data/processed src sandbox tmp tests; 移动 Excel 到 data/raw/
  - 验证: ls 项目根目录能看到全部子目录, data/raw/地垫-卖家精灵市场数据.xlsx 存在

- [x] Task 1.2: 编写 .gitignore 与 .env.example
  - 内容: 按规范模板编写, 屏蔽 data/raw, *.db, *.xlsx, node_modules, .env 等
  - 验证: cat .gitignore / .env.example 显示正确内容

- [x] Task 1.3: 编写 docs/SPEC.md
  - 内容: 完整 schema 定义 + 数据源 + 验收标准
  - 验证: cat docs/SPEC.md 显示完整

- [x] Task 1.4: 编写 docs/TASKS.md (本文件)
  - 内容: 原子化任务清单
  - 验证: cat docs/TASKS.md 显示完整

- [x] Task 1.5: sandbox 验证 xlsx 解析能力
  - 参照: SPEC.md 5.1
  - 内容: 在 sandbox/test_xlsx.js 读取 data/raw/地垫-卖家精灵市场数据.xlsx, 输出 sheet 数 + 每个 sheet 的行数 + 列数
  - 验证: node sandbox/test_xlsx.js 输出 55 个 sheet 信息

- [x] Task 1.6: 安装核心依赖
  - 内容: 选用 Node.js 内建 node:sqlite (>=22.5), 安装 xlsx + dotenv
  - 验证: node -e "require('xlsx'); require('dotenv'); const {DatabaseSync} = require('node:sqlite');" 无报错

## Phase 2: 核心转换脚本

- [x] Task 2.1: 实现列名规范化工具
  - 内容: src/import_xlsx.js 提供 colNorm(name), 替换 ($()/单引号/空格) 等
  - 验证: 月度表 + TOP 表列名均无非法字符

- [x] Task 2.2: 实现类型推断
  - 内容: inferType(values) 返回 INTEGER/REAL/TEXT
  - 验证: 月销量=INTEGER, 月销售额=REAL, 商品标题=TEXT

- [x] Task 2.3: 实现 sheet 分类器
  - 内容: classifySheet(name) 返回 top / monthly / empty
  - 验证: 50 月度 + 4 TOP + 1 empty 全部归类正确

- [x] Task 2.4: 实现主转换器 src/import_xlsx.js
  - 内容: 加载 Excel -> 分类 sheet -> 动态建表 -> 批量 INSERT -> 写 meta
  - 验证: node src/import_xlsx.js 成功生成 data/processed/market.db (48.75 MB, 73,812 rows)

## Phase 3: 验证与文档

- [x] Task 3.1: 行数对照验证
  - 内容: tests/verify.js 对比每个 sheet 的 Excel 行数 vs DB count(*)
  - 验证: 54 张表全部 PASS, 0 mismatch

- [x] Task 3.2: 抽样数据对比
  - 内容: 随机抽 3 个 sheet (202206, 2025.6, 2026.7), 前 3 + 后 3 行逐列对比
  - 验证: 3 个抽样 sheet 全部 PASS, 数据完全一致

- [x] Task 3.3: meta 表验证
  - 内容: 检查 meta 表 schema + 记录数
  - 验证: PASS (id/imported_at/source_file/schema_version 全部正确)

- [x] Task 3.4: 编写 docs/REVIEWS.md
  - 内容: REVIEWS.md 已建立 (默认空, 留给后续审查)
  - 验证: 文件存在

## Phase R: 后续重构 (按需追加)

- [x] Task R1.1: 修复 TOP 表 header 错位 (headerRow 1 -> 0) [2026-08-27]
  - 内容: import_xlsx.js processTopSheet 硬编码 headerRow=1, 实际 header 在 row 0, 导致列名错位为数字
  - 验证: full_audit.js 全表 4,441,989 cells 0 mismatch

- [x] Task R1.2: 保留空 header 列的数据 (_col_N 命名) [2026-08-27]
  - 内容: TOP销量(倍率) 等表的空 header 列含数据, 原本被 if(!headers[i]) continue 丢弃
  - 验证: full_audit.js 全表 0 mismatch, TOP销量(倍率) 3697 cell 找回

- [x] Task R1.3: tests/verify.js 修复 TOP header 检测 (1 -> 0) [2026-08-27]
  - 内容: detectHeaderRow 对 TOP 表返回 1, 与 import 不一致
  - 验证: verify.js 54 表全部 PASS

- [x] Task R1.4: tests/verify.js 修复抽样对比逻辑 (前3+后3 vs 前6) [2026-08-27]
  - 内容: getDbSample 取前 6, 但 Excel 取前 3+后 3, 范围不一致
  - 验证: 抽样对比 PASS

- [x] Task R1.5: 编写 full_audit.js 全表逐 cell 校验 [2026-08-27]
  - 内容: 覆盖所有 54 张表的逐 cell 比对
  - 历史验证记录: 4,441,989 cells, 0 mismatch
  - 当前状态更正: 仓库中未找到该脚本，现阶段不可复验；见 R3.3

## Round 2 自审 (2026-08-27): 源数据质量审查

> 经过 full_audit.js 全表逐 cell 校验确认 0 mismatch 后, 继续审查源数据本身的质量问题.  
> > 详见 `docs/REVIEWS.md` Round 2 段落.

- [x] Task R2.1: 关键列类型一致性审查 [2026-08-27]
  - 内容: 上架时间 / 子体销量 / 子体销售额 / Coupon / LQS 等列在不同 sheet 类型不一致
  - 验证: 9 项源数据质量问题已记录到 REVIEWS.md
  - 决策: 保留源数据原貌 (lossless 原则), 仅文档化

- [x] Task R2.2: 上架时间 Excel 序列号检测 [2026-08-27]
  - 内容: 验证 202206 是 cell.t=s, 202406 是 cell.t=n, 确认是源数据差异而非脚本 bug
  - 验证: verify_source.js 输出确认

- [x] Task R2.3: 重复列名 (重量 / 体积) 处理确认 [2026-08-27]
  - 内容: 202206-202406 有重量/体积 重复列, 2025.6+ 用新列名 商品重量/商品尺寸
  - 验证: dup_check.js 确认 _2 后缀逻辑生效

- [x] Task R2.4: README 增加跨表查询示例 [2026-08-27]
  - 内容: 4.4 跨月聚合 (含 CAST 处理), 4.5 数据质量注意事项表
  - 验证: test_queries.js 7 个示例全部可执行 (1 个预期失败的列缺失场景)

## 已完成

Phase 1/2/3 已完成；Round 3 复核发现的待修复项如下。

## Round 3 复核 (2026-08-27): 工作表可见性与可复验性

- [x] Task R3.1: 核对工作簿可见性口径
  - 结果: 23 张可见表、32 张隐藏表；隐藏表中 31 张为有效历史月份，`Sheet6` 无有效区域
  - 决策: 按原始“历年 + YoY/MoM”要求保留隐藏历史数据，文档分别披露内部总数与可见表数

- [x] Task R3.2: 修正文档中的“55 张业务表 + 1 张空表”表述
  - 结果: 有效业务表应为 54 张；55 是工作簿内部登记总数，`Sheet6` 只是隐藏遗留表

- [x] Task R3.3: 恢复并运行 `tests/full_audit.js`
  - 原因: TASKS/README/REVIEWS 声称存在全量逐 cell 审计，但当前文件缺失
  - 验收: 54 张有效业务表逐行逐列比对，输出实际 checked cells 与 mismatch 数

- [x] Task R3.4: 修复 `IMPORT_OVERWRITE=false` 保护逻辑
  - 现状: 主 DB 虽不删除，但导入流程仍会 `DROP TABLE IF EXISTS`，会覆盖原表
  - 验收: 目标 DB 存在且配置为 false 时非零退出，DB 哈希不变

- [x] Task R3.5: 改为空表内容判断并记录工作表可见性
  - 现状: `classifySheet` 仅把固定名称 `Sheet6` 视为空表，且 meta 不记录可见/隐藏口径
  - 验收: 依据 `!ref` 判断无有效区域；验证输出明确显示 23 可见、32 隐藏、54 有效导入、1 跳过

- [x] Task R3.6: 独立执行全量逐 cell 复核
  - 结果: 54 张有效业务表，4,441,989 个单元格，结构问题 0，值不一致 0
  - 补充: 发现 70,989 个源 Excel 数值在 SQLite 中以 TEXT 存储；数值等值但存储类型发生变化
  - 限制: 本次通过一次性只读核验命令完成，仍需 R3.3 恢复可复跑脚本

- [x] Task R3.7: 落实或移除 `IMPORT_NUMERIC_TOLERANCE`
  - 现状: `.env.example` 与 SPEC 声明该配置，但 `src/import_xlsx.js` 未读取

- [x] Task R3.8: 修正 `meta.db_size_bytes` 统计时点
  - 现状: meta 记录 51,113,984 bytes，当前 DB 实际为 52,760,576 bytes
  - 原因: WAL 模式下在 checkpoint/close 前读取主 DB 文件大小，不能代表最终文件大小

- [ ] Task R3.9: 核实 2026.04 - 2026.07 源数据统计口径
  - 结果: 全量审计确认异常值与源 Excel 一致，不是字段错位或转换错误
  - 风险: 月销量从 2026.03 的 291,479 跳至 2026.04 的 2,924,738，后续月份维持百万级，不能直接并入趋势结论

## Round 4 改造执行 (2026-08-27)

- [x] Task R4.1: 建立本地 Git 改造前基线
  - 提交: `ca2c66f chore: establish pre-optimization baseline`
  - 远程: GitHub CLI 已恢复登录；远程私有仓库推送等待对具体文件和目的地的明确确认

- [x] Task R4.2: 转换器 schema 1.1.0 与 `sheet_catalog`
  - 验证: visible=23, hidden=32, effective=54, skipped=1

- [x] Task R4.3: 全量审计与覆盖保护自动化
  - 验证: 4,441,989 cells, structural issues=0, value mismatches=0；overwrite=false DB 哈希不变

- [x] Task R4.4: 建立可复现分析生成器
  - 输出: 结构化 JSON、优化版 Markdown、优化版独立 HTML
  - 口径: 小类BSR前100、五档分层、月度MoM/YoY、年度同周期YoY、双均价、异常月份附录、GENIMO建议

- [ ] Task R4.5: 业务确认 2026.03-07 数据口径后决定正式纳入或继续隔离
