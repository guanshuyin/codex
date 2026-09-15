---
name: amazon-macro-market-research
description: 调研用户指定的 Amazon 末级类目，获取并核验滚动24个月市场、搜索、价格、产品与竞争数据，进行对抗性入局评估并生成可追溯的Excel报告。适用于类目级宏观研究；不适用于ASIN图片识别、微观选品、利润核算或补货预测。
---

# 宏观市场调研

为用户主动提供的 Amazon 末级类目生成滚动 24 个月宏观市场调研与入局评估。默认美国站、Excel 输出；用户明确指定其他站点或正式汇报格式时再调整。

## 开始前

1. 确认用户提供了末级类目名称、完整路径或节点 ID。
2. 优先检查卖家精灵、Softtime、SIF 的可用 MCP 接口。按指标质量选择来源，不把不同平台数据机械平均。
3. 使用类目节点接口验证名称、完整路径、节点 ID、站点和末级属性。无法唯一匹配时，列出候选并停止，不自行选择近似节点。
4. 识别最近完整结算月，并构造包含该月的连续 24 个月区间。不得使用未完结月份或缺月序列。

读取 [references/research-method.md](references/research-method.md) 执行数据获取、对抗性分析和入局判定。

## 生成报告

- 将各接口结果整理为 [references/input-schema.md](references/input-schema.md) 规定的标准化 JSON。
- 创建或修改 Excel 时使用当前环境的 Spreadsheets skill 和 `@oai/artifact-tool`。
- 可调用 [scripts/build_leaf_market_research_report.mjs](scripts/build_leaf_market_research_report.mjs) 生成标准工作簿；脚本参数为标准化 JSON 和输出目录。
- 报告必须保留数值底表、数据来源、样本量和证据限制。每张图表下同时写明事实、支持入局解释、反对入局解释与证据限制。
- 未提供成本、费用、利润和供应链数据时，结论最高只能为“有条件入局”，不得声称具体产品可以盈利。

## 质量要求

读取 [references/report-spec.md](references/report-spec.md) 完成工作表结构、图表和验收检查。

特别注意：

- 关键结论至少需要两个指标支持；能跨源验证时优先跨源验证。
- 数据源不可用、字段缺失、节点文字漂移、样本污染和接口口径冲突必须公开披露。
- 单月爆发、单个新品或单一关键词不得单独决定入局结论。
- 公式重算后扫描错误，并逐张工作表渲染检查空白图、错误比例轴、文字截断和布局拥挤。

## 边界

本技能评估市场层面的可进入性，不执行以下工作：

- 根据图片或 Listing 识别产品类型；
- 批量筛选、分类或下载 ASIN；
- 具体 SKU 的利润、库存、采购和补货计算；
- 未经用户授权向外部系统发布、提交或共享报告。
