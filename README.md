# Codex Skills

这个仓库用于保存个人 Codex Skill。

## 补货分析

`replenishment-analysis` 用于分析亚马逊帐篷产品的历史销量、市场与竞品趋势、库存和在途数据，并生成：

- 标准化的 Sheet1—Sheet5 工作簿；
- 月度需求及预测日均单量；
- SKU 库存充足度分类；
- 采购、发货及库存变化计划；
- 市场趋势、竞品选择和动态权重依据。

Skill 目录：[skills/replenishment-analysis](skills/replenishment-analysis/)

### 调用方式

安装后可在 Codex 中直接调用：

```text
$replenishment-analysis
```

示例：

```text
使用 $replenishment-analysis 分析这份帐篷销量库存表并生成需求预测与补货计划。
```

## 宏观市场调研

`amazon-macro-market-research` 用于调研用户指定的 Amazon 末级类目，获取滚动 24 个月市场与搜索数据，分析市场容量、季节、价格、产品画像和竞争格局，并采用对抗性设计给出三级入局判断。

Skill 目录：[skills/amazon-macro-market-research](skills/amazon-macro-market-research/)

### 调用方式

```text
$amazon-macro-market-research
```

示例：

```text
使用 $amazon-macro-market-research 调研我提供的 Amazon 末级类目，并生成宏观市场与入局评估。
```

## 安装方式

将需要的 Skill 目录复制到本机 Codex 技能目录：

```text
~/.codex/skills/
```

重新打开 Codex 会话后即可识别。

## 数据说明

本仓库只保存 Skill 的规则、参考文档和通用脚本，不包含销量、库存、ASIN 明细、Excel 输出或其他业务数据。
