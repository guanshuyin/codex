# 标准化输入 JSON

工作簿生成脚本接收一个 JSON 文件。字段名稳定，类目和指标值均由本次研究动态提供，不得在脚本中写死具体类目。

## 顶层结构

```json
{
  "meta": {},
  "sources": [],
  "history": [],
  "structures": {"price": [], "rating": [], "reviews": [], "age": [], "brands": []},
  "keywordCurrent": [],
  "keywordTrends": [],
  "qualityIssues": [],
  "recommendations": {}
}
```

## `meta`

必填：`categoryName`、`categoryPath`、`nodeId`、`nodeIdPath`、`marketplace`、`site`、`currency`、`baselineMonth`、`periodStart`、`periodEnd`、`sampleSize`、`generatedAt`、`leafMatch`、`primaryKeyword`。月份格式为 `yyyyMM`。

## `history`

必须恰好 24 条并按月升序。每条至少包含：`month`、`totalProducts`、`totalUnits`、`totalRevenue`、`avgPrice`、`avgUnits`、`top10ProductCrn`、`top10BrandCrn`、`top10SellerCrn`、`fbaProportion`、`amazonSelfProportion`、`l6NewAvgSales`。

集中度使用 0—1 小数；`fbaProportion` 和 `amazonSelfProportion` 当前脚本按 0—100 百分数读取。

## 结构数组

- `price`、`rating`、`reviews`、`age`：`label`、`products`、`units`、`share`；
- `brands`：`rank`、`brand`、`products`、`units`、`unitShare`、`revenueShare`、`avgPrice`。

份额均使用 0—1 小数。

## 关键词

`keywordCurrent` 每条包含：`keyword`、`searches`、`purchases`、`purchaseRate`、`clickConcentration`、`products`、`adProducts`、`supplyDemand`、`avgPrice`、`ppc`、`titleDensity`、`spr`、`relevancy`。

`keywordTrends` 每项包含 `keyword` 和按月升序的 `data`；`data` 至少含 `month`、`searches`。主关键词趋势应覆盖报告 24 个月。

## `recommendations`

可选类目特定建议：`priceBand`、`differentiation`、`launchTiming`、`inventoryAndAds`、`missingInputs`。这些内容必须来自本次分析，不得从其他类目样例套用。
