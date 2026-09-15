import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root=process.cwd();
const inputPath=path.resolve(process.argv[2]||"");
if(!process.argv[2]) throw new Error("Usage: node build_leaf_market_research_report.mjs <input.json> [outdir]");
const d=JSON.parse(await fs.readFile(inputPath,"utf8"));
const outdir=path.resolve(process.argv[3]||path.join(root,"outputs","01a0235f-b283-7d11-8575-d67b181086c9"));
const outfile=path.join(outdir,`${d.meta.categoryName}_${d.meta.baselineMonth}_24个月宏观市场调研与入局评估.xlsx`);
const previewDir=path.join(outdir,`${d.meta.categoryName}_${d.meta.baselineMonth}_previews`);
const h=d.history;
if(h.length!==24) throw new Error(`趋势月份必须为24，实际${h.length}`);
if(h[0].month!==d.meta.periodStart||h[23].month!==d.meta.periodEnd) throw new Error("趋势区间与元数据不一致");

const n=v=>Number(v||0), sum=a=>a.reduce((x,y)=>x+n(y),0), pct=v=>`${(100*n(v)).toFixed(1)}%`, ml=m=>`${m.slice(0,4)}-${m.slice(4)}`;
const prior=h.slice(0,12), recent=h.slice(12);
const priorUnits=sum(prior.map(x=>x.totalUnits)), recentUnits=sum(recent.map(x=>x.totalUnits));
const priorRev=sum(prior.map(x=>x.totalRevenue)), recentRev=sum(recent.map(x=>x.totalRevenue));
const unitsYoY=recentUnits/priorUnits-1, revYoY=recentRev/priorRev-1;
const peak=h.reduce((a,b)=>n(a.totalUnits)>n(b.totalUnits)?a:b), trough=h.reduce((a,b)=>n(a.totalUnits)<n(b.totalUnits)?a:b);
const last=h[23], prices=d.structures.price, ratings=d.structures.rating, reviews=d.structures.reviews, ages=d.structures.age, brands=d.structures.brands;
const under150=sum(prices.filter(x=>!x.label.startsWith("150")&&!x.label.startsWith("200")&&!x.label.startsWith("250")&&!x.label.startsWith("300")&&!x.label.startsWith("350")).map(x=>x.share));
const review500=reviews.find(x=>x.label==="500+")?.share||0, old3=ages.find(x=>x.label==="3年以上")?.share||0;
const top1=brands[0]?.unitShare||0, top10=n(last.top10BrandCrn), product10=n(last.top10ProductCrn), seller10=n(last.top10SellerCrn);
const newWeak=n(last.l6NewAvgSales)<n(last.avgUnits);
const coreKw=d.keywordCurrent.find(x=>x.keyword===d.meta.primaryKeyword)||d.keywordCurrent[0];
const kwTrend=d.keywordTrends.find(x=>x.keyword===coreKw.keyword)?.data||[];
const kwPrior=sum(kwTrend.slice(0,12).map(x=>x.searches)), kwRecent=sum(kwTrend.slice(12).map(x=>x.searches));
const kwYoY=kwPrior?kwRecent/kwPrior-1:null;
const vetoes=[
  {rule:"销量与核心搜索需求同时明显下降",hit:unitsYoY<=-0.10&&kwYoY!==null&&kwYoY<=-0.10,evidence:`销量同比${pct(unitsYoY)}；核心词搜索同比${kwYoY===null?"缺失":pct(kwYoY)}`},
  {rule:"Top10品牌≥80%或Top1品牌≥35%",hit:top10>=0.80||top1>=0.35,evidence:`Top10 ${pct(top10)}；Top1 ${pct(top1)}`},
  {rule:"高评价老品主导且新品明显偏弱",hit:review500>=0.70&&newWeak,evidence:`500+评价销量份额${pct(review500)}；新品均销${n(last.l6NewAvgSales)} vs 全体${n(last.avgUnits)}`},
  {rule:"主流价格下移且PPC、商品数上升",hit:false,evidence:"缺少连续24个月价格带/PPC/商品数同口径序列，证据不足"},
  {rule:"机会依赖未验证差异化或成本假设",hit:true,evidence:"尚未提供成本、费用、供应链与产品方案"}
];
let verdict="有条件入局";
if(vetoes.slice(0,4).filter(x=>x.hit).length>=2) verdict="不建议入局";
// 缺少成本/利润/供应链时，最高只能有条件入局。

const wb=Workbook.create();
const names=["结论总览","数据源与质量检查","市场容量与季节","搜索需求与流量","价格与产品画像","品牌与竞争","支持入局证据","反对入局证据","最终入局判定","原始数据"];
const sh=Object.fromEntries(names.map(x=>[x,wb.worksheets.add(x)]));
const C={ink:"#243746",muted:"#667985",blue:"#2D6F95",light:"#EAF2F6",line:"#C9D8E1",green:"#DDEFE6",red:"#F6E3E1",amber:"#FFF0CF",white:"#FFFFFF"};
const font="Microsoft YaHei";
function base(s){s.showGridLines=false;s.getRange("A:Q").format.font={name:font,size:10,color:C.ink};s.getRange("1:1").format.rowHeight=12;}
function title(s,t,sub){s.mergeCells("A2:Q2");s.getRange("A2").values=[[t]];s.getRange("A2:Q2").format={font:{name:font,size:18,bold:true,color:C.ink},verticalAlignment:"center",borders:{bottom:{style:"medium",color:C.blue}}};s.getRange("2:2").format.rowHeight=32;s.mergeCells("A3:Q3");s.getRange("A3").values=[[sub]];s.getRange("A3:Q3").format={font:{name:font,size:9,color:C.muted},wrapText:true};}
function section(s,row,text){s.mergeCells(`A${row}:Q${row}`);s.getRange(`A${row}`).values=[[text]];s.getRange(`A${row}:Q${row}`).format={fill:C.light,font:{name:font,bold:true,color:C.blue},borders:{bottom:{style:"thin",color:C.line}}};}
function head(r){r.format={fill:C.blue,font:{name:font,bold:true,color:C.white},wrapText:true,horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:C.line}};}
function body(r){r.format={font:{name:font,size:9,color:C.ink},wrapText:true,verticalAlignment:"top",borders:{preset:"all",style:"thin",color:C.line}};}
function block(s,range,label,text,fill=C.light){s.mergeCells(range);const a=range.split(":")[0];s.getRange(a).values=[[`${label}\n${text}`]];s.getRange(range).format={fill,font:{name:font,size:9,color:C.ink},wrapText:true,verticalAlignment:"top",borders:{preset:"outside",style:"thin",color:C.line}};}
function chart(s,type,source,p1,p2,t,yfmt="#,#0") {const c=s.charts.add(type,s.getRange(source));c.title=t;c.hasLegend=true;c.setPosition(p1,p2);c.xAxis={axisType:"textAxis",textStyle:{fontSize:8}};c.yAxis={numberFormatCode:yfmt,textStyle:{fontSize:8}};c.titleTextStyle.fontSize=11;return c;}
for(const s of Object.values(sh)) base(s);
const sub=`${d.meta.site}｜${d.meta.categoryPath}｜节点 ${d.meta.nodeId}｜${ml(d.meta.periodStart)} 至 ${ml(d.meta.periodEnd)}｜主样本 Top${d.meta.sampleSize}`;

// 原始数据
let s=sh["原始数据"];title(s,"原始数据",sub);
s.getRange("A5:O5").values=[["月份","销量","销售额($)","平均Listing价($)","隐含成交价($)","环比销量","同比销量","同比销售额","Top10商品","Top10品牌","Top10卖家","FBA","Amazon自营","商品数","来源"]];head(s.getRange("A5:O5"));
const hist=h.map((x,i)=>[ml(x.month),n(x.totalUnits),n(x.totalRevenue),n(x.avgPrice),n(x.totalRevenue)/n(x.totalUnits),i?n(x.totalUnits)/n(h[i-1].totalUnits)-1:null,i>=12?n(x.totalUnits)/n(h[i-12].totalUnits)-1:null,i>=12?n(x.totalRevenue)/n(h[i-12].totalRevenue)-1:null,n(x.top10ProductCrn),n(x.top10BrandCrn),n(x.top10SellerCrn),n(x.fbaProportion)/100,n(x.amazonSelfProportion)/100,n(x.totalProducts),"SellerSprite market_research"]);
s.getRange("A6:O29").values=hist;body(s.getRange("A6:O29"));s.getRange("B6:B29").format.numberFormat="#,#0";s.getRange("C6:E29").format.numberFormat='"$"#,#0.00';s.getRange("F6:M29").format.numberFormat="0.0%";
s.getRange("Q5:W5").values=[["关键词","月份","搜索量","购买量","购买率","同比","来源"]];head(s.getRange("Q5:W5"));
const kwRaw=[];for(const k of d.keywordTrends)for(const x of k.data)kwRaw.push([k.keyword,ml(x.month),x.searches,x.purchases||null,x.purchaseRate||null,x.yearlyGrowth||null,"SellerSprite keyword_research_trends"]);
s.getRange(`Q6:W${5+kwRaw.length}`).values=kwRaw;body(s.getRange(`Q6:W${5+kwRaw.length}`));s.getRange(`S6:S${5+kwRaw.length}`).format.numberFormat="#,#0";s.getRange(`U6:V${5+kwRaw.length}`).format.numberFormat="0.0%";s.freezePanes.freezeRows(5);

// 总览
s=sh["结论总览"];title(s,`${d.meta.categoryName}｜24个月宏观市场调研与入局评估`,sub);
section(s,5,"当前判断");s.mergeCells("A6:D8");s.getRange("A6").values=[[verdict]];s.getRange("A6:D8").format={fill:verdict==="不建议入局"?C.red:C.amber,font:{name:font,size:20,bold:true,color:C.ink},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"outside",style:"medium",color:C.blue}};
s.getRange("F6:Q8").values=[[`市场销量：最近12个月 ${recentUnits.toLocaleString()} 件，同比 ${pct(unitsYoY)}\n核心词需求：最近12个月搜索量同比 ${kwYoY===null?"证据不足":pct(kwYoY)}\n集中度：Top10品牌 ${pct(top10)}，Top10商品 ${pct(product10)}，Top10卖家 ${pct(seller10)}\n结论约束：未提供成本、利润与供应链数据，因此最高只能给出“有条件入局”。`]];s.mergeCells("F6:Q8");s.getRange("F6:Q8").format={fill:C.light,font:{name:font,size:10},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
section(s,10,"快速事实");
const facts=[["基准月",ml(d.meta.baselineMonth)],["基准月销量",last.totalUnits],["基准月销售额",last.totalRevenue],["平均Listing价格",last.avgPrice],["24个月峰值",`${ml(peak.month)} / ${n(peak.totalUnits).toLocaleString()}件`],["24个月低点",`${ml(trough.month)} / ${n(trough.totalUnits).toLocaleString()}件`],["主流价格结构",`$150以下占${pct(under150)}销量`],["评分结构",`4.3-4.5占${pct(ratings.find(x=>x.label==="4.3-4.5")?.share)}销量`],["评价壁垒",`500+评价占${pct(review500)}销量`],["老品结构",`3年以上占${pct(old3)}销量`]];
s.getRange("A11:D20").values=facts.map(x=>[x[0],x[1],null,null]);for(let r=11;r<=20;r++)s.mergeCells(`B${r}:D${r}`);body(s.getRange("A11:D20"));s.getRange("A11:A20").format={fill:C.light,font:{name:font,bold:true,color:C.blue}};s.getRange("B13:B14").format.numberFormat='"$"#,#0.00';
block(s,"F11:Q14","支持入局解释",`市场仍有显著绝对规模；Top10品牌${pct(top10)}未触发80%垄断阈值，Top1品牌${pct(top1)}亦未触发35%阈值；半年内产品贡献${pct(sum(ages.slice(0,3).map(x=>x.share)))}销量，说明新品并非完全没有窗口。`,C.green);
block(s,"F15:Q18","反对入局解释",`核心词搜索同比${kwYoY===null?"无法验证":pct(kwYoY)}；500+评价商品控制${pct(review500)}销量；Top10卖家集中度${pct(seller10)}，评价冷启动、履约和流量成本均为实质门槛。`,C.red);
block(s,"F19:Q22","证据限制",`类目存在误分类商品；SellerSprite两个接口的商品数相差约12.3%；Softtime与SIF未接入；缺少成本、费用、供应链和具体产品方案。`);
section(s,24,"阶段总结");block(s,"A25:Q28","不带倾向性的总结",`该类目是规模可观、季节明显、品牌中度集中、评价壁垒较高的成熟市场。需求端和类目销量端并非完全同向，必须以具体产品利润与差异化验证后再决策。`);

// 数据源与质量
s=sh["数据源与质量检查"];title(s,"数据源与质量检查",sub);section(s,5,"数据源参与情况");s.getRange("A6:E6").values=[["数据源","状态","使用范围","质量/限制","检查日期"]];head(s.getRange("A6:E6"));s.getRange(`A7:E${6+d.sources.length}`).values=d.sources.map(x=>[x.source,x.status,x.scope,x.quality,d.meta.generatedAt]);body(s.getRange(`A7:E${6+d.sources.length}`));s.getRange("A:A").format.columnWidth=20;s.getRange("B:B").format.columnWidth=14;s.getRange("C:C").format.columnWidth=28;s.getRange("D:D").format.columnWidth=60;
section(s,12,"固定质量评分（5分制）");s.getRange("A13:H13").values=[["来源","新鲜度","节点匹配","历史覆盖","样本量","字段完整","口径透明","可追溯"]];head(s.getRange("A13:H13"));s.getRange("A14:H16").values=[["SellerSprite MCP",5,5,5,3,4,3,5],["Softtime",null,null,null,null,null,null,null],["SIF",null,null,null,null,null,null,null]];body(s.getRange("A14:H16"));
section(s,19,"已识别的数据风险");s.getRange(`A20:B${19+d.qualityIssues.length}`).values=d.qualityIssues.map((x,i)=>[i+1,x]);for(let r=20;r<20+d.qualityIssues.length;r++)s.mergeCells(`B${r}:Q${r}`);body(s.getRange(`A20:Q${19+d.qualityIssues.length}`));s.getRange(`20:${19+d.qualityIssues.length}`).format.rowHeight=28;
section(s,28,"阶段总结");block(s,"A29:Q33","不带倾向性的总结","本报告完成了单源内多指标核验，但没有完成跨平台核验。结论可用于市场层面的方向判断，不应被视为精确财务预测。",C.amber);

// 市场容量
s=sh["市场容量与季节"];title(s,"市场容量与季节",sub);section(s,5,"图表1｜月度销量与销售额（数值保留）");s.getRange("A6:E6").values=[["月份","销量","销售额($)","平均Listing价($)","同比销量"]];head(s.getRange("A6:E6"));s.getRange("A7:E30").values=h.map((x,i)=>[ml(x.month),x.totalUnits,x.totalRevenue,x.avgPrice,i>=12?n(x.totalUnits)/n(h[i-12].totalUnits)-1:null]);body(s.getRange("A7:E30"));s.getRange("B7:B30").format.numberFormat="#,#0";s.getRange("C7:D30").format.numberFormat='"$"#,#0.00';s.getRange("E7:E30").format.numberFormat="0.0%";chart(s,"line","A6:B30","G6","Q20","24个月销量趋势","#,#0");
block(s,"G22:Q25","图表事实",`低点${ml(trough.month)}（${n(trough.totalUnits).toLocaleString()}件），峰值${ml(peak.month)}（${n(peak.totalUnits).toLocaleString()}件）。最近12个月销量同比${pct(unitsYoY)}，销售额同比${pct(revYoY)}。`);
block(s,"G26:Q29","支持入局解释","3月通常进入上行段，5—7月形成主要销售窗口；规模足以支持细分场景产品。",C.green);
block(s,"G30:Q33","反对入局解释","峰值后回落快，旺季备货错误会转化为淡季库存风险；年度总量不能替代逐月判断。",C.red);
block(s,"G34:Q37","证据限制","类目Top100统计可能混入附件或误分类商品；销量为第三方估算值。",C.amber);
section(s,40,"最近12个月 vs 前12个月逐月对齐");s.getRange("A41:F41").values=[["月份","前期销量","近期销量","销量同比","前期销售额","近期销售额"]];head(s.getRange("A41:F41"));s.getRange("A42:F53").values=prior.map((x,i)=>[`${Number(x.month.slice(4))}月`,x.totalUnits,recent[i].totalUnits,n(recent[i].totalUnits)/n(x.totalUnits)-1,x.totalRevenue,recent[i].totalRevenue]);body(s.getRange("A42:F53"));s.getRange("B42:C53").format.numberFormat="#,#0";s.getRange("D42:D53").format.numberFormat="0.0%";s.getRange("E42:F53").format.numberFormat='"$"#,#0';section(s,56,"阶段总结");block(s,"A57:Q61","不带倾向性的总结",`旺季启动点集中在3—5月，峰值多在6—7月，7—8月进入下降拐点，11月至次年2月偏淡。具体年份仍会受促销、天气和样本变化影响。`);

// 搜索
s=sh["搜索需求与流量"];title(s,"搜索需求与流量",sub);section(s,5,"图表2｜核心词24个月搜索趋势");s.getRange("A6:C6").values=[["月份",d.keywordTrends[0]?.keyword||"关键词1",d.keywordTrends[1]?.keyword||"关键词2"]];head(s.getRange("A6:C6"));s.getRange("A7:C30").values=monthsFrom(d.keywordTrends);body(s.getRange("A7:C30"));s.getRange("B7:C30").format.numberFormat="#,#0";chart(s,"line","A6:C30","E6","Q20","核心词搜索量趋势","#,#0");
block(s,"E22:Q25","图表事实",`${coreKw.keyword}最近12个月合计搜索量同比${kwYoY===null?"无法计算":pct(kwYoY)}；2026-08当月搜索量${n(coreKw.searches).toLocaleString()}。`);
block(s,"E26:Q29","支持入局解释","搜索需求保持明显季节性，旺季仍存在稳定用户意图；相近词可用于拆分场景和搭建方式。",C.green);
block(s,"E30:Q33","反对入局解释","核心类目词同比下滑，不能把季节性回升误判为长期增长；购买率较低，流量不等于成交。",C.red);
block(s,"E34:Q37","证据限制","关键词工具的商品数、供需比在相近词之间口径异常，不将其作为单独否决项。",C.amber);
section(s,40,"基准月关键词指标");s.getRange("A41:M41").values=[["关键词","搜索量","购买量","购买率","点击集中度","商品数","广告商品","供需比","均价","PPC","标题密度","SPR","相关度"]];head(s.getRange("A41:M41"));s.getRange(`A42:M${41+d.keywordCurrent.length}`).values=d.keywordCurrent.map(x=>[x.keyword,x.searches,x.purchases,x.purchaseRate,x.clickConcentration,x.products,x.adProducts,x.supplyDemand,x.avgPrice,x.ppc,x.titleDensity,x.spr,x.relevancy]);body(s.getRange(`A42:M${41+d.keywordCurrent.length}`));s.getRange(`B42:C${41+d.keywordCurrent.length}`).format.numberFormat="#,#0";s.getRange(`D42:E${41+d.keywordCurrent.length}`).format.numberFormat="0.0%";s.getRange(`I42:J${41+d.keywordCurrent.length}`).format.numberFormat='"$"0.00';section(s,49,"阶段总结");block(s,"A50:Q54","不带倾向性的总结","搜索需求显示明显季节波动和同比压力。PPC、点击集中度和购买率应结合具体产品词重新验证，不能仅凭大词决定进入。");

// 价格画像
s=sh["价格与产品画像"];title(s,"价格与产品画像",sub);section(s,5,"图表3｜价格带销量份额");writeDist(s,6,prices,"价格带",x=>x.revenue);s.getRange("S6:T14").values=[["价格带","销量份额"],...prices.map(x=>[x.label,x.share])];chart(s,"bar","S6:T14","F6","Q19","价格带销量份额","0%");blocks(s,21,[`$150以下占${pct(under150)}销量；最大价格带为${prices.reduce((a,b)=>a.share>b.share?a:b).label}。`,`低于$150存在最大需求池，细分功能仍可形成多价位。`,`低价带体量大也意味着价格战和广告效率压力。`,`价格带只有基准月快照，无法直接证明价格迁移。`]);
section(s,39,"图表4｜评分、评价数与上架时长");s.getRange("A40:D40").values=[["评分区间","商品数","销量","销量份额"]];head(s.getRange("A40:D40"));s.getRange(`A41:D${40+ratings.length}`).values=ratings.map(x=>[x.label,x.products,x.units,x.share]);body(s.getRange(`A41:D${40+ratings.length}`));s.getRange(`D41:D${40+ratings.length}`).format.numberFormat="0.0%";s.getRange(`S40:T${40+ratings.length}`).values=[["评分区间","销量份额"],...ratings.map(x=>[x.label,x.share])];chart(s,"bar",`S40:T${40+ratings.length}`,"F40","Q53","评分区间销量份额","0%");
section(s,56,"评价数与上架时长数值");s.getRange("A57:D57").values=[["评价数区间","商品数","销量","销量份额"]];head(s.getRange("A57:D57"));s.getRange(`A58:D${57+reviews.length}`).values=reviews.map(x=>[x.label,x.products,x.units,x.share]);body(s.getRange(`A58:D${57+reviews.length}`));s.getRange(`D58:D${57+reviews.length}`).format.numberFormat="0.0%";s.getRange("F57:I57").values=[["上架时长","商品数","销量","销量份额"]];head(s.getRange("F57:I57"));s.getRange(`F58:I${57+ages.length}`).values=ages.map(x=>[x.label,x.products,x.units,x.share]);body(s.getRange(`F58:I${57+ages.length}`));s.getRange(`I58:I${57+ages.length}`).format.numberFormat="0.0%";section(s,70,"阶段总结");block(s,"A71:Q76","不带倾向性的总结",`评分4.3—4.5贡献${pct(ratings.find(x=>x.label==="4.3-4.5")?.share)}销量；500+评价贡献${pct(review500)}；3年以上商品贡献${pct(old3)}。消费者基础预期是稳定评分、较强社会证明和成熟内容。`);

// 品牌竞争
s=sh["品牌与竞争"];title(s,"品牌与竞争",sub);section(s,5,"图表5｜Top10品牌份额");s.getRange("A6:G6").values=[["排名","品牌","商品数","销量","销量份额","销售额份额","均价"]];head(s.getRange("A6:G6"));s.getRange("A7:G16").values=brands.map(x=>[x.rank,x.brand,x.products,x.units,x.unitShare,x.revenueShare,x.avgPrice]);body(s.getRange("A7:G16"));s.getRange("E7:F16").format.numberFormat="0.0%";s.getRange("G7:G16").format.numberFormat='"$"0.00';s.getRange("S6:U16").values=[["品牌","销量份额","销售额份额"],...brands.map(x=>[x.brand,x.unitShare,x.revenueShare])];chart(s,"bar","S6:U16","I6","Q20","Top10品牌销量/销售额份额","0%");blocks(s,22,[`Top1品牌${pct(top1)}，Top10品牌${pct(top10)}；Top10商品${pct(product10)}；Top10卖家${pct(seller10)}。`,`品牌未达到硬垄断阈值，单品集中度低于品牌集中度，允许多款产品并存。`,`卖家集中度明显高于商品集中度，渠道、履约和广告资源可能集中。`,`品牌表为Top100样本；部分品牌“新品”标记可能受变体或重上架定义影响。`]);section(s,40,"竞争结构快照");s.getRange("A41:C46").values=[["指标","数值","判断"],["Top10商品",product10,"单品分散"],["Top10品牌",top10,"中度集中"],["Top10卖家",seller10,"渠道偏集中"],["FBA占比",n(last.fbaProportion)/100,"履约基础门槛"],["Amazon自营",n(last.amazonSelfProportion)/100,"平台参与但非绝对主导"]];head(s.getRange("A41:C41"));body(s.getRange("A42:C46"));s.getRange("B42:B46").format.numberFormat="0.0%";section(s,49,"阶段总结");block(s,"A50:Q54","不带倾向性的总结","市场并非单品牌或单品绝对垄断，但渠道和评价资产集中。进入者需要避开同质化正面竞争，并具备稳定履约能力。");

// 支持与反对
const supports=[
  ["市场绝对规模",`基准月${n(last.totalUnits).toLocaleString()}件、$${n(last.totalRevenue).toLocaleString()}`,"支持存在可服务需求","中"],
  ["品牌未触发硬垄断",`Top1 ${pct(top1)}；Top10 ${pct(top10)}`,"允许细分品牌进入","中"],
  ["单品相对分散",`Top10商品${pct(product10)}`,"需求由多款产品承接","中"],
  ["新品存在窗口",`半年内商品销量份额${pct(sum(ages.slice(0,3).map(x=>x.share)))}`,"新品并非完全无法起量","中"],
  ["价格层次丰富",`$150以下占${pct(under150)}`,"可按场景/功能构建价格梯度","中"]
];
const antis=[
  ["搜索需求下行",`核心词同比${kwYoY===null?"缺失":pct(kwYoY)}`,"长期需求承压","高"],
  ["评价壁垒",`500+评价商品占${pct(review500)}销量`,"新品转化和广告效率承压","高"],
  ["老品优势",`3年以上商品占${pct(old3)}销量`,"成熟Listing占据稳定位置","中"],
  ["卖家集中",`Top10卖家${pct(seller10)}`,"渠道和运营资源偏集中","中"],
  ["强季节性",`峰值${ml(peak.month)}，低点${ml(trough.month)}`,"预测错误导致库存风险","高"],
  ["数据污染",`存在附件/误分类；商品数口径差12.3%`,"总市场可能被高估或结构失真","高"]
];
evidenceSheet(sh["支持入局证据"],"支持入局证据",supports,C.green);evidenceSheet(sh["反对入局证据"],"反对入局证据",antis,C.red);

// 最终判定
s=sh["最终入局判定"];title(s,"最终入局判定",sub);section(s,5,"三级结论");s.mergeCells("A6:D8");s.getRange("A6").values=[[verdict]];s.getRange("A6:D8").format={fill:C.amber,font:{name:font,size:20,bold:true},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"outside",style:"medium",color:C.blue}};block(s,"F6:Q8","结论边界","这是市场层面的可进入性判断。缺少成本、费用、利润与供应链信息，因此不认定任何具体产品可以盈利。",C.amber);
section(s,10,"重点否决信号检查");s.getRange("A11:D11").values=[["否决规则","是否触发","证据","实际影响"]];head(s.getRange("A11:D11"));s.getRange(`A12:D${11+vetoes.length}`).values=vetoes.map(x=>[x.rule,x.hit?"是":"否/证据不足",x.evidence,x.hit?"压低结论":"未形成硬否决"]);body(s.getRange(`A12:D${11+vetoes.length}`));s.getRange("A:A").format.columnWidth=34;s.getRange("C:C").format.columnWidth=58;s.getRange("D:D").format.columnWidth=22;
section(s,19,"进入条件与行动建议");const actions=[
  ["当前是否适合入局",verdict],["最强支持证据",`品牌/单品未硬垄断，Top10品牌${pct(top10)}、Top10商品${pct(product10)}`],["最强反对证据",`核心搜索需求同比${kwYoY===null?"证据不足":pct(kwYoY)}，且500+评价控制${pct(review500)}销量`],
  ["结论成立条件","必须验证差异化功能、到岸成本、FBA费用、广告获客成本、退货损耗和旺季备货能力"],["终止/重评条件","核心销量与搜索需求连续两个旺季同比下降；PPC持续上升且主流价格下移；样品无法形成可感知差异；贡献利润为负"],
  ["建议价格带",d.recommendations?.priceBand||"根据主流价格带、目标贡献利润和可验证差异化确定，不以类目均价直接定价"],["差异化方向",d.recommendations?.differentiation||"基于消费者痛点、评价缺口与产品结构形成可感知差异，并通过样品和评论需求验证"],
  ["上市月份",d.recommendations?.launchTiming||"围绕需求启动月提前完成备货与内容，小规模验证后再承接旺季"],["备货与广告节奏",d.recommendations?.inventoryAndAds||"首批保守；按转化、退货和交期滚动补货；峰值前放量，下降拐点后降低补货"],["尚需补充数据",d.recommendations?.missingInputs||"到岸成本、尺寸重量、FBA费用、广告预算/CPC、退货率、MOQ、交期、目标毛利、竞品评论痛点"]
];s.getRange("A20:B29").values=actions;for(let r=20;r<=29;r++)s.mergeCells(`B${r}:Q${r}`);body(s.getRange("A20:Q29"));s.getRange("A20:A29").format={fill:C.light,font:{name:font,bold:true,color:C.blue}};section(s,32,"阶段总结");block(s,"A33:Q37","最终结论",`${verdict}。市场有规模且未形成品牌硬垄断，但核心搜索需求下行、评价壁垒和季节库存风险突出。只有当具体产品能够验证差异化、利润和供应链条件时，市场进入才具有可执行性。`);

// Formatting
for(const x of Object.values(sh)){x.getUsedRange().format.font.name=font;x.getRange("A:Q").format.columnWidth=12;x.getRange("A:A").format.columnWidth=Math.max(x.getRange("A:A").format.columnWidth||12,16);x.freezePanes.freezeRows(5);}
sh["结论总览"].getRange("A:Q").format.columnWidth=12;sh["市场容量与季节"].getRange("A:F").format.columnWidth=15;sh["搜索需求与流量"].getRange("A:M").format.columnWidth=14;sh["价格与产品画像"].getRange("A:I").format.columnWidth=15;sh["品牌与竞争"].getRange("A:G").format.columnWidth=15;sh["支持入局证据"].getRange("A:D").format.columnWidth=28;sh["反对入局证据"].getRange("A:D").format.columnWidth=28;sh["原始数据"].getRange("A:W").format.columnWidth=15;
sh["数据源与质量检查"].getRange("A:A").format.columnWidth=20;sh["数据源与质量检查"].getRange("B:B").format.columnWidth=14;sh["数据源与质量检查"].getRange("C:C").format.columnWidth=28;sh["数据源与质量检查"].getRange("D:D").format.columnWidth=60;
sh["最终入局判定"].getRange("A:A").format.columnWidth=32;sh["最终入局判定"].getRange("B:B").format.columnWidth=22;sh["最终入局判定"].getRange("C:C").format.columnWidth=48;sh["最终入局判定"].getRange("D:D").format.columnWidth=22;

wb.recalculate();
await fs.mkdir(outdir,{recursive:true});await fs.mkdir(previewDir,{recursive:true});
const f=await SpreadsheetFile.exportXlsx(wb);await f.save(outfile);
const renderRanges={"结论总览":"A1:Q28","数据源与质量检查":"A1:Q33","市场容量与季节":"A1:Q61","搜索需求与流量":"A1:Q54","价格与产品画像":"A1:Q76","品牌与竞争":"A1:Q54","支持入局证据":"A1:Q28","反对入局证据":"A1:Q29","最终入局判定":"A1:Q37","原始数据":"A1:W35"};
for(const [name,range] of Object.entries(renderRanges)){const p=await wb.render({sheetName:name,range,scale:0.75,format:"png"});await fs.writeFile(path.join(previewDir,`${name}.png`),new Uint8Array(await p.arrayBuffer()));}
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"formula errors"});
const check=await wb.inspect({kind:"table",sheetId:"最终入局判定",range:"A1:D29",include:"values,formulas",tableMaxRows:29,tableMaxCols:4,maxChars:8000});
console.log(JSON.stringify({outfile,previewDir,verdict,unitsYoY,revYoY,kwYoY,errors:errors.ndjson,check:check.ndjson.slice(0,2500)},null,2));

function monthsFrom(trends){const a=trends[0]?.data||[];return a.map((x,i)=>[ml(x.month),x.searches,trends[1]?.data?.[i]?.searches||null]);}
function writeDist(s,start,arr,label){s.getRange(`A${start}:D${start}`).values=[[label,"商品数","销量","销量份额"]];head(s.getRange(`A${start}:D${start}`));s.getRange(`A${start+1}:D${start+arr.length}`).values=arr.map(x=>[x.label,x.products,x.units,x.share]);body(s.getRange(`A${start+1}:D${start+arr.length}`));s.getRange(`D${start+1}:D${start+arr.length}`).format.numberFormat="0.0%";}
function blocks(s,row,texts){block(s,`F${row}:Q${row+3}`,"图表事实",texts[0]);block(s,`F${row+4}:Q${row+7}`,"支持入局解释",texts[1],C.green);block(s,`F${row+8}:Q${row+11}`,"反对入局解释",texts[2],C.red);block(s,`F${row+12}:Q${row+15}`,"证据限制",texts[3],C.amber);}
function evidenceSheet(s,t,rows,fill){title(s,t,sub);section(s,5,"证据清单");s.getRange("A6:D6").values=[["指标","事实证据","对判断的意义","证据强度"]];head(s.getRange("A6:D6"));s.getRange(`A7:D${6+rows.length}`).values=rows;body(s.getRange(`A7:D${6+rows.length}`));s.getRange(`A7:D${6+rows.length}`).format.fill=fill;s.getRange("A:D").format.columnWidth=30;section(s,16,"反面检验与限制");block(s,"A17:Q22","证据限制","每项证据均需与相反指标共同阅读；类目污染、单一数据源与缺少成本数据会降低结论强度。单月爆发、单个新品或单一关键词不会被作为入局依据。",C.amber);section(s,25,"阶段总结");block(s,"A26:Q28","不带倾向性的总结","本页仅陈列单侧证据，最终判断以支持与反对证据的共同权衡及否决规则为准。");}
