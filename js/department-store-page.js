// ============================================================
// 百貨店セクター分析レポート - データ定義 (ドラフト版)
// ============================================================
// ※ 本ファイルの数値はすべてレポート構成確認用の仮置き値。
//    SPEEDA・各社IR月次開示・日本百貨店協会統計の実データ受領後に差し替える。
// 月次KPI: 店頭(既存店)売上高前年比 / 免税売上高前年比 (各社月次IR)
// 業界平均: 日本百貨店協会「全国百貨店売上高」前年同月比
// ============================================================

var companies = [
  { code:'3099', name:'三越伊勢丹ホールディングス', segment:'URBAN', fiscalYear:'2026/3期(仮)', brands:['伊勢丹新宿','三越日本橋','三越銀座'],
    stockPrice:2300, marketCap:880000, revenue:570000, opProfit:65000, netProfit:45000, opMargin:11.4, netMargin:7.9,
    per:18.5, pbr:1.4, roe:8.2, dividendYield:1.4, stores:15,
    monthlySSS:[6.2,4.8,-3.5,-2.1,1.8,2.5,3.9,5.2,3.1,2.4,4.0,5.5],
    monthlyDutyFree:[-8.0,-18.5,-32.0,-28.5,-15.2,-9.8,-3.5,2.0,6.5,9.0,14.5,18.0] },
  { code:'8233', name:'高島屋', segment:'URBAN', fiscalYear:'2026/2期(仮)', brands:['高島屋日本橋','高島屋大阪','高島屋横浜'],
    stockPrice:1250, marketCap:390000, revenue:500000, opProfit:46000, netProfit:35000, opMargin:9.2, netMargin:7.0,
    per:11.5, pbr:0.9, roe:8.6, dividendYield:2.0, stores:17,
    monthlySSS:[4.8,3.5,-4.2,-3.0,0.9,1.8,2.8,4.1,2.5,1.9,3.2,4.6],
    monthlyDutyFree:[-10.5,-20.0,-35.5,-30.0,-18.0,-12.5,-5.0,0.5,5.0,7.5,12.0,15.5] },
  { code:'3086', name:'J.フロント リテイリング', segment:'URBAN', fiscalYear:'2026/2期(仮)', brands:['大丸','松坂屋','パルコ'],
    stockPrice:1800, marketCap:470000, revenue:420000, opProfit:45000, netProfit:32000, opMargin:10.7, netMargin:7.6,
    per:14.0, pbr:1.1, roe:8.9, dividendYield:2.2, stores:20,
    monthlySSS:[5.5,4.0,-5.0,-3.8,1.2,2.0,3.5,4.8,2.8,2.1,3.8,5.0],
    monthlyDutyFree:[-6.5,-16.0,-30.0,-26.5,-13.0,-8.0,-2.0,3.5,8.0,10.5,16.0,20.0] },
  { code:'8242', name:'エイチ・ツー・オー リテイリング', segment:'RAIL', fiscalYear:'2026/3期(仮)', brands:['阪急うめだ本店','阪神梅田本店','イズミヤ'],
    stockPrice:2000, marketCap:250000, revenue:550000, opProfit:30000, netProfit:20000, opMargin:5.5, netMargin:3.6,
    per:13.0, pbr:1.0, roe:7.5, dividendYield:1.5, stores:80,
    monthlySSS:[7.0,5.5,-2.8,-1.5,2.2,3.0,4.5,6.0,3.5,2.8,4.5,6.2],
    monthlyDutyFree:[-4.0,-12.5,-25.0,-22.0,-10.0,-5.5,1.0,6.0,11.0,14.0,19.5,24.0] },
  { code:'8237', name:'松屋', segment:'URBAN', fiscalYear:'2026/2期(仮)', brands:['松屋銀座','松屋浅草'],
    stockPrice:1050, marketCap:56000, revenue:45000, opProfit:4200, netProfit:2900, opMargin:9.3, netMargin:6.4,
    per:19.0, pbr:1.6, roe:8.8, dividendYield:1.0, stores:2,
    monthlySSS:[3.0,1.5,-8.5,-6.8,-1.5,0.5,1.8,3.2,1.0,0.4,2.0,3.5],
    monthlyDutyFree:[-15.0,-28.0,-42.0,-38.5,-24.0,-18.0,-10.5,-4.0,1.5,4.0,9.0,12.5] },
  { code:'8244', name:'近鉄百貨店', segment:'RAIL', fiscalYear:'2026/2期(仮)', brands:['あべのハルカス近鉄本店'],
    stockPrice:2300, marketCap:93000, revenue:120000, opProfit:4500, netProfit:3000, opMargin:3.8, netMargin:2.5,
    per:16.5, pbr:1.3, roe:7.8, dividendYield:1.1, stores:11,
    monthlySSS:[4.2,3.0,-3.2,-2.0,1.0,1.6,2.5,3.8,2.0,1.5,2.8,4.0],
    monthlyDutyFree:[-9.0,-19.5,-33.0,-29.0,-16.5,-11.0,-4.5,1.0,5.5,8.0,13.0,16.5] },
  { code:'8260', name:'井筒屋', segment:'LOCAL', fiscalYear:'2026/2期(仮)', brands:['井筒屋小倉本店'],
    stockPrice:320, marketCap:8500, revenue:60000, opProfit:1800, netProfit:1200, opMargin:3.0, netMargin:2.0,
    per:9.5, pbr:0.7, roe:7.0, dividendYield:1.6, stores:3,
    monthlySSS:null, monthlyDutyFree:null },
  { code:'8252', name:'丸井グループ', segment:'OTHER', fiscalYear:'2026/3期(仮)', brands:['マルイ','モディ','エポスカード'],
    stockPrice:2500, marketCap:490000, revenue:230000, opProfit:44000, netProfit:28000, opMargin:19.1, netMargin:12.2,
    per:17.0, pbr:1.9, roe:10.5, dividendYield:2.4, stores:22,
    monthlySSS:null, monthlyDutyFree:null },
];

window.SECTOR_CONFIG = {
  slug: 'department-store',
  sectorName: '百貨店',
  storesLabel: '店舗数',
  months: ['25/4','25/5','25/6','25/7','25/8','25/9','25/10','25/11','25/12','26/1','26/2','26/3'],
  segments: { URBAN:'都市型大手', RAIL:'電鉄系', LOCAL:'地方百貨店', OTHER:'SC・フィンテック型' },
  segColors: { URBAN:'#1a2d4f', RAIL:'#5a7fa8', LOCAL:'#c8946e', OTHER:'#9b8b6e' },
  heatmapPosThreshold: 3, // 前年比+3%以上を緑
  heatmapNegThreshold: 0, // 前年比マイナスを赤
  monthlySeries: [
    { field:'monthlySSS', label:'店頭(既存店)売上高 前年同月比 (%)', sub:'各社IR月次売上速報 / 破線=日本百貨店協会 全国売上前年比 (仮置き値)',
      industryAvg:[3.0,2.1,-2.5,-1.8,0.9,1.5,2.2,3.0,1.8,1.2,2.5,3.1], industryAvgLabel:'全国百貨店平均' },
    { field:'monthlyDutyFree', label:'免税売上高 前年同月比 (%)', sub:'インバウンド需要の代理指標 / 各社IR月次開示 (仮置き値)' },
  ],
  execCommentary: `
    <strong>セクター概況 (ドラフト):</strong> 都市型大手・電鉄系・地方百貨店の主要上場${'8'}社をカバー。
    2024年に免税売上を中心とするインバウンド消費が過去最高を更新した後、2025年は<strong>円高転換・中国景気減速による免税売上の反動減</strong>が最大のテーマ。
    店頭売上は富裕層の外商・ラグジュアリー消費が下支えする一方、免税比率の高い都心旗艦店(伊勢丹新宿・阪急うめだ等)ほど月次のボラティリティが高い。<br><br>
    <strong>注目ポイント(実データ反映後に検証):</strong>
    (1) <strong>免税売上のベース効果一巡</strong>のタイミングと店頭売上のオーガニック成長率、
    (2) <strong>不動産含み益</strong>(都心一等地の店舗資産)とPBR評価のギャップ — アクティビスト観点での注目領域、
    (3) 地方・郊外店の<strong>構造的閉店</strong>と収益性改善の進捗、
    (4) 外商・富裕層ビジネスの<strong>ストック化</strong>(高島屋ファイナンシャル・パートナーズ等)。`,
  monthlyCommentary: `
    <strong>月次KPIの設計 (ドラフト):</strong> 百貨店は各社が<strong>月次店頭売上高(既存店ベース)と免税売上高</strong>を開示しており、
    インバウンド動向と国内消費を分解して追跡できる。業界平均(破線)は日本百貨店協会「全国百貨店売上高概況」の前年同月比を使用予定。
    井筒屋・丸井Gは月次開示形式が異なるため、実データ反映時に開示粒度を確認の上で追加する。`,
  monthlyMethodology: `
    <strong>データ方法論 (ドラフト / 実データ反映時に確定):</strong><br>
    ・<strong>対象期間:</strong> 2025年4月～2026年3月 (暦月ベース、決算期の異なる企業も暦月で統一)<br>
    ・<strong>店頭売上高:</strong> 各社IR月次売上速報の既存店前年同月比。百貨店事業のみ(SC・金融事業は含まず)<br>
    ・<strong>免税売上高:</strong> 各社月次開示の免税売上前年比。非開示企業は日本百貨店協会の全国免税売上で代替<br>
    ・<strong>業界平均:</strong> 日本百貨店協会「全国百貨店売上高概況」前年同月比(店舗数調整後)<br>
    ・<strong>更新頻度:</strong> 各社月次速報は翌月初旬～中旬に開示。毎月中旬にデータ更新を実施予定`,
  macroCommentary: `
    <strong>外部環境 (ドラフト):</strong> 全国百貨店売上高はコロナ禍の2020年に4兆円台前半まで落ち込んだ後、
    インバウンド回復と富裕層消費により<strong>約6兆円規模まで回復</strong>。訪日外客数は2024年に3,600万人超と過去最高を更新し、
    免税売上高は業界全体で年間5,000億円規模。一方で地方百貨店の閉店が続き、<strong>店舗数は構造的な減少トレンド</strong>にある。
    ※以下のチャートはいずれも仮置き値。実データは日本百貨店協会・JNTO・観光庁統計から取得予定。`,
  macroCharts: [
    { title:'全国百貨店売上高の推移', sub:'出所(予定): 日本百貨店協会 / 単位: 兆円 (仮置き値)', type:'bar', yTitle:'兆円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'全国百貨店売上高', data:[6.2,6.0,5.9,5.9,5.8,4.2,4.4,5.0,5.5,5.9,5.8] }] },
    { title:'訪日外客数の推移', sub:'出所(予定): JNTO / 単位: 万人 (仮置き値)', type:'line', yTitle:'万人',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'訪日外客数', data:[1974,2404,2869,3119,3188,412,25,383,2507,3687,3900] }] },
    { title:'百貨店免税売上高の推移', sub:'出所(予定): 日本百貨店協会 / 単位: 億円 (仮置き値)', type:'bar', yTitle:'億円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'免税売上高', data:[1930,1810,2700,3400,3450,300,100,500,3100,6000,5200] }] },
    { title:'商品カテゴリ別 売上前年比 (直近年)', sub:'出所(予定): 日本百貨店協会 商品別売上 (仮置き値)', type:'bar',
      labels:['美術・宝飾・貴金属','ラグジュアリー衣料','化粧品','婦人服・洋品','紳士服・洋品','食料品','家庭用品'],
      series:[{ label:'前年比(%)', data:[8.5,6.0,2.5,-1.0,-0.5,1.8,-2.5] }] },
  ],
};
