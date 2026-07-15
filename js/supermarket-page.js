// ============================================================
// スーパーマーケットセクター分析レポート - データ定義 (ドラフト版)
// ============================================================
// ※ 本ファイルの数値はすべてレポート構成確認用の仮置き値。
//    SPEEDA・各社IR月次開示・スーパーマーケット販売統計の実データ受領後に差し替える。
// 月次KPI: 既存店売上高前年比 / 既存店客数前年比 / 既存店客単価前年比
// 業界平均: 全国スーパーマーケット協会等3団体「スーパーマーケット販売統計」既存店前年比
// ============================================================

var companies = [
  { code:'8267', name:'イオン', segment:'GMS', fiscalYear:'2026/2期(仮)', brands:['イオン','マックスバリュ','まいばすけっと'],
    stockPrice:3800, marketCap:3300000, revenue:9900000, opProfit:255000, netProfit:45000, opMargin:2.6, netMargin:0.5,
    per:70.0, pbr:2.8, roe:4.0, dividendYield:1.1, stores:5800,
    monthlySSS:[2.8,3.1,2.5,2.2,2.9,3.4,2.6,2.3,3.0,3.5,2.7,2.4],
    monthlyCust:[0.5,0.8,0.3,0.1,0.6,1.0,0.4,0.2,0.7,1.1,0.5,0.3],
    monthlyBasket:[2.3,2.3,2.2,2.1,2.3,2.4,2.2,2.1,2.3,2.4,2.2,2.1] },
  { code:'3038', name:'神戸物産', segment:'DISC', fiscalYear:'2026/10期(仮)', brands:['業務スーパー'],
    stockPrice:3500, marketCap:960000, revenue:530000, opProfit:34000, netProfit:23000, opMargin:6.4, netMargin:4.3,
    per:26.0, pbr:5.2, roe:20.5, dividendYield:0.7, stores:1080,
    monthlySSS:[5.5,6.0,4.8,4.4,5.6,6.5,5.0,4.6,5.8,6.8,5.2,4.8],
    monthlyCust:[3.0,3.4,2.5,2.2,3.1,3.8,2.6,2.3,3.2,4.0,2.8,2.4],
    monthlyBasket:[2.5,2.6,2.3,2.2,2.5,2.7,2.4,2.3,2.6,2.8,2.4,2.4] },
  { code:'8279', name:'ヤオコー', segment:'FOOD', fiscalYear:'2026/3期(仮)', brands:['ヤオコー','エイヴイ','せんどう'],
    stockPrice:4200, marketCap:440000, revenue:660000, opProfit:31000, netProfit:20000, opMargin:4.7, netMargin:3.0,
    per:16.5, pbr:2.1, roe:13.0, dividendYield:1.2, stores:200,
    monthlySSS:[4.5,4.8,4.0,3.6,4.6,5.2,4.1,3.8,4.8,5.5,4.2,3.9],
    monthlyCust:[1.5,1.8,1.2,0.9,1.6,2.1,1.3,1.0,1.7,2.2,1.4,1.1],
    monthlyBasket:[3.0,3.0,2.8,2.7,3.0,3.1,2.8,2.8,3.1,3.3,2.8,2.8] },
  { code:'8194', name:'ライフコーポレーション', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['ライフ','ビオラル'],
    stockPrice:3300, marketCap:310000, revenue:840000, opProfit:29000, netProfit:19000, opMargin:3.5, netMargin:2.3,
    per:13.0, pbr:1.5, roe:12.0, dividendYield:1.4, stores:310,
    monthlySSS:[3.8,4.1,3.4,3.0,3.9,4.5,3.5,3.2,4.1,4.8,3.6,3.3],
    monthlyCust:[1.0,1.3,0.8,0.5,1.1,1.6,0.9,0.6,1.2,1.7,0.9,0.7],
    monthlyBasket:[2.8,2.8,2.6,2.5,2.8,2.9,2.6,2.6,2.9,3.1,2.7,2.6] },
  { code:'141A', name:'トライアルホールディングス', segment:'DISC', fiscalYear:'2026/6期(仮)', brands:['トライアル','スーパーセンター'],
    stockPrice:2800, marketCap:350000, revenue:800000, opProfit:26000, netProfit:15000, opMargin:3.3, netMargin:1.9,
    per:24.0, pbr:3.1, roe:14.0, dividendYield:0.5, stores:290,
    monthlySSS:[6.5,7.0,5.8,5.4,6.6,7.5,6.0,5.6,6.8,7.8,6.2,5.8],
    monthlyCust:[4.0,4.4,3.4,3.0,4.1,4.8,3.5,3.2,4.2,5.0,3.7,3.3],
    monthlyBasket:[2.5,2.6,2.4,2.4,2.5,2.7,2.5,2.4,2.6,2.8,2.5,2.5] },
  { code:'8273', name:'イズミ', segment:'GMS', fiscalYear:'2026/2期(仮)', brands:['ゆめタウン','ゆめマート'],
    stockPrice:3500, marketCap:250000, revenue:500000, opProfit:26000, netProfit:16000, opMargin:5.2, netMargin:3.2,
    per:11.0, pbr:0.8, roe:8.0, dividendYield:2.6, stores:190,
    monthlySSS:null, monthlyCust:null, monthlyBasket:null },
  { code:'9948', name:'アークス', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['ラルズ','ユニバース','ベルジョイス'],
    stockPrice:2900, marketCap:160000, revenue:640000, opProfit:17500, netProfit:12000, opMargin:2.7, netMargin:1.9,
    per:11.5, pbr:0.9, roe:7.8, dividendYield:2.2, stores:380,
    monthlySSS:[2.5,2.8,2.2,1.9,2.6,3.1,2.3,2.0,2.7,3.2,2.4,2.1],
    monthlyCust:[0.2,0.5,0.0,-0.3,0.3,0.8,0.1,-0.2,0.4,0.9,0.2,-0.1],
    monthlyBasket:[2.3,2.3,2.2,2.2,2.3,2.3,2.2,2.2,2.3,2.3,2.2,2.2] },
  { code:'9956', name:'バローホールディングス', segment:'FOOD', fiscalYear:'2026/3期(仮)', brands:['バロー','タチヤ','中部薬品'],
    stockPrice:2400, marketCap:130000, revenue:830000, opProfit:25000, netProfit:14000, opMargin:3.0, netMargin:1.7,
    per:9.0, pbr:0.7, roe:8.2, dividendYield:2.5, stores:1200,
    monthlySSS:null, monthlyCust:null, monthlyBasket:null },
  { code:'9974', name:'ベルク', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['ベルク','フォルテ'],
    stockPrice:6300, marketCap:130000, revenue:400000, opProfit:16500, netProfit:11000, opMargin:4.1, netMargin:2.8,
    per:10.5, pbr:1.2, roe:12.5, dividendYield:1.8, stores:140,
    monthlySSS:[4.2,4.5,3.8,3.4,4.3,4.9,3.9,3.6,4.5,5.2,4.0,3.7],
    monthlyCust:[1.8,2.1,1.4,1.1,1.9,2.4,1.5,1.2,2.0,2.5,1.6,1.3],
    monthlyBasket:[2.4,2.4,2.4,2.3,2.4,2.5,2.4,2.4,2.5,2.7,2.4,2.4] },
  { code:'2742', name:'ハローズ', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['ハローズ'],
    stockPrice:4400, marketCap:90000, revenue:210000, opProfit:10000, netProfit:6800, opMargin:4.8, netMargin:3.2,
    per:11.0, pbr:1.4, roe:13.2, dividendYield:1.3, stores:110,
    monthlySSS:[5.0,5.4,4.5,4.1,5.1,5.8,4.6,4.3,5.3,6.0,4.8,4.4],
    monthlyCust:[2.4,2.8,2.0,1.7,2.5,3.1,2.1,1.8,2.6,3.2,2.2,1.9],
    monthlyBasket:[2.6,2.6,2.5,2.4,2.6,2.7,2.5,2.5,2.7,2.8,2.6,2.5] },
  { code:'3222', name:'ユナイテッド・スーパーマーケット・ホールディングス', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['マルエツ','カスミ','マックスバリュ関東'],
    stockPrice:900, marketCap:110000, revenue:700000, opProfit:8000, netProfit:4000, opMargin:1.1, netMargin:0.6,
    per:26.0, pbr:0.9, roe:3.5, dividendYield:1.8, stores:520,
    monthlySSS:[1.8,2.1,1.5,1.2,1.9,2.4,1.6,1.3,2.0,2.5,1.7,1.4],
    monthlyCust:[-0.5,-0.2,-0.8,-1.1,-0.4,0.1,-0.7,-1.0,-0.3,0.2,-0.6,-0.9],
    monthlyBasket:[2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3] },
  { code:'8276', name:'平和堂', segment:'GMS', fiscalYear:'2026/2期(仮)', brands:['平和堂','アル・プラザ','フレンドマート'],
    stockPrice:2400, marketCap:130000, revenue:440000, opProfit:12500, netProfit:8000, opMargin:2.8, netMargin:1.8,
    per:13.5, pbr:0.8, roe:6.0, dividendYield:2.3, stores:160,
    monthlySSS:null, monthlyCust:null, monthlyBasket:null },
];

window.SECTOR_CONFIG = {
  slug: 'supermarket',
  sectorName: 'スーパーマーケット',
  storesLabel: '店舗数',
  months: ['25/4','25/5','25/6','25/7','25/8','25/9','25/10','25/11','25/12','26/1','26/2','26/3'],
  segments: { FOOD:'食品スーパー', GMS:'総合スーパー(GMS)', DISC:'ディスカウント・業務用' },
  segColors: { FOOD:'#2d7a4f', GMS:'#1a2d4f', DISC:'#c8946e' },
  heatmapPosThreshold: 4,
  heatmapNegThreshold: 0,
  monthlySeries: [
    { field:'monthlySSS', label:'既存店売上高 前年同月比 (%)', sub:'各社IR月次売上速報 / 破線=スーパーマーケット販売統計 既存店前年比 (仮置き値)',
      industryAvg:[3.8,4.1,3.5,3.2,3.9,4.4,3.6,3.3,4.0,4.5,3.7,3.4], industryAvgLabel:'業界平均(販売統計)' },
    { field:'monthlyCust', label:'既存店客数 前年同月比 (%)', sub:'各社IR月次開示 (仮置き値)' },
    { field:'monthlyBasket', label:'既存店客単価 前年同月比 (%)', sub:'各社IR月次開示 (仮置き値)' },
  ],
  execCommentary: `
    <strong>セクター概況 (ドラフト):</strong> 食品スーパー・GMS・ディスカウント業態の主要上場${'12'}社をカバー。
    食品インフレの継続により<strong>客単価上昇が既存店売上を押し上げる</strong>一方、実質賃金の伸び悩みで客数と買上点数には節約志向の影が残る。
    値上げ疲れの消費者を取り込む<strong>ディスカウント業態(業務スーパー・トライアル)の高成長</strong>と、リテールメディア・AI発注などのDX投資が二大テーマ。<br><br>
    <strong>注目ポイント(実データ反映後に検証):</strong>
    (1) <strong>客数×客単価の分解</strong> — インフレ一巡後にオーガニックな客数成長を確保できているか、
    (2) <strong>PB比率と粗利率</strong>の関係(コスト転嫁力の代理指標)、
    (3) 地方スーパーの<strong>再編・M&A</strong>(イオン系再編、アークス連合等)とPBR1倍割れ企業のバリュエーション、
    (4) ネットスーパー・クイックコマースの損益分岐点への距離。`,
  monthlyCommentary: `
    <strong>月次KPIの設計 (ドラフト):</strong> スーパー各社は<strong>既存店売上高・客数・客単価の前年同月比</strong>を月次開示しており、
    「インフレによる単価上昇」と「実需の客数」を分解して追跡できる。
    業界平均(破線)は全国スーパーマーケット協会等3団体「スーパーマーケット販売統計調査」の既存店前年比を使用予定。
    イズミ・バロー・平和堂は月次非開示のため四半期データで補完する。`,
  monthlyMethodology: `
    <strong>データ方法論 (ドラフト / 実データ反映時に確定):</strong><br>
    ・<strong>対象期間:</strong> 2025年4月～2026年3月 (暦月ベース、2月期・3月期・6月期・10月期企業も暦月で統一)<br>
    ・<strong>既存店売上・客数・客単価:</strong> 各社IR月次営業速報の既存店前年同月比。GMSは食品部門ではなく全社ベース<br>
    ・<strong>業界平均:</strong> スーパーマーケット販売統計調査(全国スーパーマーケット協会・日本スーパーマーケット協会・オール日本スーパーマーケット協会)の既存店前年比<br>
    ・<strong>非開示企業:</strong> イズミ・バローHD・平和堂は月次非開示。四半期決算データから按分推計するか空欄とする<br>
    ・<strong>更新頻度:</strong> 各社月次速報は翌月上旬に開示。毎月中旬にデータ更新を実施予定`,
  macroCommentary: `
    <strong>外部環境 (ドラフト):</strong> スーパーマーケット販売額は食品インフレを追い風に<strong>約15兆円規模へ拡大</strong>。
    ただし成長の大半は単価上昇によるもので、食品CPIの伸びが鈍化する局面では既存店成長率の減速が想定される。
    世帯の食料支出は名目では過去最高水準に達しエンゲル係数は約28%と歴史的高水準 — <strong>節約志向とプレミアム消費の二極化</strong>が業態間の勝敗を分ける。
    ※以下のチャートはいずれも仮置き値。実データは商業動態統計・総務省CPI・家計調査から取得予定。`,
  macroCharts: [
    { title:'スーパーマーケット販売額の推移', sub:'出所(予定): 経済産業省 商業動態統計 / 単位: 兆円 (仮置き値)', type:'bar', yTitle:'兆円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'販売額', data:[13.2,13.0,13.1,13.2,13.1,13.7,13.5,13.6,14.1,14.7,15.2] }] },
    { title:'食品CPI(生鮮除く)前年比の推移', sub:'出所(予定): 総務省 消費者物価指数 / (仮置き値)', type:'line', yTitle:'%',
      labels:['2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'食品CPI前年比', data:[0.5,1.2,0.0,4.5,8.2,3.5,4.0] }] },
    { title:'二人以上世帯 食料支出(月額)の推移', sub:'出所(予定): 総務省 家計調査 / 単位: 千円 (仮置き値)', type:'bar', yTitle:'千円',
      labels:['2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'食料支出月額', data:[73,75,76,78,81,86,90] }] },
    { title:'業態別 販売額前年比 (直近年)', sub:'出所(予定): 商業動態統計・各業界団体 (仮置き値)', type:'bar',
      labels:['ディスカウント・業務用','食品スーパー','ドラッグストア(食品)','コンビニ','GMS','百貨店(食品)'],
      series:[{ label:'前年比(%)', data:[7.5,3.8,8.5,2.5,1.5,1.8] }] },
  ],
};
