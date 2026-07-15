// ============================================================
// ドラッグストアセクター分析レポート - データ定義 (ドラフト版)
// ============================================================
// ※ 本ファイルの数値はすべてレポート構成確認用の仮置き値。
//    SPEEDA・各社IR月次開示・経産省商業動態統計の実データ受領後に差し替える。
// 月次KPI: 既存店売上高前年比 / 既存店客数前年比 / 既存店客単価前年比
// 業界平均: 経済産業省「商業動態統計」ドラッグストア販売額前年比
// 留意: ツルハHD・ウエルシアHDの経営統合(イオン主導)の扱いは実データ反映時に確認
// ============================================================

var companies = [
  { code:'3391', name:'ツルハホールディングス', segment:'MEGA', fiscalYear:'2026/5期(仮)', brands:['ツルハドラッグ','くすりの福太郎','ドラッグイレブン'],
    stockPrice:12500, marketCap:620000, revenue:1050000, opProfit:56000, netProfit:36000, opMargin:5.3, netMargin:3.4,
    per:19.0, pbr:1.9, roe:10.2, dividendYield:1.6, stores:2650,
    monthlySSS:[4.2,3.8,5.1,4.5,3.2,4.8,5.5,4.0,3.6,5.2,4.4,3.9],
    monthlyCust:[1.5,1.2,2.0,1.8,0.9,1.6,2.2,1.4,1.1,2.0,1.5,1.2],
    monthlyBasket:[2.7,2.6,3.1,2.7,2.3,3.2,3.3,2.6,2.5,3.2,2.9,2.7] },
  { code:'3141', name:'ウエルシアホールディングス', segment:'MEGA', fiscalYear:'2026/2期(仮)', brands:['ウエルシア薬局','ダックス','ハックドラッグ'],
    stockPrice:2200, marketCap:460000, revenue:1250000, opProfit:45000, netProfit:26000, opMargin:3.6, netMargin:2.1,
    per:20.5, pbr:1.6, roe:8.0, dividendYield:1.5, stores:2800,
    monthlySSS:[3.5,3.0,4.4,3.8,2.5,4.0,4.8,3.3,2.9,4.5,3.7,3.2],
    monthlyCust:[1.0,0.8,1.6,1.3,0.5,1.2,1.8,1.0,0.7,1.6,1.1,0.8],
    monthlyBasket:[2.5,2.2,2.8,2.5,2.0,2.8,3.0,2.3,2.2,2.9,2.6,2.4] },
  { code:'3088', name:'マツキヨココカラ&カンパニー', segment:'MEGA', fiscalYear:'2026/3期(仮)', brands:['マツモトキヨシ','ココカラファイン'],
    stockPrice:2250, marketCap:920000, revenue:1080000, opProfit:82000, netProfit:56000, opMargin:7.6, netMargin:5.2,
    per:16.5, pbr:1.8, roe:11.5, dividendYield:1.7, stores:3400,
    monthlySSS:[5.5,4.8,6.5,5.8,4.2,6.0,7.2,5.0,4.5,6.8,5.6,5.0],
    monthlyCust:[2.8,2.4,3.5,3.0,2.0,3.2,4.0,2.6,2.2,3.6,2.9,2.5],
    monthlyBasket:[2.7,2.4,3.0,2.8,2.2,2.8,3.2,2.4,2.3,3.2,2.7,2.5] },
  { code:'3349', name:'コスモス薬品', segment:'FOOD', fiscalYear:'2026/5期(仮)', brands:['ディスカウントドラッグコスモス'],
    stockPrice:6500, marketCap:410000, revenue:950000, opProfit:41000, netProfit:28000, opMargin:4.3, netMargin:2.9,
    per:17.5, pbr:2.1, roe:12.5, dividendYield:0.9, stores:1550,
    monthlySSS:null, monthlyCust:null, monthlyBasket:null },
  { code:'9989', name:'サンドラッグ', segment:'DISC', fiscalYear:'2026/3期(仮)', brands:['サンドラッグ','ドラッグトップス'],
    stockPrice:4100, marketCap:470000, revenue:820000, opProfit:41000, netProfit:28000, opMargin:5.0, netMargin:3.4,
    per:16.0, pbr:1.9, roe:12.0, dividendYield:2.3, stores:1400,
    monthlySSS:[3.8,3.2,4.6,4.0,2.8,4.2,5.0,3.5,3.1,4.8,3.9,3.4],
    monthlyCust:[1.8,1.4,2.4,2.0,1.1,1.9,2.6,1.6,1.3,2.4,1.8,1.4],
    monthlyBasket:[2.0,1.8,2.2,2.0,1.7,2.3,2.4,1.9,1.8,2.4,2.1,2.0] },
  { code:'7649', name:'スギホールディングス', segment:'RX', fiscalYear:'2026/2期(仮)', brands:['スギ薬局','スギドラッグ'],
    stockPrice:2800, marketCap:340000, revenue:820000, opProfit:38000, netProfit:25000, opMargin:4.6, netMargin:3.0,
    per:14.5, pbr:1.6, roe:11.0, dividendYield:1.4, stores:1800,
    monthlySSS:[5.0,4.4,5.8,5.2,3.8,5.4,6.2,4.6,4.1,6.0,5.0,4.5],
    monthlyCust:[2.2,1.8,2.8,2.4,1.5,2.4,3.0,2.0,1.7,2.8,2.2,1.9],
    monthlyBasket:[2.8,2.6,3.0,2.8,2.3,3.0,3.2,2.6,2.4,3.2,2.8,2.6] },
  { code:'3549', name:'クスリのアオキホールディングス', segment:'FOOD', fiscalYear:'2026/5期(仮)', brands:['クスリのアオキ'],
    stockPrice:2900, marketCap:290000, revenue:500000, opProfit:22000, netProfit:14000, opMargin:4.4, netMargin:2.8,
    per:18.0, pbr:2.2, roe:12.8, dividendYield:0.8, stores:950,
    monthlySSS:[6.5,5.8,7.5,6.8,5.0,7.0,8.2,6.0,5.4,7.8,6.5,5.9],
    monthlyCust:[3.5,3.0,4.2,3.8,2.6,3.9,4.8,3.2,2.8,4.4,3.6,3.1],
    monthlyBasket:[3.0,2.8,3.3,3.0,2.4,3.1,3.4,2.8,2.6,3.4,2.9,2.8] },
  { code:'3148', name:'クリエイトSDホールディングス', segment:'RX', fiscalYear:'2026/5期(仮)', brands:['クリエイト'],
    stockPrice:3100, marketCap:210000, revenue:450000, opProfit:21000, netProfit:14000, opMargin:4.7, netMargin:3.1,
    per:14.0, pbr:1.6, roe:12.2, dividendYield:1.3, stores:770,
    monthlySSS:[4.5,4.0,5.4,4.8,3.4,5.0,5.8,4.2,3.8,5.5,4.6,4.1],
    monthlyCust:[2.0,1.6,2.6,2.2,1.3,2.2,2.8,1.8,1.5,2.6,2.0,1.7],
    monthlyBasket:[2.5,2.4,2.8,2.6,2.1,2.8,3.0,2.4,2.3,2.9,2.6,2.4] },
  { code:'9267', name:'Genky DrugStores', segment:'FOOD', fiscalYear:'2026/6期(仮)', brands:['ゲンキー'],
    stockPrice:3800, marketCap:58000, revenue:200000, opProfit:9000, netProfit:5800, opMargin:4.5, netMargin:2.9,
    per:13.0, pbr:1.9, roe:15.0, dividendYield:0.8, stores:450,
    monthlySSS:[5.8,5.2,6.8,6.0,4.5,6.2,7.4,5.4,4.9,7.0,5.8,5.3],
    monthlyCust:[3.0,2.6,3.8,3.2,2.2,3.4,4.2,2.8,2.4,3.9,3.1,2.7],
    monthlyBasket:[2.8,2.6,3.0,2.8,2.3,2.8,3.2,2.6,2.5,3.1,2.7,2.6] },
  { code:'7679', name:'薬王堂ホールディングス', segment:'FOOD', fiscalYear:'2026/2期(仮)', brands:['薬王堂'],
    stockPrice:2300, marketCap:45000, revenue:150000, opProfit:6000, netProfit:4000, opMargin:4.0, netMargin:2.7,
    per:11.5, pbr:1.5, roe:13.5, dividendYield:1.1, stores:400,
    monthlySSS:[4.8,4.2,5.6,5.0,3.6,5.2,6.0,4.4,4.0,5.8,4.8,4.3],
    monthlyCust:[2.4,2.0,3.0,2.6,1.7,2.6,3.2,2.2,1.9,3.0,2.4,2.1],
    monthlyBasket:[2.4,2.2,2.6,2.4,1.9,2.6,2.8,2.2,2.1,2.8,2.4,2.2] },
  { code:'2664', name:'カワチ薬品', segment:'DISC', fiscalYear:'2026/3期(仮)', brands:['カワチ薬品'],
    stockPrice:2600, marketCap:56000, revenue:290000, opProfit:8500, netProfit:5500, opMargin:2.9, netMargin:1.9,
    per:10.0, pbr:0.6, roe:6.2, dividendYield:2.2, stores:360,
    monthlySSS:null, monthlyCust:null, monthlyBasket:null },
];

window.SECTOR_CONFIG = {
  slug: 'drugstore',
  sectorName: 'ドラッグストア',
  storesLabel: '店舗数',
  months: ['25/4','25/5','25/6','25/7','25/8','25/9','25/10','25/11','25/12','26/1','26/2','26/3'],
  segments: { MEGA:'全国大手', FOOD:'フード&ドラッグ', RX:'調剤併設強化型', DISC:'ディスカウント型' },
  segColors: { MEGA:'#1a2d4f', FOOD:'#2d7a4f', RX:'#5a7fa8', DISC:'#c8946e' },
  heatmapPosThreshold: 5,
  heatmapNegThreshold: 0,
  monthlySeries: [
    { field:'monthlySSS', label:'既存店売上高 前年同月比 (%)', sub:'各社IR月次売上速報 / 破線=商業動態統計 ドラッグストア販売額前年比 (仮置き値)',
      industryAvg:[5.8,6.2,5.1,4.8,5.5,6.0,5.2,4.9,5.6,6.3,5.8,5.4], industryAvgLabel:'業界平均(商業動態)' },
    { field:'monthlyCust', label:'既存店客数 前年同月比 (%)', sub:'各社IR月次開示 (仮置き値)' },
    { field:'monthlyBasket', label:'既存店客単価 前年同月比 (%)', sub:'各社IR月次開示 (仮置き値)' },
  ],
  execCommentary: `
    <strong>セクター概況 (ドラフト):</strong> 全国大手・フード&ドラッグ・調剤併設型の主要上場${'11'}社をカバー。
    市場規模は約9兆円と<strong>百貨店を上回る国内小売の主戦場</strong>に成長し、食品構成比の拡大によりスーパーマーケット・コンビニとの業態間競争が激化。
    <strong>イオン主導のツルハ・ウエルシア経営統合</strong>により売上2兆円超の巨大連合が誕生し、業界再編が最終局面へ。<br><br>
    <strong>注目ポイント(実データ反映後に検証):</strong>
    (1) <strong>調剤事業の成長</strong>(調剤報酬改定・敷地内薬局規制の影響と調剤併設率の推移)、
    (2) <strong>フード&ドラッグ業態</strong>(コスモス・ゲンキー・クスリのアオキ)の地方ドミナント拡大と既存店成長率の持続性、
    (3) 統合再編後の<strong>PB共同開発・物流統合によるマージン改善</strong>、
    (4) インバウンド(化粧品・OTC医薬品)復調の都市型店舗への寄与。`,
  monthlyCommentary: `
    <strong>月次KPIの設計 (ドラフト):</strong> ドラッグストア各社は<strong>既存店売上高・客数・客単価の前年同月比</strong>を月次開示しており、
    「客数×客単価」への分解で成長ドライバー(集客か単価か)を判別できる。
    業界平均(破線)は経産省「商業動態統計」のドラッグストア販売額前年比を使用予定。コスモス薬品・カワチ薬品は月次非開示のため四半期データで補完する。`,
  monthlyMethodology: `
    <strong>データ方法論 (ドラフト / 実データ反映時に確定):</strong><br>
    ・<strong>対象期間:</strong> 2025年4月～2026年3月 (暦月ベース、2月期・3月期・5月期企業も暦月で統一)<br>
    ・<strong>既存店売上・客数・客単価:</strong> 各社IR月次営業速報の既存店前年同月比<br>
    ・<strong>業界平均:</strong> 経済産業省「商業動態統計」ドラッグストア販売額前年比(全店ベースのため既存店比較では1-2pt高めに出る点に留意)<br>
    ・<strong>非開示企業:</strong> コスモス薬品・カワチ薬品は月次非開示。四半期決算の既存店データから按分推計するか空欄とする<br>
    ・<strong>更新頻度:</strong> 各社月次速報は翌月上旬に開示。毎月中旬にデータ更新を実施予定`,
  macroCommentary: `
    <strong>外部環境 (ドラフト):</strong> ドラッグストア販売額は食品取扱拡大を追い風に<strong>年率5%前後の成長を継続</strong>し、約9兆円規模へ。
    店舗数は2.3万店を超え飽和感が指摘されるが、調剤併設化とフード&ドラッグ化で客数を維持。
    調剤医療費は高齢化を背景に拡大が続き、<strong>門前薬局からドラッグストア調剤へのシフト</strong>が構造的な追い風。
    ※以下のチャートはいずれも仮置き値。実データは経産省商業動態統計・厚労省調剤医療費動向から取得予定。`,
  macroCharts: [
    { title:'ドラッグストア販売額の推移', sub:'出所(予定): 経済産業省 商業動態統計 / 単位: 兆円 (仮置き値)', type:'bar', yTitle:'兆円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'販売額', data:[5.4,5.7,6.0,6.3,6.8,7.2,7.4,7.7,8.4,9.0,9.5] }] },
    { title:'ドラッグストア店舗数の推移', sub:'出所(予定): 日本チェーンドラッグストア協会 / 単位: 千店 (仮置き値)', type:'line', yTitle:'千店',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'店舗数', data:[18.8,19.3,19.9,20.6,21.0,21.5,21.9,22.1,22.6,23.0,23.4] }] },
    { title:'調剤医療費の推移', sub:'出所(予定): 厚生労働省 調剤医療費の動向 / 単位: 兆円 (仮置き値)', type:'bar', yTitle:'兆円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024'],
      series:[{ label:'調剤医療費', data:[7.9,7.5,7.7,7.5,7.7,7.5,7.8,7.7,8.3,8.5] }] },
    { title:'商品カテゴリ別 販売額前年比 (直近年)', sub:'出所(予定): 商業動態統計 品目別 (仮置き値)', type:'bar',
      labels:['食品','調剤医薬品','OTC医薬品','化粧品','日用雑貨','ヘルスケア用品'],
      series:[{ label:'前年比(%)', data:[8.5,7.2,3.5,4.8,2.5,3.0] }] },
  ],
};
