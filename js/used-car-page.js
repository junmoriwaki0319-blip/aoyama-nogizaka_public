// ============================================================
// 中古車小売セクター分析レポート - データ定義 (ドラフト版)
// ============================================================
// ※ 本ファイルの数値はすべてレポート構成確認用の仮置き値。
//    SPEEDA・各社IR月次開示・自販連/USS統計の実データ受領後に差し替える。
// 月次KPI: 小売販売台数前年比 / 買取(仕入)台数前年比 (各社月次IR)
// 業界平均: 日本自動車販売協会連合会「中古車登録台数」前年同月比
// 留意: 旧ビッグモーター(WECARS)は非上場のため対象外。業界需給への影響はコメントで補足
// ============================================================

var companies = [
  { code:'4732', name:'ユー・エス・エス', segment:'AUCTION', fiscalYear:'2026/3期(仮)', brands:['USSオートオークション'],
    stockPrice:1450, marketCap:720000, revenue:115000, opProfit:58000, netProfit:40000, opMargin:50.4, netMargin:34.8,
    per:16.0, pbr:2.8, roe:18.0, dividendYield:2.4, stores:21,
    monthlyRetail:[2.5,3.0,1.8,2.8,3.5,2.0,1.4,2.6,3.8,3.0,2.2,3.4],
    monthlyBuy:null },
  { code:'3186', name:'ネクステージ', segment:'RETAIL', fiscalYear:'2026/11期(仮)', brands:['ネクステージ','ユニバース'],
    stockPrice:2300, marketCap:190000, revenue:560000, opProfit:19000, netProfit:12000, opMargin:3.4, netMargin:2.1,
    per:14.0, pbr:1.8, roe:14.5, dividendYield:1.4, stores:340,
    monthlyRetail:[4.5,6.0,3.2,5.5,7.0,4.0,2.8,5.2,8.0,6.5,4.8,7.2],
    monthlyBuy:[3.5,5.0,2.4,4.5,6.0,3.2,2.0,4.2,7.0,5.5,3.8,6.2] },
  { code:'7599', name:'IDOM', segment:'RETAIL', fiscalYear:'2026/2期(仮)', brands:['ガリバー','LIBERALA'],
    stockPrice:1300, marketCap:130000, revenue:510000, opProfit:19500, netProfit:12500, opMargin:3.8, netMargin:2.5,
    per:10.0, pbr:1.4, roe:14.8, dividendYield:2.0, stores:460,
    monthlyRetail:[6.0,7.5,4.5,6.8,8.5,5.2,4.0,6.5,9.5,8.0,6.0,8.8],
    monthlyBuy:[5.0,6.5,3.8,5.8,7.5,4.4,3.2,5.5,8.5,7.0,5.2,7.8] },
  { code:'9856', name:'ケーユーホールディングス', segment:'RETAIL', fiscalYear:'2026/3期(仮)', brands:['ケーユー'],
    stockPrice:1500, marketCap:58000, revenue:135000, opProfit:8800, netProfit:6000, opMargin:6.5, netMargin:4.4,
    per:9.5, pbr:0.8, roe:9.0, dividendYield:3.0, stores:60,
    monthlyRetail:null, monthlyBuy:null },
  { code:'7593', name:'VTホールディングス', segment:'DEALER', fiscalYear:'2026/3期(仮)', brands:['ホンダカーズ東海','日産サティオ','J-net'],
    stockPrice:600, marketCap:70000, revenue:360000, opProfit:16000, netProfit:9500, opMargin:4.4, netMargin:2.6,
    per:7.5, pbr:0.9, roe:12.0, dividendYield:3.6, stores:250,
    monthlyRetail:null, monthlyBuy:null },
  { code:'7676', name:'グッドスピード', segment:'RETAIL', fiscalYear:'2026/9期(仮)', brands:['グッドスピード'],
    stockPrice:700, marketCap:13000, revenue:125000, opProfit:2800, netProfit:1500, opMargin:2.2, netMargin:1.2,
    per:9.0, pbr:1.1, roe:11.0, dividendYield:0.0, stores:45,
    monthlyRetail:[3.0,4.5,1.8,4.0,5.5,2.6,1.5,3.8,6.5,5.0,3.4,5.8],
    monthlyBuy:[2.0,3.5,1.0,3.0,4.5,1.8,0.8,2.8,5.5,4.0,2.4,4.8] },
  { code:'4298', name:'プロトコーポレーション', segment:'SERVICE', fiscalYear:'2026/3期(仮)', brands:['グーネット','グー買取'],
    stockPrice:1500, marketCap:57000, revenue:68000, opProfit:7800, netProfit:5200, opMargin:11.5, netMargin:7.6,
    per:10.5, pbr:0.9, roe:9.5, dividendYield:2.8, stores:null,
    monthlyRetail:null, monthlyBuy:null },
  { code:'7199', name:'プレミアグループ', segment:'SERVICE', fiscalYear:'2026/3期(仮)', brands:['プレミアファイナンシャルサービス','カーセブン'],
    stockPrice:2600, marketCap:100000, revenue:42000, opProfit:10500, netProfit:6800, opMargin:25.0, netMargin:16.2,
    per:14.0, pbr:2.6, roe:19.5, dividendYield:2.6, stores:null,
    monthlyRetail:null, monthlyBuy:null },
  { code:'9832', name:'オートバックスセブン', segment:'SERVICE', fiscalYear:'2026/3期(仮)', brands:['オートバックス','カーズ'],
    stockPrice:1600, marketCap:130000, revenue:250000, opProfit:12500, netProfit:8500, opMargin:5.0, netMargin:3.4,
    per:14.5, pbr:0.9, roe:6.5, dividendYield:3.8, stores:590,
    monthlyRetail:null, monthlyBuy:null },
];

window.SECTOR_CONFIG = {
  slug: 'used-car',
  sectorName: '中古車小売',
  storesLabel: '店舗・拠点数',
  months: ['25/4','25/5','25/6','25/7','25/8','25/9','25/10','25/11','25/12','26/1','26/2','26/3'],
  segments: { RETAIL:'中古車小売専業', AUCTION:'オートオークション', DEALER:'ディーラー系', SERVICE:'情報・金融・関連サービス' },
  segColors: { RETAIL:'#1a2d4f', AUCTION:'#9b8b6e', DEALER:'#5a7fa8', SERVICE:'#2d7a4f' },
  heatmapPosThreshold: 5,
  heatmapNegThreshold: 0,
  monthlySeries: [
    { field:'monthlyRetail', label:'小売販売台数 前年同月比 (%)', sub:'各社IR月次開示 (USSはオークション成約台数) / 破線=自販連 中古車登録台数前年比 (仮置き値)',
      industryAvg:[1.5,2.2,0.8,1.9,2.5,1.2,0.5,1.8,2.9,2.1,1.4,2.6], industryAvgLabel:'中古車登録台数(業界)' },
    { field:'monthlyBuy', label:'買取(仕入)台数 前年同月比 (%)', sub:'各社IR月次開示 / 仕入競争力の代理指標 (仮置き値)' },
  ],
  execCommentary: `
    <strong>セクター概況 (ドラフト):</strong> 中古車小売専業・オートオークション・ディーラー系・関連サービスの主要上場${'9'}社をカバー。
    2023年のビッグモーター問題(現WECARS)による業界信頼の毀損と大手1社の急収縮を経て、<strong>ネクステージ・IDOM(ガリバー)がシェアを吸収</strong>する再編局面。
    新車供給正常化により中古車の玉不足は緩和へ向かう一方、<strong>円安を背景とする中古車輸出の高水準</strong>が国内オークション相場を下支えしている。<br><br>
    <strong>注目ポイント(実データ反映後に検証):</strong>
    (1) <strong>小売台数×台当たり粗利</strong>の分解(相場下落局面では在庫評価損リスク)、
    (2) <strong>買取台数の成長率</strong> — 小売専業の競争力は仕入(買取)チャネルで決まる、
    (3) 付帯収益(整備・保証・ファイナンス・保険)の<strong>ストック化比率</strong>とコンプライアンス体制、
    (4) オークション(USS)は台数×手数料単価のインフラ型収益で景気耐性が高く、セクター内でディフェンシブな位置付け。`,
  monthlyCommentary: `
    <strong>月次KPIの設計 (ドラフト):</strong> 小売専業各社は<strong>月次で小売販売台数・買取台数</strong>を開示しており(IDOM・ネクステージ等)、
    USSは<strong>オークション出品・成約台数</strong>を月次開示。業界全体の需要は自販連「中古車登録台数」で捕捉する。
    外食の既存店売上にあたる「月次の体温計」として、小売台数前年比を第一KPIに設定。ケーユー・VT・オートバックス等は月次非開示のため四半期データで補完する。`,
  monthlyMethodology: `
    <strong>データ方法論 (ドラフト / 実データ反映時に確定):</strong><br>
    ・<strong>対象期間:</strong> 2025年4月～2026年3月 (暦月ベース、2月期・9月期・11月期企業も暦月で統一)<br>
    ・<strong>小売販売台数:</strong> 各社IR月次開示の小売台数前年同月比。USSはオークション成約台数で代替<br>
    ・<strong>買取台数:</strong> 各社IR月次開示。買取専門店・店頭買取・出張買取の合算<br>
    ・<strong>業界平均:</strong> 日本自動車販売協会連合会「中古車登録台数(乗用車)」前年同月比 ※軽自動車は全軽自協データを加算するか実データ反映時に決定<br>
    ・<strong>非開示企業:</strong> ケーユーHD・VTHD・プロト・プレミアG・オートバックスは月次非開示。四半期データで補完<br>
    ・<strong>更新頻度:</strong> 各社月次速報・自販連統計は翌月上旬に開示。毎月中旬にデータ更新を実施予定`,
  macroCommentary: `
    <strong>外部環境 (ドラフト):</strong> 国内中古車登録台数は年間約380万台(乗用車)で成熟市場だが、
    <strong>円安による中古車輸出の拡大(年間180万台規模)</strong>が国内流通の需給をタイト化させ、オークション平均落札価格は
    コロナ前対比で約1.5倍の高水準を維持。新車の納期正常化・EV下取り価格の不安定さ・OBD検査義務化などが
    今後の相場と業者収益のスイングファクター。※以下のチャートはいずれも仮置き値。実データは自販連・全軽自協・USS月次データ・財務省貿易統計から取得予定。`,
  macroCharts: [
    { title:'中古車登録台数の推移 (乗用車)', sub:'出所(予定): 日本自動車販売協会連合会 / 単位: 万台 (仮置き値)', type:'bar', yTitle:'万台',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'中古車登録台数', data:[386,384,381,377,372,355,340,330,345,352,358] }] },
    { title:'中古車輸出台数の推移', sub:'出所(予定): 財務省 貿易統計 / 単位: 万台 (仮置き値)', type:'line', yTitle:'万台',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'中古車輸出台数', data:[110,120,130,135,140,120,130,150,160,170,180] }] },
    { title:'オークション平均落札価格の推移', sub:'出所(予定): USS月次データ等 / 単位: 万円 (仮置き値)', type:'line', yTitle:'万円',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'平均落札価格', data:[65,68,70,74,80,95,110,105,102,108,112] }] },
    { title:'新車販売台数の推移 (参考)', sub:'出所(予定): 自販連・全軽自協 / 単位: 万台 (仮置き値)', type:'bar', yTitle:'万台',
      labels:['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
      series:[{ label:'新車販売台数', data:[505,497,523,527,520,460,445,420,478,442,455] }] },
  ],
};
