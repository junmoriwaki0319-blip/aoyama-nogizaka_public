// ===== SCRIPT 1: Data Definition =====
// ============================================================
// 外食産業KPIダッシュボード - データ定義 v4
// ============================================================
// MUMSSレポート準拠のTop15 + 追加銘柄で構成
// デュポン分解・TOPIXパフォーマンス・マクロ指標を追加
// 既存店売上高データは各社IR月次開示に基づく (sameStoreSales = 前年同月比%)
// SFP・イートアンド・クリエイトRは月次非開示のためnull、リンガーハット1月は補間推定
// ============================================================

const SEGMENTS = {
  FF: 'ファストフード',
  FR: 'カジュアルレストラン',
  IZK: '居酒屋・バー',
  SUSHI: '持ち帰り・回転寿司',
  CAFE: '喫茶・カフェ',
  GYU: '定食・そば・うどん',
  RAMEN: '中華麺',
  YAKINIKU: '焼肉',
  OTHER: 'その他',
};

const SUB_SECTORS_EN = {
  FF: 'Fast Food',
  FR: 'Casual Restaurants',
  IZK: 'Pubs Bars Izakayas',
  SUSHI: 'Conveyor Belt Sushi',
  CAFE: 'Tea Coffee Shops',
  GYU: 'Casual Japanese',
  RAMEN: 'Ramen',
  YAKINIKU: 'Yakiniku',
  OTHER: 'Other',
};

// ---------- 企業マスタ ----------
// 企業データはFirestoreから動的取得（認証後にloadPremiumData()で取得）
let companies = [];


// ---------- 月ラベル ----------
const MONTHS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];

// 日本フードサービス協会 既存店売上高ベース 業界平均(前年同月比%)
// 注: JF-NET公表の「全店」データより系統的に1-2pt低い(既存店ベースのため)
// 24/4=105.0, 5=104.5, 6=111.2, 7=103.8, 8=108.1, 9=107.6, 10=104.6
const SECTOR_AVG_SSS = [105.0,104.5,111.2,103.8,108.1,107.6,104.6,110.5,110.6,109.7,110.6,110.6];


// ---------- データ基準日 ----------
const DATA_AS_OF = {
  // stockPrice / topixPeriod は loadPremiumData で実データ(updatedAt / INDEX_MONTHS)から動的設定する
  stockPrice: '各社最新株価',
  financials: '各社直近本決算(Yahoo!ファイナンス自動取得)',
  sameStoreSales: '2024年4月～2025年3月',
  topixPeriod: '直近1年',
  minimumWage: '2024年10月改定(厚生労働省)',
};
// ---------- TOPIX Performance ----------
// マーケットデータ: Firestoreから動的取得（認証後）
let TOPIX_RETURN_1Y, TOPIX_MONTHLY, RESTAURANT_INDEX_MONTHLY, INDEX_MONTHS;
let SECTOR_AVG_SSS_HISTORY = {}; // { FY2023: [12], FY2022: [12] }

// TOPIX-17比較(直近1年, 100基準)
const TOPIX17_SECTORS = [
  { name: '情報通信・サービス', point: 142.5 },
  { name: '銀行', point: 140.2 },
  { name: '小売', point: 135.8 },
  { name: '金融(除く銀行)', point: 134.5 },
  { name: 'Restaurant', point: 130.0 },
  { name: 'TOPIX', point: 129.8 },
  { name: '建設・資材', point: 128.5 },
  { name: '不動産', point: 127.2 },
  { name: '機械', point: 126.0 },
  { name: '運輸・物流', point: 125.5 },
  { name: '食品', point: 122.8 },
  { name: '電機・精密', point: 121.5 },
  { name: '鉄鋼・非鉄', point: 118.2 },
  { name: '医薬品', point: 115.5 },
  { name: '商社・卸売', point: 112.8 },
  { name: '素材・化学', point: 110.5 },
  { name: '自動車・輸送機', point: 108.2 },
  { name: 'エネルギー資源', point: 105.8 },
  { name: '電力・ガス', point: 102.5 },
];

// ---------- マクロ指標 ----------
const MACRO_DATA = {
  // 主要食品価格指数 (2015年10月=100)
  // 2025年はCPI品目別(2020年基準)の2025年平均前年比(米類+67.5%/鶏卵+10.3%/鶏肉+6.3%/輸入牛肉+5.6%/国産豚肉+5.3%/輸入豚肉+5.2%)を2024年値に乗じて延長。輸入鶏肉はCPIに区分品目がないためnull
  foodPrices: {
    years: ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025'],
    items: {
      '米':        [100,101,102,103,104,103,105,108,120,166,278],
      '鶏卵':      [100,98,99,101,103,110,115,125,118,110,121],
      'ブロイラー(国産)': [100,99,100,101,103,102,105,110,108,101,107],
      'ブロイラー(輸入)': [100,98,99,102,105,108,115,128,125,122,null],
      '牛肉(輸入)':  [100,102,105,108,112,110,118,135,140,138,146],
      '豚肉(国産)':  [100,101,103,105,108,112,120,132,136,134,141],
      '豚肉(輸入)':  [100,99,101,104,108,110,118,130,132,130,137],
    },
    // 2025年平均の絶対価格(小売物価統計調査・東京都区部年平均)
    absNote: '2025年平均の店頭価格(東京都区部): 米(コシヒカリ)5kg 4,928円 / 鶏卵10個 307円 / 鶏もも肉100g 153円 / 輸入牛肉100g 362円 / 国産豚バラ100g 283円 / 輸入豚ロース100g 176円 (出所: 総務省 小売物価統計調査)',
  },
  // 家計調査 二人以上世帯 年間飲食費(万円) 2025年は品目別年間支出金額(e-Stat 第4-1表)実数を万円換算
  householdSpending: {
    years: ['2015','2017','2019','2021','2023','2025'],
    categories: {
      '日本そば・うどん':  [0.6,0.6,0.7,0.5,0.7,0.8],
      '中華そば':          [0.6,0.6,0.6,0.7,0.6,1.0],
      'すし(外食)':        [1.5,1.5,1.5,1.3,1.6,1.8],
      '和食':              [2.3,2.3,2.3,1.8,2.4,2.8],
      '中華食':            [0.5,0.5,0.5,0.4,0.5,0.5],
      '洋食':              [1.3,1.3,1.2,0.8,1.2,1.5],
      '焼肉':              [0.6,0.7,0.7,0.6,0.8,0.9],
      'ハンバーガー':      [0.3,0.4,0.5,0.6,0.6,0.7],
      '喫茶代':            [0.6,0.6,0.8,0.6,0.9,1.2],
      '飲酒代':            [1.9,1.8,2.0,0.5,1.6,1.9],
    },
    totals: [13.4, 13.5, 13.9, 10.4, 14.0, 13.1],
  },
  // 外食産業市場規模(狭義・外食産業計、兆円) 出所: 食の安全・安心財団推計。2024・2025年推計は2026年7月時点で未公表
  marketSize: {
    years: ['2019','2020','2021','2022','2023'],
    values: [26.3, 18.2, 17.0, 20.1, 24.2],
  },
  // 外食産業主要指標推移 (2019/12=100)
  // 25/06・25/12はJF月次調査の前年同月比(25/06: 売上106.0/店舗100.7/客数101.9/客単価104.1、25/12: 売上106.0/店舗101.1/客数102.4/客単価103.5)を前年点に乗じて延長
  industryIndex: {
    months: ['19/12','20/06','20/12','21/06','21/12','22/06','22/12','23/06','23/12','24/06','24/12','25/06','25/12'],
    sales:     [100,62,78,72,85,90,95,100,105,112,118.8,118.7,125.9],
    stores:    [100,97,95,93,92,91,91,91.5,92,93,93.2,93.7,94.2],
    customers: [100,58,72,68,80,84,88,90,92,93,94.4,94.8,96.7],
    unitPrice: [100,107,108,106,106,107,108,111,114,120,125.9,124.9,130.3],
  },
  // 人件費指数(2019=100)
  laborCost: {
    years: ['2019','2020','2021','2022','2023','2024','2025'],
    minWage: [901,902,930,961,1004,1055,1121],  // 全国加重平均最低賃金(円) 各年度改定額
    index:   [100,100.1,103.2,106.7,111.4,117.1,124.4],
  },
};

// ---------- ユーティリティ ----------
function fmtBil(v) {
  if (v >= 1000000) return (v/1000000).toFixed(2)+'兆円';
  if (v >= 10000) return (v/10000).toFixed(1)+'百億円';
  return v.toLocaleString()+'百万円';
}
function shortName(n) { return n.replace(/ホールディングス|HD|COMPANIES/g,'').trim(); }
function avg(arr, key) { const vals = arr.map(c => c[key]).filter(v => v != null); return vals.length ? (vals.reduce((s,v) => s+v, 0)/vals.length).toFixed(1) : '-'; }
function topN(arr, key, n) { return [...arr].filter(c => c[key] != null).sort((a,b) => b[key]-a[key]).slice(0,n); }
// null安全フォーマッタ (54社拡張で一部企業に未算出のnull指標が存在するため。saas-page.js の nv と同等)
function nv(v, suf='', fmt) { if (v == null) return '-'; if (fmt === 'loc') return v.toLocaleString()+suf; if (fmt === 'f1') return v.toFixed(1)+suf; if (fmt === 'f2') return v.toFixed(2)+suf; return v+suf; }
function sv(v, d=1) { return v == null ? '-' : (v > 0 ? '+' : '') + v.toFixed(d) + '%'; }


// ===== SCRIPT 2: Dashboard Logic =====
// ============================================================
// Restaurants セクター分析 KPIダッシュボード v3
// aoyama-nogizaka.com デザイン準拠
// ============================================================
(() => {
  'use strict';
  Chart.defaults.color = '#777777';
  Chart.defaults.borderColor = '#e0ddd6';
  Chart.defaults.font.family = "'Noto Sans JP', system-ui";
  Chart.defaults.font.size = 11;
  Chart.register(ChartDataLabels);
  Chart.defaults.plugins.datalabels = { display: false };

  const P = ['#1a2d4f','#9b8b6e','#2d7a4f','#b53a3a','#5a7fa8','#c8946e','#6b8e5e','#8b6b8e','#4a8b8b','#a89b5a','#7a5a3a','#5a6b8e'];
  const SC = { FF:'#1a2d4f', FR:'#5a7fa8', IZK:'#b53a3a', SUSHI:'#4a8b8b', CAFE:'#9b8b6e', GYU:'#c8946e', RAMEN:'#8b6b8e', YAKINIKU:'#a89b5a', OTHER:'#777777' };

  let tab = 'exec', selComp = null, mCodes = [], mSeg = 'ALL', mFY = 'latest';
  const C = {};

  function recalcYutai() {
    // 優待利回り再計算: yutaiYield = yutaiValue / (stockPrice × yutaiMinShares) × 100
    companies.forEach(c => {
      if (c.yutaiValue > 0 && c.yutaiMinShares > 0) {
        c.yutaiYield = parseFloat((c.yutaiValue / (c.stockPrice * c.yutaiMinShares) * 100).toFixed(2));
      } else { c.yutaiYield = 0; }
      c.totalYutaiYield = parseFloat((c.dividendYield + c.yutaiYield).toFixed(2));
    });
  }

  // Firestoreからプレミアムデータを取得してダッシュボード初期化
  window.loadPremiumData = async function() {
    if (!window.firebaseDb) return;
    try {
      const { doc, getDoc } = await import('https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js');
      const db = window.firebaseDb;
      // マーケットデータ取得
      const marketSnap = await getDoc(doc(db, 'premiumContent', 'food-market'));
      if (marketSnap.exists()) {
        const m = marketSnap.data();
        TOPIX_RETURN_1Y = m.TOPIX_RETURN_1Y;
        TOPIX_MONTHLY = m.TOPIX_MONTHLY;
        RESTAURANT_INDEX_MONTHLY = m.RESTAURANT_INDEX_MONTHLY;
        INDEX_MONTHS = m.INDEX_MONTHS;
        if (m.SECTOR_AVG_SSS_FY2023) SECTOR_AVG_SSS_HISTORY['FY2023'] = m.SECTOR_AVG_SSS_FY2023;
        if (m.SECTOR_AVG_SSS_FY2022) SECTOR_AVG_SSS_HISTORY['FY2022'] = m.SECTOR_AVG_SSS_FY2022;
        // TOPIX比較期間ラベルを INDEX_MONTHS の直近1年(13点)から動的生成
        if (Array.isArray(INDEX_MONTHS) && INDEX_MONTHS.length) {
          const fmtM = s => { const p = String(s).split('/'); return p.length === 2 ? `20${p[0]}年${parseInt(p[1], 10)}月` : s; };
          const n = INDEX_MONTHS.length, s0 = n >= 13 ? n - 13 : 0;
          DATA_AS_OF.topixPeriod = `${fmtM(INDEX_MONTHS[s0])}～${fmtM(INDEX_MONTHS[n - 1])}`;
        }
      }
      // 企業データ取得
      const compSnap = await getDoc(doc(db, 'premiumContent', 'food-companies'));
      if (compSnap.exists()) {
        const cd = compSnap.data();
        companies = cd.companies || [];
        // 株価基準日ラベルを doc の updatedAt(=アップロード日)から動的設定
        if (cd.updatedAt) {
          const d = new Date(cd.updatedAt);
          if (!isNaN(d)) DATA_AS_OF.stockPrice = `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日時点`;
        }
      }
      const asOfEl = document.getElementById('stockAsOf');
      if (asOfEl) asOfEl.textContent = DATA_AS_OF.stockPrice;
      recalcYutai();
      document.getElementById('companyCount').textContent = companies.length;
      render();
    } catch (e) {
      console.error('Premium data load failed:', e);
    }
  };

  // 未認証時はナビだけ初期化
  initNav();

  function initNav() {
    const navInner=document.getElementById('mainNav');
    const btnL=document.getElementById('navScrollLeft');
    const btnR=document.getElementById('navScrollRight');
    function updateScrollBtns(){
      if(!navInner||!btnL||!btnR)return;
      btnL.classList.toggle('hidden',navInner.scrollLeft<=4);
      btnR.classList.toggle('hidden',navInner.scrollLeft+navInner.clientWidth>=navInner.scrollWidth-4);
    }
    if(navInner){
      navInner.addEventListener('scroll',updateScrollBtns);
      window.addEventListener('resize',updateScrollBtns);
      setTimeout(updateScrollBtns,100);
    }
    if(btnL)btnL.addEventListener('click',()=>{navInner.scrollBy({left:-200,behavior:'smooth'});});
    if(btnR)btnR.addEventListener('click',()=>{navInner.scrollBy({left:200,behavior:'smooth'});});
    document.querySelectorAll('.nav-item').forEach(n => n.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach(x => x.classList.remove('active'));
      n.classList.add('active');
      tab = n.dataset.tab;
      document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
      document.getElementById('sec-'+tab).classList.add('active');
      n.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
      render();
    }));
  }
  function render() {
    const fns = { exec:rExec, market:rMarket, financial:rFinancial, dupont:rDupont, shareholder:rSharehold, monthly:rMonthly, detail:rDetail, macro:rMacro };
    if (fns[tab]) fns[tab]();
  }

  // ============ 01 EXECUTIVE SUMMARY ============
  function rExec() {
    const el = g('sec-exec');
    const tm=companies.reduce((s,c)=>s+c.marketCap,0), tr=companies.reduce((s,c)=>s+c.revenue,0);
    const ts=companies.reduce((s,c)=>s+c.stores,0);
    const aOP=avg(companies,'opMargin'), aROE=avg(companies,'roe'), aPER=avg(companies,'per');
    const aYut=(companies.reduce((s,c)=>s+c.totalYutaiYield,0)/companies.length).toFixed(2);
    const topMC=topN(companies,'marketCap',3), topOP=topN(companies,'opMargin',3), topROE=topN(companies,'roe',3);
    // TOPIX比較の動的算出(データ連動・ハードコード排除)
    const rrV=companies.map(c=>c.relativeReturn).filter(v=>v!=null), exAbove=rrV.filter(v=>v>0).length, exValid=rrV.length;
    const restPt=Array.isArray(RESTAURANT_INDEX_MONTHLY)&&RESTAURANT_INDEX_MONTHLY.length?RESTAURANT_INDEX_MONTHLY[RESTAURANT_INDEX_MONTHLY.length-1]:null;
    const topixPt=Array.isArray(TOPIX_MONTHLY)&&TOPIX_MONTHLY.length?TOPIX_MONTHLY[TOPIX_MONTHLY.length-1]:null;
    const exPt=(restPt!=null&&topixPt!=null)?+(restPt-topixPt).toFixed(1):null;
    const topPerf=[...companies].filter(c=>c.stockReturn1Y!=null).sort((a,b)=>b.stockReturn1Y-a.stockReturn1Y).slice(0,3);

    el.innerHTML = `
      ${secH('01','Executive Summary','外食セクター全体概況と主要指標ハイライト')}
      <div class="commentary">
        <strong>セクター概況 (${DATA_AS_OF.stockPrice}基準):</strong> 対象${companies.length}社の合計時価総額は<strong>${fmtBil(tm)}</strong>、売上高合計<strong>${fmtBil(tr)}</strong>。
        外食セクター指数(時価総額上位15社・加重)は直近(${DATA_AS_OF.topixPeriod})で${exPt!=null?`TOPIXを<strong>${exPt>=0?'+':''}${exPt}pt</strong>${exPt>=0?'上回る':'下回る'}`:'<strong>TOPIX比 —</strong>'}。一方でTOPIX(1年騰落率 ${nv(TOPIX_RETURN_1Y,'%')})を上回った企業は${exValid}社中<strong>${exAbove}社</strong>に留まり、上昇は大型株に集中(企業間格差が顕著)。
        ${topPerf.length?`1年騰落率の上位は${topPerf.map(c=>`${shortName(c.name)}(${sv(c.stockReturn1Y)})`).join('、')}。`:''}<br><br>
        <strong>足元のリスク要因:</strong>
        (1) <strong>中東情勢の緊迫化</strong>(イラン紛争)に伴う原油価格上昇が輸送費・包材費・光熱費のコスト増として波及するリスク、
        (2) <strong>郊外型ファミリーレストラン</strong>における客数減少トレンド(人口動態変化・デリバリーサービスの代替進行)、
        (3) <strong>米価指数は前年比+38%と急騰</strong>(2015年基準指数120→166、小売店頭価格ベースでは約1.8倍)し値上げ転嫁にも限界が意識される局面。
        最低賃金は2024年に<strong>全国加重平均1,055円</strong>(${DATA_AS_OF.minimumWage})に到達し、パート比率の高い外食産業では人件費率30-40%×賃金上昇率約5%により、<strong>営業利益の数%〜十数%</strong>の下押し圧力と試算される(自社推計、パート比率・業態により幅あり)。<br>
        <strong>企業選別の提言:</strong> (1)ブランド力に裏打ちされた<strong>価格転嫁力</strong>、(2)<strong>海外展開</strong>による為替ヘッジ効果と成長余地、(3)<strong>DX・省人化投資</strong>によるコスト構造改革 の3軸を重視。
      </div>
      <div class="kpi-grid">
        ${kpi('対象企業数',companies.length+'社','','c-navy')}
        ${kpi('時価総額合計',fmtBil(tm),'','c-navy')}
        ${kpi('売上高合計',fmtBil(tr),'','c-gold')}
        ${kpi('平均営業利益率',aOP+'%','','c-green')}
        ${kpi('平均ROE',aROE+'%','セクター平均','c-gold')}
        ${kpi('平均PER',aPER+'倍','FY+1ベース','c-navy')}
        ${kpi('vs TOPIX(加重指数)',exPt!=null?`${exPt>=0?'+':''}${exPt}pt`:'-','','c-navy')}
        ${kpi('総店舗数',ts.toLocaleString(),'','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 百億円 / ${DATA_AS_OF.stockPrice}基準</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">営業利益率 vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / セグメント色分け</div><div class="chart-area tall"><canvas id="exBub"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">セグメント別 時価総額構成比</div><div class="chart-area"><canvas id="exPie"></canvas></div></div>
        <div class="chart-panel">
          <div class="chart-panel-title">主要ランキング</div>
          <div style="display:grid;gap:10px;padding-top:8px;">
            ${rankCard('時価総額',topMC,c=>fmtBil(c.marketCap))}
            ${rankCard('営業利益率',topOP,c=>c.opMargin+'%')}
            ${rankCard('ROE',topROE,c=>c.roe+'%')}
          </div>
        </div>
      </div>`;
    dc(['exMC','exPie','exBub']);
    const s15=[...companies].sort((a,b)=>b.marketCap-a.marketCap).slice(0,15);
    mc('exMC','bar',{labels:s15.map(c=>shortName(c.name)),datasets:[{data:s15.map(c=>Math.round(c.marketCap/10000*10)/10),backgroundColor:s15.map(c=>SC[c.segment]||'#777'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:v=>v.toLocaleString()+'百億'}}});
    const segMC={}; companies.forEach(c=>{segMC[c.segment]=(segMC[c.segment]||0)+c.marketCap;});
    mc('exPie','doughnut',{labels:Object.keys(segMC).map(k=>SEGMENTS[k]),datasets:[{data:Object.values(segMC),backgroundColor:Object.keys(segMC).map(k=>SC[k]),borderWidth:0}]},{plugins:{legend:{position:'right'},datalabels:{display:true,color:'#fff',font:{size:10,weight:600},formatter:(v,ctx)=>{const t=ctx.dataset.data.reduce((a,b)=>a+b,0);return(v/t*100).toFixed(1)+'%';}}}});
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('exBub','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.opMargin,y:c.roe,r:Math.max(4,Math.sqrt(c.marketCap/mx)*30),name:c.name})),backgroundColor:SC[seg]+'88',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'営業利益率 (%)'}},y:{title:{display:true,text:'ROE (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ${x.raw.x}% / ROE${x.raw.y}%`}}}});
  }

  // ============ 02 STOCK MARKET PERFORMANCE ============
  function rMarket() {
    const el = g('sec-market');
    const sorted = [...companies].sort((a,b)=>(b.relativeReturn??-1e9)-(a.relativeReturn??-1e9));
    const above = companies.filter(c=>c.relativeReturn!=null&&c.relativeReturn>0).length;
    const validN = companies.filter(c=>c.relativeReturn!=null).length;
    const restPt=Array.isArray(RESTAURANT_INDEX_MONTHLY)&&RESTAURANT_INDEX_MONTHLY.length?RESTAURANT_INDEX_MONTHLY[RESTAURANT_INDEX_MONTHLY.length-1]:null;
    const topixPt=Array.isArray(TOPIX_MONTHLY)&&TOPIX_MONTHLY.length?TOPIX_MONTHLY[TOPIX_MONTHLY.length-1]:null;
    const exPt=(restPt!=null&&topixPt!=null)?+(restPt-topixPt).toFixed(1):null;
    const baseM=Array.isArray(INDEX_MONTHS)&&INDEX_MONTHS.length?INDEX_MONTHS[0]:'起点';
    el.innerHTML = `
      ${secH('02','株式市場パフォーマンス','TOPIX対比の相対株価推移と企業別騰落率')}
      <div class="commentary">
        <strong>市場動向 (${DATA_AS_OF.topixPeriod}):</strong> 外食セクター指数(時価総額上位15社・加重)は<strong>${nv(restPt,'pt')}</strong>(TOPIX ${nv(topixPt,'pt')}、${baseM}=100基準)で、${exPt!=null?`TOPIXを<strong>${exPt>=0?'+':''}${exPt}pt</strong>${exPt>=0?'上回る':'下回る'}`:'TOPIX比 —'}。
        ただし個別ではTOPIX(1年騰落率 ${nv(TOPIX_RETURN_1Y,'%')})を上回った企業は<strong>${above}社</strong>(有効${validN}社中)に留まり、多くの企業はTOPIXを下回った(上昇の大型株集中)。<br>
        <span style="font-size:0.75rem;color:#888;">※ 外食セクター指数の定義: 弊社収録${companies.length}社のうち更新時点の<strong>時価総額上位15社</strong>を構成銘柄とする時価総額加重指数(${baseM}=100基準、構成15社で収録全社の時価総額の約8割をカバー)。上位企業への絞り込みは、小型株の流動性に起因するノイズを抑えながらセクターの実勢を代表させるための弊社選定基準による。構成銘柄は各回の更新時に時価総額基準で自動的に見直される。</span>
      </div>
      <div class="kpi-grid">
        ${kpi('外食セクター',nv(restPt,'pt'),'加重指数','c-navy')}
        ${kpi('TOPIX',nv(topixPt,'pt'),'同期間','c-navy')}
        ${kpi('超過(指数)',exPt!=null?`${exPt>=0?'+':''}${exPt}pt`:'-','','c-navy')}
        ${kpi('TOPIX超過企業',above+'/'+validN+'社','1年騰落率','c-gold')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">TOPIX-17セクター別パフォーマンス比較</div><div class="chart-panel-sub">直近1年騰落率 / 100基準</div><div class="chart-area tall"><canvas id="mkT17"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">外食産業 vs TOPIX 推移</div><div class="chart-panel-sub">時価総額上位15社・月次指数 (起点=100) / ${DATA_AS_OF.topixPeriod}</div><div class="chart-area tall"><canvas id="mkIdx"></canvas></div></div>
      </div>
      <div class="chart-row single">
        <div class="chart-panel"><div class="chart-panel-title">企業別 株価騰落率 (1年)</div><div class="chart-panel-sub">紺色=TOPIX超過 / 赤色=TOPIX未達 / 赤点線=TOPIX (${TOPIX_RETURN_1Y}%)</div><div class="chart-area tall"><canvas id="mkStocks"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">企業別パフォーマンス一覧</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">株価: ${DATA_AS_OF.stockPrice} / PER・PBR・ROE: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>決算期</th><th>セグメント</th><th>株価</th><th>時価総額(百億円)</th><th>1Y騰落率</th><th>vs TOPIX</th><th>PER</th><th>PBR</th><th>ROE</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem;white-space:nowrap">${c.fiscalYear||'-'}</td><td><span class="badge">${SEGMENTS[c.segment]}</span></td><td>${c.stockPrice.toLocaleString()}</td><td>${Math.round(c.marketCap/10000).toLocaleString()}</td><td class="${c.stockReturn1Y==null?'':c.stockReturn1Y>=0?'pos':'neg'}">${sv(c.stockReturn1Y)}</td><td class="${c.relativeReturn==null?'':c.relativeReturn>=0?'pos':'neg'}" style="font-weight:600">${sv(c.relativeReturn)}</td><td>${nv(c.per,'','f1')}</td><td>${nv(c.pbr,'','f1')}</td><td>${nv(c.roe,'%')}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);
    dc(['mkT17','mkIdx','mkStocks']);
    mc('mkT17','bar',{labels:TOPIX17_SECTORS.map(s=>s.name),datasets:[{data:TOPIX17_SECTORS.map(s=>s.point-100),backgroundColor:TOPIX17_SECTORS.map(s=>s.name==='Restaurant'?'#1a2d4f':s.name==='TOPIX'?'#9b8b6e':'#c8c4bb'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:v=>(v>=0?'+':'')+v.toFixed(1)+'pt'}}});
    const ml=INDEX_MONTHS||['24/05','24/06','24/07','24/08','24/09','24/10','24/11','24/12','25/01','25/02','25/03','25/04','25/05'];
    mc('mkIdx','line',{labels:ml,datasets:[{label:'外食産業',data:RESTAURANT_INDEX_MONTHLY,borderColor:'#1a2d4f',backgroundColor:'rgba(26,45,79,0.08)',fill:true,tension:0.3,pointRadius:4,borderWidth:2},{label:'TOPIX',data:TOPIX_MONTHLY,borderColor:'#9b8b6e',borderDash:[5,3],fill:false,tension:0.3,pointRadius:3,borderWidth:2}]},{});
    const hasAnnotation = Chart.registry?.plugins?.get?.('annotation') || (typeof window !== 'undefined' && window['chartjs-plugin-annotation']);
    mc('mkStocks','bar',{labels:sorted.map(c=>shortName(c.name)),datasets:[{data:sorted.map(c=>c.stockReturn1Y),backgroundColor:sorted.map(c=>c.stockReturn1Y>=TOPIX_RETURN_1Y?'#1a2d4f':'rgba(181,58,58,0.6)'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},annotation:{annotations:{topix:{type:'line',xMin:TOPIX_RETURN_1Y,xMax:TOPIX_RETURN_1Y,borderColor:'#b53a3a',borderWidth:2,borderDash:[4,4],label:{display:true,content:'TOPIX '+TOPIX_RETURN_1Y+'%',position:'start'}}}}}});
  }

  // ============ 03 FINANCIAL ============
  function rFinancial() {
    const el = g('sec-financial');
    const sorted=[...companies].sort((a,b)=>b.revenue-a.revenue);
    el.innerHTML = `
      ${secH('03','財務指標・バリュエーション','売上高・利益率・PER/PBRの横断比較')}
      <div class="commentary">
        <strong>バリュエーション分析 (${DATA_AS_OF.financials}ベース):</strong> セクター平均PERは${avg(companies,'per')}倍、PBRは${avg(companies,'pbr')}倍。
        全企業がPBR1倍超のバリュエーションを獲得しており、高PBR・高ROEのグロース株セクター。
        PBR-ROE間にはR²=0.53の正の相関が確認され、<strong>ROE改善が株価評価に直結</strong>する構造。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">売上高 vs 営業利益</div><div class="chart-panel-sub">TOP15 / 百万円</div><div class="chart-area tall"><canvas id="fnRP"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">PBR vs ROE マトリクス</div><div class="chart-panel-sub">4象限分析 / バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="fnPBR"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">PER vs PBR</div><div class="chart-panel-sub">左下=割安ゾーン</div><div class="chart-area"><canvas id="fnPERPBR"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">営業利益率分布</div><div class="chart-area"><canvas id="fnHist"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">財務指標一覧 (売上高順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">財務データ: ${DATA_AS_OF.financials} / 株価: ${DATA_AS_OF.stockPrice}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>決算期</th><th>売上高</th><th>営業利益</th><th>純利益</th><th>営業利益率</th><th>純利益率</th><th>PER</th><th>PBR</th><th>ROE</th><th>D/E</th><th>Net D/E</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem;white-space:nowrap">${c.fiscalYear||'-'}</td><td>${c.revenue.toLocaleString()}</td><td>${c.opProfit.toLocaleString()}</td><td>${c.netProfit.toLocaleString()}</td><td class="${c.opMargin>=8?'pos':c.opMargin<3?'neg':''}">${c.opMargin}%</td><td>${c.netMargin}%</td><td>${nv(c.per,'','f1')}</td><td>${nv(c.pbr,'','f1')}</td><td class="${c.roe!=null&&c.roe>=15?'pos':''}">${nv(c.roe,'%')}</td><td>${nv(c.deRatio,'','f2')}</td><td>${nv(c.netDeRatio,'','f2')}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el); dc(['fnRP','fnPBR','fnPERPBR','fnHist']);
    const t15=sorted.slice(0,15);
    mc('fnRP','bar',{labels:t15.map(c=>shortName(c.name)),datasets:[{label:'売上高',data:t15.map(c=>c.revenue),backgroundColor:'rgba(26,45,79,0.6)',borderWidth:0},{label:'営業利益',data:t15.map(c=>c.opProfit),backgroundColor:'rgba(45,122,79,0.6)',borderWidth:0}]},{indexAxis:'y'});
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('fnPBR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.roe,y:c.pbr,r:Math.max(4,Math.sqrt(c.marketCap/mx)*28),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ROE${x.raw.x}% / PBR${x.raw.y}倍`}}}});
    mc('fnPERPBR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.per,y:c.pbr,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'PER (倍)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: PER${x.raw.x}倍 / PBR${x.raw.y}倍`}}}});
    const bins=[0,2,4,6,8,10,15,30]; const hist=bins.slice(0,-1).map((_,i)=>companies.filter(c=>c.opMargin>=bins[i]&&c.opMargin<bins[i+1]).length);
    mc('fnHist','bar',{labels:bins.slice(0,-1).map((b,i)=>`${b}-${bins[i+1]}%`),datasets:[{data:hist,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});
  }

  // ============ 04 DUPONT DECOMPOSITION ============
  function rDupont() {
    const el = g('sec-dupont');
    const sorted=[...companies].sort((a,b)=>b.roe-a.roe);
    // エー・ピーHD(3175)は自己資本が僅少でROE・レバレッジが発散する外れ値のため、グラフと平均から除外(一覧表には掲載)
    const dpSorted=sorted.filter(c=>c.code!=='3175');
    const dpComps=companies.filter(c=>c.code!=='3175');
    // 平均値はROEが±100%超に発散する銘柄(KOZO HD・SANKO等)も除外して算出
    const dpAvgComps=companies.filter(c=>c.roe!=null&&Math.abs(c.roe)<100);
    el.innerHTML = `
      ${secH('04','DuPont分解','ROE = 売上高純利益率 × 総資産回転率 × 財務レバレッジ')}
      <div class="commentary">
        <strong>DuPont分解:</strong> ROE平均は<strong>${avg(dpAvgComps,'roe')}%</strong>、売上高純利益率平均<strong>${avg(dpAvgComps,'netMargin')}%</strong>、
        総資産回転率平均<strong>${avg(dpAvgComps,'assetTurnover')}x</strong>、財務レバレッジ平均<strong>${avg(dpAvgComps,'leverage')}x</strong>。
        エターナルホスピタリティG(ROE${companies.find(c=>c.code==='3193')?.roe}%)やクリエイト・レストランツHD等は高レバレッジが高ROEに寄与。
        一方、コメダHDは純利益率${companies.find(c=>c.code==='3543')?.netMargin}%と突出するが回転率は0.5xと低い(FCモデルの特性)。
        ※自己資本僅少でROEが発散するエー・ピーHD(212%)はグラフ・平均から、ROEが±100%を超えるKOZO HD(-213%)・SANKO MARKETING FOODS(-304%)は平均から除外(いずれも一覧表には掲載)。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">ROE(FY) DuPont分解</div><div class="chart-panel-sub">ROE順 / 3要素の寄与 / エー・ピーHDは外れ値のため除外</div><div class="chart-area tall"><canvas id="dpBar"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">純利益率 vs 総資産回転率</div><div class="chart-panel-sub">バブルサイズ=ROE / エー・ピーHDは外れ値のため除外</div><div class="chart-area tall"><canvas id="dpScatter"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">DuPont分解一覧 (ROE順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">財務データ: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>決算期</th><th>ROE</th><th>純利益率</th><th>総資産回転率</th><th>財務レバレッジ</th><th>D/E</th><th>Net D/E</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row ${c.code==='8163'?'row-hl':''}" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem;white-space:nowrap">${c.fiscalYear||'-'}</td><td class="${c.roe!=null&&c.roe>=15?'pos':''}" style="font-weight:700">${nv(c.roe,'%')}</td><td>${c.netMargin}%</td><td>${nv(c.assetTurnover,'x','f1')}</td><td>${nv(c.leverage,'x','f1')}</td><td>${nv(c.deRatio,'','f2')}</td><td>${nv(c.netDeRatio,'','f2')}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el); dc(['dpBar','dpScatter']);
    mc('dpBar','bar',{labels:dpSorted.map(c=>shortName(c.name)),datasets:[{label:'純利益率(%)',data:dpSorted.map(c=>c.netMargin),backgroundColor:'#1a2d4f'},{label:'回転率(x)',data:dpSorted.map(c=>c.assetTurnover*5),backgroundColor:'#9b8b6e'},{label:'レバレッジ(x)',data:dpSorted.map(c=>c.leverage),backgroundColor:'#5a7fa8'}]},{indexAxis:'y',scales:{x:{stacked:false}},plugins:{legend:{position:'top'}}});
    const maxROE=Math.max(...dpComps.map(c=>c.roe));
    mc('dpScatter','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:dpComps.filter(c=>c.segment===seg).map(c=>({x:c.netMargin,y:c.assetTurnover,r:Math.max(4,c.roe/maxROE*30),name:c.name,roe:c.roe})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'売上高純利益率 (%)'}},y:{title:{display:true,text:'総資産回転率 (x)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 純利益率${x.raw.x}% / 回転率${x.raw.y}x / ROE${x.raw.roe}%`}}}});
  }

  // ============ 05 SHAREHOLDER ============
  function rSharehold() {
    const el = g('sec-shareholder');
    const ys=[...companies].sort((a,b)=>b.totalYutaiYield-a.totalYutaiYield);
    const is=[...companies].sort((a,b)=>b.individualRatio-a.individualRatio);
    el.innerHTML = `
      ${secH('05','株主還元・投資指標','配当+優待利回り・個人投資家比率の横断分析')}
      <div class="commentary">
        <strong>株主還元 (${DATA_AS_OF.stockPrice}基準):</strong> 優待総利回りでは<strong>${ys[0].name}</strong>が${ys[0].totalYutaiYield.toFixed(2)}%でトップ。
        個人投資家比率と優待総利回りには<strong>正の相関</strong>が認められ、優待制度が株主構成に直結する構造。<br>
        ※優待利回り = 優待価値(年間) ÷ 最低投資額(株価×最低必要株数)。最低必要株数は各社優待制度に準拠。
      </div>
      <div class="kpi-grid">
        ${kpi('優待総利回りTOP',ys[0].totalYutaiYield.toFixed(2)+'%',ys[0].name,'c-green')}
        ${kpi('個人比率TOP',is[0].individualRatio+'%',is[0].name,'c-gold')}
        ${kpi('平均配当利回り',avg(companies,'dividendYield')+'%','','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">優待総利回りランキング</div><div class="chart-panel-sub">配当+優待 / TOP20</div><div class="chart-area tall"><canvas id="shBar"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">優待総利回り vs 個人投資家比率</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="shScat"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">株主還元指標一覧</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">株価: ${DATA_AS_OF.stockPrice}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>株価</th><th>最低株数</th><th>最低投資額</th><th>配当利回り</th><th>優待価値</th><th>優待利回り</th><th>総利回り</th><th>個人比率</th></tr></thead>
        <tbody>${ys.map(c=>{const inv=c.stockPrice*c.yutaiMinShares;return `<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td>${c.stockPrice.toLocaleString()}</td><td>${c.yutaiMinShares}株</td><td>${inv.toLocaleString()}円</td><td>${c.dividendYield.toFixed(2)}%</td><td>${c.yutaiValue.toLocaleString()}円</td><td>${c.yutaiYield.toFixed(2)}%</td><td class="${c.totalYutaiYield>=3?'pos':''}" style="font-weight:700">${c.totalYutaiYield.toFixed(2)}%</td><td>${nv(c.individualRatio,'%')}</td></tr>`;}).join('')}</tbody>
      </table></div></div>`;
    bindRows(el); dc(['shBar','shScat']);
    const y20=ys.slice(0,20);
    mc('shBar','bar',{labels:y20.map(c=>shortName(c.name)),datasets:[{label:'配当利回り',data:y20.map(c=>c.dividendYield),backgroundColor:'#2d7a4f'},{label:'優待利回り',data:y20.map(c=>c.yutaiYield),backgroundColor:'#9b8b6e'}]},{indexAxis:'y',scales:{x:{stacked:true},y:{stacked:true}}});
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('shScat','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.totalYutaiYield,y:c.individualRatio,r:Math.max(4,Math.sqrt(c.marketCap/mx)*22),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'優待総利回り (%)'}},y:{title:{display:true,text:'個人投資家比率 (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ${x.raw.x.toFixed(2)}% / 個人${x.raw.y}%`}}}});
  }

  // ============ 06 MONTHLY KPI ============
  // 複数年度データ取得ヘルパー
  // 企業データに monthlyHistory があれば複数年度表示可能
  // monthlyHistory: { 'FY2023': { revenue:[12], profit:[12], costRatio:[12], sameStoreSales:[12] }, ... }
  function getAvailableFYs() {
    const fys = new Set(['latest']);
    companies.forEach(c => {
      if (c.monthlyHistory) Object.keys(c.monthlyHistory).forEach(fy => fys.add(fy));
    });
    return [...fys].sort();
  }

  function getMonthlyData(company, field, fy) {
    if (fy === 'latest') return company[field] || [];
    if (company.monthlyHistory && company.monthlyHistory[fy]) {
      const fieldMap = { monthlyRevenue:'revenue', monthlyProfit:'profit', costRatio:'costRatio', sameStoreSales:'sameStoreSales' };
      return company.monthlyHistory[fy][fieldMap[field] || field] || [];
    }
    return [];
  }

  function getMonthLabels(selectedFYs) {
    if (selectedFYs.length <= 1) return MONTHS;
    // 複数年度: 年度+月のラベルを生成
    const labels = [];
    selectedFYs.forEach(fy => {
      const year = fy === 'latest' ? DATA_AS_OF.sameStoreSales.match(/(\d{4})年/)?.[1] || '' : fy.replace('FY','');
      MONTHS.forEach(m => labels.push(`${year}/${m}`));
    });
    return labels;
  }

  function getConcatData(company, field, selectedFYs) {
    if (selectedFYs.length <= 1) return getMonthlyData(company, field, selectedFYs[0] || 'latest');
    return selectedFYs.flatMap(fy => getMonthlyData(company, field, fy));
  }

  function rMonthly() {
    const el = g('sec-monthly');
    if (!mCodes.length) mCodes = companies.slice(0,5).map(c=>c.code);
    const vis = mSeg==='ALL' ? companies : companies.filter(c=>c.segment===mSeg);
    const availFYs = getAvailableFYs();
    const hasMultiYear = availFYs.length > 1;
    const latestYear = DATA_AS_OF.sameStoreSales.match(/(\d{4})年\d+月～(\d{4})年/);
    const fyLabel = latestYear ? `${latestYear[1]}年度` : '直近1年';
    el.innerHTML = `
      ${secH('06','月次KPI横断比較','複数企業・セグメント横断で売上高・利益・原価率・既存店売上前年比を比較')}
      <div class="commentary">
        <strong>月次動向:</strong> セグメントフィルタと企業チップで複数企業を同時選択・比較可能。
        外食セクターは<strong>7-8月(夏休み)と12月(忘年会)</strong>にピーク、2月が底となる共通の季節性がある。
        既存店売上高データは<strong>各社IR・月次売上情報</strong>を参照。
        業界平均(破線)は主要30社超の既存店売上高単純平均。24年10月の業界平均は<strong>前年同月比+4.6%</strong>とやや鈍化傾向。
      </div>
      <div class="chip-wrap"><div class="chip-label">表示期間</div>
        <div class="chip-select" id="mFYC">
          <div class="chip ${mFY==='latest'?'selected':''}" data-fy="latest">${fyLabel}</div>
          ${availFYs.filter(f=>f!=='latest').map(fy=>`<div class="chip ${mFY===fy?'selected':''}" data-fy="${fy}">${fy.replace('FY','')+'年度'}</div>`).join('')}
          ${hasMultiYear?`<div class="chip ${mFY==='all'?'selected':''}" data-fy="all">全期間 (最大3年)</div>`:''}
        </div>
        ${!hasMultiYear?'<div style="font-size:0.72rem;color:#999;margin-top:4px;">※ 過去年度のデータが追加されると、ここで期間を切り替えて最長3年分を表示できます</div>':''}
      </div>
      <div class="chip-wrap"><div class="chip-label">セグメントフィルタ</div>
        <div class="chip-select" id="mSegC">
          <div class="chip ${mSeg==='ALL'?'selected':''}" data-seg="ALL">全セグメント</div>
          ${Object.entries(SEGMENTS).map(([k,v])=>`<div class="chip ${mSeg===k?'selected':''}" data-seg="${k}">${v}</div>`).join('')}
        </div>
      </div>
      <div class="chip-wrap"><div class="chip-label">比較企業を選択 (複数可)</div>
        <div class="chip-select" id="mCompC">${vis.map(c=>`<div class="chip ${mCodes.includes(c.code)?'selected':''}" data-code="${c.code}">${shortName(c.name)}</div>`).join('')}</div>
      </div>
      <div id="mKpi" class="kpi-grid"></div>
      <div class="chart-row"><div class="chart-panel"><div class="chart-panel-title">月次売上高推移</div><div class="chart-area"><canvas id="mRev"></canvas></div></div><div class="chart-panel"><div class="chart-panel-title">月次利益推移</div><div class="chart-area"><canvas id="mProf"></canvas></div></div></div>
      <div class="chart-row"><div class="chart-panel"><div class="chart-panel-title">原価率推移</div><div class="chart-area"><canvas id="mCost"></canvas></div></div><div class="chart-panel"><div class="chart-panel-title">既存店売上高前年比</div><div class="chart-area"><canvas id="mSSS"></canvas></div></div></div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">既存店売上高前年比 ヒートマップ</div></div><div class="table-scroll"><table class="heatmap-table" id="hmSSS"></table></div></div>
      <div class="commentary" style="margin-top:24px;font-size:0.78rem;">
        <strong>データ方法論:</strong><br>
        ・<strong>対象期間:</strong> ${DATA_AS_OF.sameStoreSales}(4月始まり会計年度)<br>
        ・<strong>既存店売上高の採用ブランド:</strong> 複数ブランド展開企業は、売上高構成比が最大のブランドを代表値として使用。
        ゼンショーHD→すき家、トリドールHD→丸亀製麺、FOOD&LIFE→スシロー、物語コーポレーション→焼肉きんぐ、
        コメダHD→コメダ珈琲店、あみやき亭→焼肉部門、ペッパーフード→いきなり！ステーキ、
        ハイデイ日高→日高屋、鳥貴族HD→鳥貴族(エターナルホスピタリティG)<br>
        ・<strong>月次売上高・利益:</strong> 各社IR開示の月次売上速報を百万円単位で記載。非開示企業は四半期実績を月次按分して推計<br>
        ・<strong>原価率(食材費率):</strong> 有価証券報告書の売上原価÷売上高で算出。月次変動は季節パターンから推計<br>
        ・<strong>業界平均(破線):</strong> 外食主要チェーンの既存店売上高前年比の単純平均(各社IR月次速報ベース)<br>
        ・<strong>更新頻度:</strong> 各社月次速報は翌月中旬に開示されるため、毎月下旬にデータ更新を実施
      </div>`;
    g('mFYC').querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{mFY=ch.dataset.fy;updateMonthlyCharts();g('mFYC').querySelectorAll('.chip').forEach(x=>x.classList.toggle('selected',x.dataset.fy===mFY));}));
    g('mSegC').querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{mSeg=ch.dataset.seg;mCodes=[];rMonthly();}));
    g('mCompC').querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{const c=ch.dataset.code;mCodes=mCodes.includes(c)?mCodes.filter(x=>x!==c):[...mCodes,c];updateMonthlyCharts();}));
    updateMonthlyCharts();
  }
  function updateMonthlyCharts() {
    const sel=companies.filter(c=>mCodes.includes(c.code));
    g('mCompC')?.querySelectorAll('.chip').forEach(ch=>ch.classList.toggle('selected',mCodes.includes(ch.dataset.code)));
    if(!sel.length) return;
    dc(['mRev','mProf','mCost','mSSS']);

    // 表示期間に応じたデータ取得
    const availFYs = getAvailableFYs();
    let selectedFYs;
    if (mFY === 'all') {
      selectedFYs = availFYs.filter(f=>f!=='latest').concat(['latest']);
    } else {
      selectedFYs = [mFY];
    }
    const labels = getMonthLabels(selectedFYs);

    const ds=(field)=>sel.map((c,i)=>({label:shortName(c.name),data:getConcatData(c,field,selectedFYs),borderColor:P[i%P.length],fill:false,tension:0.3,pointRadius:3,borderWidth:2}));
    mc('mRev','line',{labels,datasets:ds('monthlyRevenue')},{});
    mc('mProf','line',{labels,datasets:ds('monthlyProfit')},{});
    mc('mCost','line',{labels,datasets:ds('costRatio')},{});

    // 既存店売上高: 業界平均も期間に合わせて年度別データを使用
    const getSectorAvg = (fy) => {
      if (fy === 'latest') return SECTOR_AVG_SSS;
      return SECTOR_AVG_SSS_HISTORY[fy] || SECTOR_AVG_SSS;
    };
    const sssAvgData = selectedFYs.flatMap(fy => getSectorAvg(fy));
    mc('mSSS','line',{labels,datasets:[...ds('sameStoreSales'),{label:'業界平均',data:sssAvgData,borderColor:'#999',borderDash:[6,3],fill:false,tension:0.3,pointRadius:0,borderWidth:2}]},{});
    // ヒートマップ: 現在選択中の年度のデータのみ
    if (selectedFYs.length <= 1) {
      heatmap('hmSSS',sel,'sameStoreSales',v=>v==null?'-':v.toFixed(1),v=>v==null?'':v>=105?'pos':v<100?'neg':'');
    } else {
      // 複数年度: カスタムヒートマップ
      const t=g('hmSSS');if(t){
        let h=`<thead><tr><th style="text-align:left">企業名</th>${labels.map(m=>`<th style="font-size:0.65rem">${m}</th>`).join('')}</tr></thead><tbody>`;
        sel.forEach(c=>{const v=getConcatData(c,'sameStoreSales',selectedFYs);h+=`<tr><td>${shortName(c.name)}</td>${v.map(x=>`<td class="${x==null?'':x>=105?'pos':x<100?'neg':''}">${x==null?'-':x.toFixed(1)}</td>`).join('')}</tr>`;});
        h+='</tbody>';t.innerHTML=h;
      }
    }
  }

  // ============ 07 DETAIL ============
  function rDetail() {
    const el = g('sec-detail');
    if(!selComp) selComp = companies[0];
    const c = selComp;
    const peers = companies.filter(x=>x.segment===c.segment&&x.code!==c.code);
    el.innerHTML = `
      ${secH('07','個別企業分析','選択企業の詳細指標と同業比較')}
      <div class="inline-filters"><span class="f-label">企業選択</span>
        <select id="dtSel">${companies.map(x=>`<option value="${x.code}" ${x.code===c.code?'selected':''}>${x.code} ${x.name}</option>`).join('')}</select>
      </div>
      <div style="font-family:'Noto Serif JP',serif;font-size:1.4rem;font-weight:700;color:var(--navy);margin-bottom:4px;">${c.name}</div>
      <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:8px;">${c.code} / ${SEGMENTS[c.segment]} / ${c.stores.toLocaleString()}店舗 / 決算期: ${c.fiscalYear||'-'}</div>
      <div class="chip-select" style="margin-bottom:24px;">${c.brands.map(b=>`<span class="chip selected">${b}</span>`).join('')}</div>
      <div class="kpi-grid">
        ${kpi('株価',c.stockPrice.toLocaleString()+'円',DATA_AS_OF.stockPrice,'c-navy')}
        ${kpi('時価総額',Math.round(c.marketCap/10000).toLocaleString()+'百億円','','c-navy')}
        ${kpi('営業利益率',c.opMargin+'%','','c-green')}
        ${kpi('ROE',c.roe+'%','','c-gold')}
        ${kpi('PER / PBR',nv(c.per,'','f1')+' / '+nv(c.pbr,'','f1'),'','c-navy')}
        ${kpi('優待総利回り',c.totalYutaiYield.toFixed(2)+'%','配当'+c.dividendYield.toFixed(2)+'%+優待'+c.yutaiYield.toFixed(2)+'%','c-green')}
        ${kpi('vs TOPIX (1Y)',sv(c.relativeReturn),'','c-gold')}
        ${kpi('個人比率',nv(c.individualRatio,'%'),'','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">月次売上高</div><div class="chart-area short"><canvas id="dtR"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">既存店売上前年比</div><div class="chart-area short"><canvas id="dtS"></canvas></div></div>
      </div>
      ${peers.length?`<div class="table-panel"><div class="table-header"><div class="table-header-title">同業比較 - ${SEGMENTS[c.segment]}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">企業名</th><th>決算期</th><th>売上高</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>総利回り</th><th>vs TOPIX</th></tr></thead>
        <tbody><tr class="row-hl"><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem">${c.fiscalYear||'-'}</td><td>${c.revenue.toLocaleString()}</td><td>${c.opMargin}%</td><td>${nv(c.roe,'%')}</td><td>${nv(c.per,'','f1')}</td><td>${c.totalYutaiYield.toFixed(2)}%</td><td>${sv(c.relativeReturn)}</td></tr>
        ${peers.map(p=>`<tr><td>${shortName(p.name)}</td><td style="font-size:0.75rem">${p.fiscalYear||'-'}</td><td>${p.revenue.toLocaleString()}</td><td>${p.opMargin}%</td><td>${nv(p.roe,'%')}</td><td>${nv(p.per,'','f1')}</td><td>${p.totalYutaiYield.toFixed(2)}%</td><td>${sv(p.relativeReturn)}</td></tr>`).join('')}
        </tbody></table></div></div>`:''}`;
    g('dtSel').addEventListener('change',e=>{selComp=companies.find(c=>c.code===e.target.value);rDetail();});
    dc(['dtR','dtS']);
    mc('dtR','bar',{labels:MONTHS,datasets:[{data:c.monthlyRevenue,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});
    const sssValid = c.sameStoreSales.filter(v=>v!=null);
    if(sssValid.length) mc('dtS','bar',{labels:MONTHS,datasets:[{data:c.sameStoreSales,backgroundColor:c.sameStoreSales.map(v=>v==null?'rgba(200,200,200,0.3)':v>=100?'rgba(45,122,79,0.5)':'rgba(181,58,58,0.5)'),borderWidth:0}]},{scales:{y:{min:Math.min(...sssValid)-3}},plugins:{legend:{display:false}}});
    else g('dtS').parentElement.innerHTML='<p style="text-align:center;color:#999;padding:2em">月次既存店データ非開示</p>';
  }

  // ============ APPENDIX: MACRO ============
  function rMacro() {
    const el = g('sec-macro');
    const md = MACRO_DATA;
    el.innerHTML = `
      ${secH('A','マクロ指標付録','家計支出・原材料価格・人件費等の外部環境データ')}
      <div class="commentary">
        <strong>外部環境:</strong> 米価は2025年平均で<strong>前年比+67.5%</strong>(CPI米類)と一段と急騰し、指数は<strong>278</strong>(2015年10月=100基準)に到達(東京都区部の店頭価格でコシヒカリ5kg 4,928円)。
        最低賃金は2025年度改定で<strong>全国加重平均1,121円</strong>(前年度比+66円、目安制度開始以降で最大の引き上げ)に到達し、人件費上昇がオペレーションコストを圧迫。
        外食産業全体の売上は2025年も前年比107.3%(JF会員社全店ベース)と拡大が続くが、押し上げの主因は客単価上昇(104.3%)で、客数の伸び(102.9%)には頭打ち感。
        家計の外食支出(二人以上世帯)は2025年に<strong>年間198,759円</strong>(前年比+6.0%)と増加が続く一方、<strong>飲酒代(19,498円)はCovid前(2019年約2.0万円)水準どまり</strong>。
        市場規模(狭義)は2023年に24.2兆円まで回復したが、なお2019年比△8.1%。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">主要食品価格指数推移</div><div class="chart-panel-sub">2015年10月=100 / 出所: 農畜産業振興機構, 農林水産省 (2025年は総務省CPI品目別前年比で延長。輸入鶏肉はCPI区分なしのため2024年まで)</div><div class="chart-area tall"><canvas id="maFood"></canvas></div><div style="font-size:0.68rem;color:#999;margin-top:8px;line-height:1.6;">${md.foodPrices.absNote}</div></div>
        <div class="chart-panel"><div class="chart-panel-title">二人以上世帯 年間飲食費推移</div><div class="chart-panel-sub">1世帯当たり年間支出金額(万円) / 出所: 総務省「家計調査」品目分類(全国・二人以上の世帯)</div><div class="chart-area tall"><canvas id="maHH"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">外食産業主要指標推移</div><div class="chart-panel-sub">2019/12=100 / 出所: 日本フードサービス協会 月次「外食産業市場動向調査」(全店ベース。協会公表値は前年比のため2019/12起点の累積指数として表示。金額の絶対水準は下の市場規模チャート参照)</div><div class="chart-area"><canvas id="maInd"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">最低賃金推移 (全国加重平均)</div><div class="chart-panel-sub">各年度改定額(円) / 出所: 厚生労働省</div><div class="chart-area"><canvas id="maWage"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">外食産業市場規模推移 (絶対額)</div><div class="chart-panel-sub">狭義・外食産業計(兆円、消費税込) / 出所: 食の安全・安心財団「外食産業市場規模推計」。2024・2025年推計は未公表(2026年7月時点、公表され次第更新)</div><div class="chart-area"><canvas id="maSize"></canvas></div></div>
      </div>`;
    dc(['maFood','maHH','maInd','maWage','maSize']);
    const fp=md.foodPrices;
    mc('maFood','line',{labels:fp.years,datasets:Object.entries(fp.items).map(([k,v],i)=>({label:k,data:v,borderColor:P[i%P.length],fill:false,tension:0.3,pointRadius:3,borderWidth:2}))},{});
    const hs=md.householdSpending;
    mc('maHH','bar',{labels:hs.years,datasets:Object.entries(hs.categories).map(([k,v],i)=>({label:k,data:v,backgroundColor:P[i%P.length]+'cc'}))},{scales:{x:{stacked:true},y:{stacked:true,title:{display:true,text:'万円'}}}});
    const ms=md.marketSize;
    mc('maSize','bar',{labels:ms.years,datasets:[{label:'市場規模(兆円)',data:ms.values,backgroundColor:'rgba(155,139,110,0.55)',borderWidth:0}]},{plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'top',color:'#777',font:{size:10},formatter:v=>v.toFixed(1)+'兆円'}}});
    const ii=md.industryIndex;
    mc('maInd','line',{labels:ii.months,datasets:[{label:'売上高',data:ii.sales,borderColor:'#1a2d4f',tension:0.3,pointRadius:3,borderWidth:2,fill:false},{label:'店舗数',data:ii.stores,borderColor:'#9b8b6e',tension:0.3,pointRadius:3,borderWidth:2,fill:false},{label:'客数',data:ii.customers,borderColor:'#5a7fa8',tension:0.3,pointRadius:3,borderWidth:2,fill:false},{label:'客単価',data:ii.unitPrice,borderColor:'#2d7a4f',tension:0.3,pointRadius:3,borderWidth:2,fill:false}]},{});
    const lc=md.laborCost;
    mc('maWage','bar',{labels:lc.years,datasets:[{label:'最低賃金(円)',data:lc.minWage,backgroundColor:'rgba(26,45,79,0.6)',borderWidth:0}]},{plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'top',color:'#777',font:{size:10},formatter:v=>v+'円'}}});
  }

  // ============ HELPERS ============
  function g(id){return document.getElementById(id);}
  function mc(id,type,data,opts={}){
    const ctx=g(id);if(!ctx)return;
    fitBarHeight(ctx,data,opts);
    const base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};
    if(['bar','line','scatter','bubble'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};
    C[id]=new Chart(ctx,{type,data,options:dm(base,opts)});
  }
  /* 横棒グラフ(indexAxis:'y')の高さを項目数に合わせて確保する。
     .chart-area はモバイルで 240px / .tall で 300px（480px以下では 200/260px）に
     固定されるため、54社の企業別騰落率などは1本あたり5px未満に潰れ、
     Chart.js が目盛ラベルを大量に間引いて判読不能になっていた。
     モバイル幅でのみ、1項目あたり20pxを確保した高さへ引き上げる。
     .chart-area の高さは mobile-touch-font-fix.min.css が !important で
     指定しているため、インラインでも priority 'important' が必要。 */
  function fitBarHeight(canvas,data,opts){
    if(opts.indexAxis!=='y')return;
    if(window.innerWidth>768)return;
    const n=(data.labels||[]).length;
    if(n<=12)return;
    const area=canvas.closest('.chart-area');
    if(!area)return;
    area.style.setProperty('height',(n*20+60)+'px','important');
  }
  function dc(ids){ids.forEach(id=>{if(C[id]){C[id].destroy();delete C[id];}});}
  function dm(t,s){const o={...t};for(const k of Object.keys(s)){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]);else o[k]=s[k];}return o;}
  function secH(n,t,d){return `<div class="sec-header"><div class="sec-num">SECTION ${n}</div><div class="sec-title">${t}</div><div class="sec-desc">${d}</div></div>`;}
  function kpi(l,v,s,cls){return `<div class="kpi-card ${cls}"><div class="kpi-label">${l}</div><div class="kpi-value">${v}</div>${s?`<div class="kpi-sub">${s}</div>`:''}</div>`;}
  function rankCard(t,items,fn){return `<div class="ranking-card"><div class="ranking-title">${t}</div>${items.map((c,i)=>`<div class="ranking-row"><span><span class="ranking-num">${i+1}</span>${shortName(c.name)}</span><span style="font-weight:600">${fn(c)}</span></div>`).join('')}</div>`;}
  function heatmap(id,comps,field,fmt,cls){
    const t=g(id);if(!t)return;
    let h=`<thead><tr><th style="text-align:left">企業名</th>${MONTHS.map(m=>`<th>${m}</th>`).join('')}<th>平均</th></tr></thead><tbody>`;
    comps.forEach(c=>{const v=c[field],a=v.reduce((s,x)=>s+x,0)/12;h+=`<tr><td>${shortName(c.name)}</td>${v.map(x=>`<td class="${cls(x)}">${fmt(x)}</td>`).join('')}<td style="font-weight:700">${fmt(a)}</td></tr>`;});
    h+='</tbody>';t.innerHTML=h;
  }
  function bindRows(el){el.querySelectorAll('.clickable-row').forEach(tr=>tr.addEventListener('click',()=>{selComp=companies.find(c=>c.code===tr.dataset.code);document.querySelector('.nav-item[data-tab="detail"]').click();}));}

  // ============ DOWNLOAD FUNCTIONS ============
  function requireAuth() {
    if (!window.currentUser) { window.openModal('login'); return false; }
    return true;
  }

  window.downloadPDF = function() {
    if (!requireAuth()) return;
    window.print();
  };

  window.downloadCSV = function() {
    if (!requireAuth()) return;
    const BOM = '\uFEFF';
    const headers = [
      'コード','企業名','セグメント','決算期','株価(円)','時価総額(百万円)','発行済株式数(千株)',
      '売上高(百万円)','営業利益(百万円)','純利益(百万円)','営業利益率(%)','純利益率(%)',
      'PER(倍)','PBR(倍)','ROE(%)','D/Eレシオ','NetD/Eレシオ',
      '総資産回転率','財務レバレッジ',
      '個人投資家比率(%)','配当利回り(%)','優待価値(円/年)','優待最低株数','優待利回り(%)','総合利回り(%)',
      '店舗数','従業員数',
      '株価騰落率1Y(%)','vs TOPIX(pp)',
      '月次売上4月','月次売上5月','月次売上6月','月次売上7月','月次売上8月','月次売上9月',
      '月次売上10月','月次売上11月','月次売上12月','月次売上1月','月次売上2月','月次売上3月',
      '既存店売上4月','既存店売上5月','既存店売上6月','既存店売上7月','既存店売上8月','既存店売上9月',
      '既存店売上10月','既存店売上11月','既存店売上12月','既存店売上1月','既存店売上2月','既存店売上3月',
      '原価率4月','原価率5月','原価率6月','原価率7月','原価率8月','原価率9月',
      '原価率10月','原価率11月','原価率12月','原価率1月','原価率2月','原価率3月'
    ];
    const rows = companies.map(c => [
      c.code, c.name, SEGMENTS[c.segment]||c.segment, c.fiscalYear||'',
      c.stockPrice, c.marketCap, c.shares,
      c.revenue, c.opProfit, c.netProfit, c.opMargin, c.netMargin,
      c.per, c.pbr, c.roe, c.deRatio, c.netDeRatio,
      c.assetTurnover, c.leverage,
      c.individualRatio, c.dividendYield, c.yutaiValue, c.yutaiMinShares, c.yutaiYield, c.totalYutaiYield,
      c.stores, c.employees,
      c.stockReturn1Y, c.relativeReturn,
      ...c.monthlyRevenue,
      ...c.sameStoreSales,
      ...c.costRatio
    ]);
    const esc = v => {
      const s = v == null ? '' : String(v);
      return s.includes(',') || s.includes('"') || s.includes('\n') ? '"'+s.replace(/"/g,'""')+'"' : s;
    };
    const csv = BOM + headers.map(esc).join(',') + '\n' + rows.map(r => r.map(esc).join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '外食産業セクター分析_データ一覧_' + new Date().toISOString().slice(0,10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  };
})();

/* ── Event Delegation: data-action ── */
document.addEventListener('click', function(e) {
  var el = e.target.closest('[data-action]');
  if (!el) return;
  var fn = el.dataset.action;
  if (fn === 'downloadPDF') downloadPDF();
  else if (fn === 'downloadCSV') downloadCSV();
});

