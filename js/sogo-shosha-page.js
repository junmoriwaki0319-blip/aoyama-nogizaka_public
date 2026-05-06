/**
 * 総合商社セクター ダッシュボード
 * 7タブ: Executive Summary / 株主還元・資本政策 / 資本効率・バリュエーション /
 *        財務健全性 / アクティビスト・シグナル / 個別企業 / データソース
 */
(function(){
'use strict';

const TARGETS=[
  {ticker:'8058',name:'三菱商事',  tier:'big5'},
  {ticker:'8031',name:'三井物産',  tier:'big5'},
  {ticker:'8001',name:'伊藤忠商事',tier:'big5'},
  {ticker:'8053',name:'住友商事',  tier:'big5'},
  {ticker:'8002',name:'丸紅',      tier:'big5'},
  {ticker:'2768',name:'双日',      tier:'mid'},
  {ticker:'8015',name:'豊田通商',  tier:'mid'},
];
const TIERS={big5:'5大商社',mid:'準大手'};
const TC={big5:'#1a2d4f',mid:'#9b8b6e'};
const TB={big5:'badge-big5',mid:'badge-mid'};

let companies=[];
let buffett=null;
let strategic=null;
let narratives=null;
let tab='exec';
const CH={};

/* ── helpers ── */
function g(id){return document.getElementById(id);}
function shortName(n){return n.replace(/ホールディングス|HD|グループ|商事|物産/g,'').trim()||n;}
function fmtMcap(v){if(v==null)return'-';var oku=v/1e8;if(oku>=10000)return(oku/10000).toFixed(2)+'兆円';return oku.toFixed(0).toLocaleString()+'億円';}
function fmt(v,unit,d){if(v==null)return'<span class="na-cell">-</span>';if(typeof v!=='number')return v+(unit||'');return v.toFixed(d==null?2:d)+(unit||'');}
function nv(v,unit,fmtType){if(v==null)return'<span class="na-cell">-</span>';if(fmtType==='f1')return v.toFixed(1)+(unit||'');if(fmtType==='f2')return v.toFixed(2)+(unit||'');if(fmtType==='loc')return v.toLocaleString()+(unit||'');return v+(unit||'');}
function avg(arr,k){var v=arr.filter(function(c){return c.kpi(k)!=null;});return v.length?(v.reduce(function(s,c){return s+c.kpi(k);},0)/v.length):null;}
function topN(arr,k,n,desc){var c=arr.filter(function(x){return x.kpi(k)!=null;}).sort(function(x,y){return desc===false?x.kpi(k)-y.kpi(k):y.kpi(k)-x.kpi(k);});return c.slice(0,n);}
function byTier(t){return companies.filter(function(c){return c.tier===t;});}
function dc(ids){ids.forEach(function(id){if(CH[id]){CH[id].destroy();delete CH[id];}});}
function dm(t,s){var o=Object.assign({},t);Object.keys(s).forEach(function(k){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]);else o[k]=s[k];});return o;}
function mc(id,type,data,opts){var ctx=g(id);if(!ctx)return;var base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};if(['bar','line','scatter','bubble','radar'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};CH[id]=new Chart(ctx,{type:type,data:data,options:dm(base,opts||{})});}
function secH(n,t,d){return'<div class="sec-header"><div class="sec-num">SECTION '+n+'</div><div class="sec-title">'+t+'</div><div class="sec-desc">'+d+'</div></div>';}
function kpi(l,v,s,cls){return'<div class="kpi-card '+cls+'"><div class="kpi-label">'+l+'</div><div class="kpi-value">'+v+'</div>'+(s?'<div class="kpi-sub">'+s+'</div>':'')+'</div>';}
function rankCard(t,items,fn){return'<div class="ranking-card"><div class="ranking-title">'+t+'</div>'+items.map(function(c,i){return'<div class="ranking-row"><span><span class="ranking-num">'+(i+1)+'</span>'+shortName(c.name)+'</span><span style="font-weight:600">'+fn(c)+'</span></div>';}).join('')+'</div>';}

/* ── enable simple table sort ── */
function enableTableSort(tbl,data,cols,rowFn){
  var sortK=null,sortD='desc';
  var ths=tbl.querySelectorAll('thead th');
  ths.forEach(function(th,i){
    if(i>=cols.length||!cols[i])return;
    th.addEventListener('click',function(){
      var def=cols[i];
      if(sortK===def.key){sortD=sortD==='desc'?'asc':'desc';}else{sortK=def.key;sortD='desc';}
      ths.forEach(function(h){h.classList.remove('sort-asc','sort-desc');});
      th.classList.add(sortD==='asc'?'sort-asc':'sort-desc');
      var sorted=[].concat(data).sort(function(a,b){
        var va=def.kpi?a.kpi(def.key):a[def.key],vb=def.kpi?b.kpi(def.key):b[def.key];
        if(va==null)va=def.type==='num'?-Infinity:'';
        if(vb==null)vb=def.type==='num'?-Infinity:'';
        if(def.type==='num')return sortD==='desc'?vb-va:va-vb;
        return sortD==='desc'?String(vb).localeCompare(String(va)):String(va).localeCompare(String(vb));
      });
      tbl.querySelector('tbody').innerHTML=sorted.map(rowFn).join('');
    });
  });
}

/* ── rExec ── */
function rExec(){
  var el=g('sec-exec');
  var tm=companies.reduce(function(s,c){return s+(c.marketCap||0);},0);
  var aROE=avg(companies,'roe'),aPBR=avg(companies,'pbr'),aPER=avg(companies,'per'),aDOE=avg(companies,'doe');
  var aNDE=avg(companies,'net_debt_ebitda');
  var topMC=[].concat(companies).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);}).slice(0,7);
  var topROE=topN(companies,'roe',5);
  var topDOE=topN(companies,'doe',5);
  var b5=byTier('big5').length,m=byTier('mid').length;

  el.innerHTML=
    secH('01','Executive Summary','日本の総合商社7社（5大商社+双日・豊田通商）の株主還元・資本効率・財務健全性を19 KPIで横断分析')+
    '<div class="commentary">'+
      '<strong>セクター概況 (2026年5月時点 / FY24 通期実績ベース):</strong> 対象<strong>'+companies.length+'社</strong>の合計時価総額は<strong>'+fmtMcap(tm)+'</strong>。'+
      '5大商社（'+b5+'社）が時価総額の大半を占める一方、双日・豊田通商（'+m+'社）は事業ポートフォリオ構造が異なり「総合商社」というカテゴリ内でも比較軸の選択が重要。'+
      '平均ROE<strong>'+fmt(aROE,'%',1)+'</strong>、平均PBR<strong>'+fmt(aPBR,'倍',2)+'</strong>、平均DOE<strong>'+fmt(aDOE,'%',2)+'</strong>。'+
      'バークシャー・ハサウェイの5大商社保有（2020年以降の継続的買い増し）により、海外投資家からのバリュエーション再評価が進行中。'+
    '</div>'+
    '<div class="commentary navy">'+
      '<div class="commentary-title">フェーズ1 の論点</div>'+
      '<strong>(1) 株主還元政策の差別化軸:</strong> 三菱商事の DOE '+fmt(companies.find(function(c){return c.ticker==='8058';}).kpi('doe'),'%',2)+' / 配当性向 '+fmt(companies.find(function(c){return c.ticker==='8058';}).kpi('payout_ratio'),'%',1)+' は他社比突出して高く、株主還元のリーダー。<br>'+
      '<strong>(2) 資本効率の階層:</strong> 伊藤忠 (ROE '+fmt(companies.find(function(c){return c.ticker==='8001';}).kpi('roe'),'%',1)+') ・住友 ('+fmt(companies.find(function(c){return c.ticker==='8053';}).kpi('roe'),'%',1)+')・豊田通商 ('+fmt(companies.find(function(c){return c.ticker==='8015';}).kpi('roe'),'%',1)+') が ROE 13% 超のトップ層。三菱・三井は資源価格依存度が高く ROE 振れ幅が大きい。<br>'+
      '<strong>(3) ネット有利子負債/EBITDA レンジ:</strong> '+fmt(Math.min.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2)+' 〜 '+fmt(Math.max.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2)+' と社間で大きく分散。Yahoo 由来の EBITDA 定義は IFRS 開示の "Underlying Earnings" と乖離するため絶対水準ではなく**相対比較**で見る (詳細: データソースタブ)。<br>'+
      '<strong>(4) 開示構造の差:</strong> 「資源/非資源」分解は5大商社のみ有効。双日・豊田通商はそれぞれ独自セグメント区分のため、横断比較には別軸が必要。<br>'+
      '<strong>(5) 政策保有株式:</strong> 縮減進捗の比較は本フェーズでは未取得 (EDINET API キー取得後にフェーズ2 で補強予定)。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('対象企業数',companies.length+'社','5大商社:'+b5+' / 準大手:'+m,'c-navy')+
      kpi('時価総額合計',fmtMcap(tm),'','c-navy')+
      kpi('平均ROE',fmt(aROE,'%',1),'','c-green')+
      kpi('平均PBR',fmt(aPBR,'倍',2),'','c-gold')+
      kpi('平均PER',fmt(aPER,'倍',1),'','c-navy')+
      kpi('平均DOE',fmt(aDOE,'%',2),'株主資本配当率','c-gold')+
      kpi('平均ネット負債/EBITDA',fmt(aNDE,'倍',2),'','c-red')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">時価総額ランキング</div><div class="chart-panel-sub">単位: 億円</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PBR vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 赤線=PBR1.0</div><div class="chart-area tall"><canvas id="exPBR"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row tri">'+
      '<div class="chart-panel"><div class="chart-panel-title">ROE TOP 5</div><div class="chart-area"><canvas id="exROE"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">DOE TOP 5</div><div class="chart-area"><canvas id="exDOE"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">主要ランキング</div>'+
        '<div style="display:grid;gap:10px;padding-top:8px;">'+
          rankCard('時価総額',topMC.slice(0,5),function(c){return fmtMcap(c.marketCap);})+
        '</div></div>'+
    '</div>';

  dc(['exMC','exPBR','exROE','exDOE']);
  mc('exMC','bar',{labels:topMC.map(function(c){return shortName(c.name);}),datasets:[{data:topMC.map(function(c){return Math.round((c.marketCap||0)/1e8);}),backgroundColor:topMC.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toLocaleString()+'億';}}}});

  var mx=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));
  mc('exPBR','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.kpi('roe')!=null&&c.kpi('pbr')!=null;}).map(function(c){return{x:c.kpi('roe'),y:c.kpi('pbr'),r:Math.max(8,Math.sqrt((c.marketCap||1)/mx)*32),name:shortName(c.name)};}),backgroundColor:TC[tier]+'77',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'},min:0}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': ROE'+x.raw.x.toFixed(1)+'% / PBR'+x.raw.y.toFixed(2)+'倍';}}},annotation:{annotations:{pbr1:{type:'line',yMin:1.0,yMax:1.0,borderColor:'#b53a3a',borderWidth:1.5,borderDash:[4,4]}}},datalabels:{display:true,color:'#444',font:{size:10,weight:600},formatter:function(v,ctx){return v.name;},align:'top',offset:4}}});

  var roeTop=topN(companies,'roe',5);
  mc('exROE','bar',{labels:roeTop.map(function(c){return shortName(c.name);}),datasets:[{data:roeTop.map(function(c){return c.kpi('roe');}),backgroundColor:roeTop.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1)+'%';}}}});

  var doeTop=topN(companies,'doe',5);
  mc('exDOE','bar',{labels:doeTop.map(function(c){return shortName(c.name);}),datasets:[{data:doeTop.map(function(c){return c.kpi('doe');}),backgroundColor:doeTop.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(2)+'%';}}}});
}

/* ── rShareholder ── */
function rShareholder(){
  var el=g('sec-shareholder');
  var aDOE=avg(companies,'doe'),aPay=avg(companies,'payout_ratio');
  var sorted=[].concat(companies).sort(function(a,b){return(b.kpi('doe')||0)-(a.kpi('doe')||0);});

  el.innerHTML=
    secH('02','株主還元・資本政策','DOE・配当性向・自己株買い・政策保有株式の比較')+
    '<div class="commentary">'+
      '<strong>株主還元政策の俯瞰:</strong> 7社平均 DOE は<strong>'+fmt(aDOE,'%',2)+'</strong>、配当性向は<strong>'+fmt(aPay,'%',1)+'</strong>。'+
      'DOE が高い順に <strong>伊藤忠 → 丸紅 → 三菱 / 豊田通商 → 住友 → 三井 → 双日</strong>。'+
      '配当性向だけ見ると三菱が突出 (60%水準) だが、これは FY24 一過性減益によるもの。'+
      'DOE の方が本質的な「株主資本に対するキャッシュ・リターン」を示す。'+
    '</div>'+
    '<div class="commentary danger">'+
      '<strong>データ取得不能項目:</strong> 本タブで未取得 ― ① 総還元性向 (Yahoo Finance に自己株買い実施額がないため)、② 自己株買い実施額 (12M累計)、③ 政策保有 / 純資産比率 (有報 XBRL の追加パース必要)。'+
      '詳細は「データソース」タブを参照。フェーズ2 で EDINET API キー取得後に補強予定。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('平均DOE',fmt(aDOE,'%',2),'株主資本に対する配当','c-gold')+
      kpi('平均配当性向',fmt(aPay,'%',1),'純利益に対する配当','c-navy')+
      kpi('総還元性向','—','データ未取得','c-red')+
      kpi('自己株買い実施額','—','データ未取得','c-red')+
      kpi('政策保有/純資産','—','データ未取得','c-red')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">DOE 比較 (株主資本配当率)</div><div class="chart-panel-sub">単位: %</div><div class="chart-area tall"><canvas id="shDOE"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">配当性向 比較</div><div class="chart-panel-sub">単位: %</div><div class="chart-area tall"><canvas id="shPayout"></canvas></div></div>'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">株主還元 一覧 (DOE 降順)</div></div><div class="table-scroll"><table id="tblSh">'+
      '<thead><tr><th>コード</th><th>企業</th><th>区分</th><th>DOE (%)</th><th>配当性向 (%)</th><th>総還元性向</th><th>自己株買い</th><th>政策保有/純資産</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+TB[c.tier]+'">'+TIERS[c.tier]+'</span></td>'+
          '<td>'+nv(c.kpi('doe'),'%','f2')+'</td>'+
          '<td>'+nv(c.kpi('payout_ratio'),'%','f1')+'</td>'+
          '<td><span class="na-cell">未取得</span></td><td><span class="na-cell">未取得</span></td><td><span class="na-cell">未取得</span></td></tr>';
      }).join('')+'</tbody></table></div></div>';

  dc(['shDOE','shPayout']);
  var doeData=[].concat(companies).sort(function(a,b){return(b.kpi('doe')||0)-(a.kpi('doe')||0);});
  mc('shDOE','bar',{labels:doeData.map(function(c){return shortName(c.name);}),datasets:[{data:doeData.map(function(c){return c.kpi('doe');}),backgroundColor:doeData.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(2)+'%';}}}});
  var payData=[].concat(companies).sort(function(a,b){return(b.kpi('payout_ratio')||0)-(a.kpi('payout_ratio')||0);});
  mc('shPayout','bar',{labels:payData.map(function(c){return shortName(c.name);}),datasets:[{data:payData.map(function(c){return c.kpi('payout_ratio');}),backgroundColor:payData.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1)+'%';}}}});
}

/* ── rValuation ── */
function rValuation(){
  var el=g('sec-valuation');
  var aROE=avg(companies,'roe'),aROIC=avg(companies,'roic'),aPBR=avg(companies,'pbr'),aPER=avg(companies,'per'),aEV=avg(companies,'ev_ebitda');
  var sorted=[].concat(companies).sort(function(a,b){return(b.kpi('roe')||0)-(a.kpi('roe')||0);});

  el.innerHTML=
    secH('03','資本効率・バリュエーション','ROE / ROIC / PBR / PER / EV·EBITDA の横断比較')+
    '<div class="commentary">'+
      '<strong>資本効率の3階層:</strong> ROE 13% 超のトップ層 (伊藤忠・住友・豊田通商・丸紅)、10〜13% の中位 (三井・双日)、10% 以下のキャッチアップ層 (三菱)。'+
      'PBR は伊藤忠・豊田通商・丸紅が 2.2〜2.3 倍で最上位、双日 1.19 倍が最下位。<strong>PBR と ROE の相関は強く</strong>、市場は資本効率の高い商社にプレミアムを付与している。'+
    '</div>'+
    '<div class="commentary info">'+
      '<strong>数値解釈の注意 ― EV/EBITDA:</strong> Yahoo Finance ベースのため、商社が IR で開示する「Underlying Earnings」「Core Operating Cash Flow」ベースとは乖離する。絶対水準ではなく社間相対比較で参照のこと。'+
      '<strong>ROIC:</strong> 「簡易 ROIC」(NI / (Equity + Debt)) を採用。標準的な NOPAT/IC とは定義が異なる点に留意。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('平均ROE',fmt(aROE,'%',1),'7社平均','c-green')+
      kpi('平均ROIC',fmt(aROIC,'%',1),'簡易 (NI/IC)','c-gold')+
      kpi('平均PBR',fmt(aPBR,'倍',2),'7社平均','c-navy')+
      kpi('平均PER',fmt(aPER,'倍',1),'7社平均','c-navy')+
      kpi('平均EV/EBITDA',fmt(aEV,'倍',1),'Yahoo 基準','c-gold')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">PBR vs ROE</div><div class="chart-panel-sub">右上ほど高評価。バブル=時価総額</div><div class="chart-area tall"><canvas id="vlPBR"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PER vs EV/EBITDA</div><div class="chart-panel-sub">バリュエーション・コンプ比較</div><div class="chart-area tall"><canvas id="vlPER"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">ROE / ROIC 比較</div><div class="chart-area tall"><canvas id="vlROE"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PBR / PER ランキング</div><div class="chart-area tall"><canvas id="vlPBRBar"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row single">'+
      '<div class="chart-panel">'+
        '<div class="chart-panel-title">7 社総合スコア レーダー</div>'+
        '<div class="chart-panel-sub">5 軸 (ROE / ROIC / PBR / PER 反転 / EV·EBITDA 反転) を 0–100 に正規化。外側ほど高評価。同一軸内では最良値=100、最悪値=20 として線形配点</div>'+
        '<div class="chart-area" style="height:520px;"><canvas id="vlRadar"></canvas></div>'+
      '</div>'+
    '</div>'+
    '<div class="chart-panel" style="margin-bottom:32px;">'+
      '<div class="chart-panel-title">19 KPI × 7 社 ヒートマップ</div>'+
      '<div class="chart-panel-sub">取得済 KPI のみ色付け。各行 (KPI 軸) で最良 / 良い / 中位 / 悪い / 最悪の 5 段階。反転指標 (PER ・ EV/EBITDA ・ ネット負債/EBITDA) は低い順に色付け</div>'+
      '<div class="table-scroll" style="margin-top:16px;"><table class="heatmap-table" id="tblHeat"><thead></thead><tbody></tbody></table></div>'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">資本効率・バリュエーション 一覧 (ROE 降順)</div></div><div class="table-scroll"><table id="tblVal">'+
      '<thead><tr><th>コード</th><th>企業</th><th>区分</th><th>ROE (%)</th><th>ROIC (%)</th><th>PBR (倍)</th><th>PER (倍)</th><th>EV/EBITDA (倍)</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+TB[c.tier]+'">'+TIERS[c.tier]+'</span></td>'+
          '<td>'+nv(c.kpi('roe'),'%','f1')+'</td><td>'+nv(c.kpi('roic'),'%','f2')+'</td>'+
          '<td>'+nv(c.kpi('pbr'),'倍','f2')+'</td><td>'+nv(c.kpi('per'),'倍','f1')+'</td>'+
          '<td>'+nv(c.kpi('ev_ebitda'),'倍','f1')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';

  // ヒートマップ描画 (19 KPI × 7 社)
  buildHeatmap('tblHeat');

  dc(['vlPBR','vlPER','vlROE','vlPBRBar','vlRadar']);
  var mx=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));
  mc('vlPBR','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.kpi('roe')!=null&&c.kpi('pbr')!=null;}).map(function(c){return{x:c.kpi('roe'),y:c.kpi('pbr'),r:Math.max(8,Math.sqrt((c.marketCap||1)/mx)*32),name:shortName(c.name)};}),backgroundColor:TC[tier]+'77',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'},min:0}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': ROE'+x.raw.x.toFixed(1)+'% / PBR'+x.raw.y.toFixed(2)+'倍';}}},datalabels:{display:true,color:'#444',font:{size:10,weight:600},formatter:function(v){return v.name;},align:'top',offset:4}}});
  mc('vlPER','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.kpi('per')!=null&&c.kpi('ev_ebitda')!=null;}).map(function(c){return{x:c.kpi('per'),y:c.kpi('ev_ebitda'),r:Math.max(8,Math.sqrt((c.marketCap||1)/mx)*32),name:shortName(c.name)};}),backgroundColor:TC[tier]+'77',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'PER (倍)'}},y:{title:{display:true,text:'EV/EBITDA (倍)'}}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': PER'+x.raw.x.toFixed(1)+' / EV·EBITDA'+x.raw.y.toFixed(1);}}},datalabels:{display:true,color:'#444',font:{size:10,weight:600},formatter:function(v){return v.name;},align:'top',offset:4}}});

  var roeOrd=[].concat(companies).sort(function(a,b){return(b.kpi('roe')||0)-(a.kpi('roe')||0);});
  mc('vlROE','bar',{labels:roeOrd.map(function(c){return shortName(c.name);}),datasets:[{label:'ROE',data:roeOrd.map(function(c){return c.kpi('roe');}),backgroundColor:'rgba(45,122,79,0.7)',borderWidth:0},{label:'ROIC (簡易)',data:roeOrd.map(function(c){return c.kpi('roic');}),backgroundColor:'rgba(155,139,110,0.7)',borderWidth:0}]},{plugins:{legend:{display:true,position:'top'},datalabels:{display:false}},scales:{y:{title:{display:true,text:'%'}}}});

  var pbrOrd=[].concat(companies).sort(function(a,b){return(b.kpi('pbr')||0)-(a.kpi('pbr')||0);});
  mc('vlPBRBar','bar',{labels:pbrOrd.map(function(c){return shortName(c.name);}),datasets:[{label:'PBR',data:pbrOrd.map(function(c){return c.kpi('pbr');}),backgroundColor:'rgba(26,45,79,0.7)',borderWidth:0,yAxisID:'y'},{label:'PER',data:pbrOrd.map(function(c){return c.kpi('per');}),backgroundColor:'rgba(155,139,110,0.7)',borderWidth:0,yAxisID:'y2'}]},{plugins:{legend:{display:true,position:'top'},datalabels:{display:false}},scales:{y:{title:{display:true,text:'PBR (倍)'},position:'left'},y2:{title:{display:true,text:'PER (倍)'},position:'right',grid:{display:false}}}});

  // レーダーチャート: 5 軸を 0-100 に線形正規化。反転指標は反転後にスコア化
  var radarKeys=[
    {key:'roe',     label:'ROE',         invert:false},
    {key:'roic',    label:'ROIC',        invert:false},
    {key:'pbr',     label:'PBR',         invert:false},
    {key:'per',     label:'PER (反転)',  invert:true},
    {key:'ev_ebitda',label:'EV/EBITDA (反転)', invert:true},
  ];
  var ranges={};
  radarKeys.forEach(function(rk){
    var vals=companies.map(function(c){return c.kpi(rk.key);}).filter(function(v){return v!=null;});
    ranges[rk.key]={min:Math.min.apply(null,vals),max:Math.max.apply(null,vals)};
  });
  function score(v,rk){
    if(v==null)return 0;
    var r=ranges[rk.key];
    if(r.max===r.min)return 60;
    var raw=(v-r.min)/(r.max-r.min); // 0..1
    if(rk.invert)raw=1-raw;
    return 20+raw*80; // 20..100
  }
  var radarColors=['#1a2d4f','#2a4470','#9b8b6e','#7a6d55','#2d7a4f','#b53a3a','#5555aa'];
  var radarDatasets=companies.map(function(c,i){
    return{
      label:shortName(c.name),
      data:radarKeys.map(function(rk){return score(c.kpi(rk.key),rk);}),
      backgroundColor:radarColors[i]+'18',
      borderColor:radarColors[i],
      borderWidth:1.5,
      pointBackgroundColor:radarColors[i],
      pointRadius:3,
    };
  });
  mc('vlRadar','radar',{labels:radarKeys.map(function(rk){return rk.label;}),datasets:radarDatasets},{plugins:{legend:{display:true,position:'bottom',labels:{font:{size:11},padding:14,boxWidth:14}},datalabels:{display:false},tooltip:{callbacks:{label:function(x){var c=companies[x.datasetIndex];var rk=radarKeys[x.dataIndex];var raw=c.kpi(rk.key);return c.name+' / '+rk.label+': '+(raw==null?'未取得':raw.toFixed(2))+(rk.key==='roe'||rk.key==='roic'?'%':'倍')+' (score '+x.parsed.r.toFixed(0)+')';}}}},scales:{r:{min:0,max:100,ticks:{stepSize:20,color:'#999',font:{size:9},backdropColor:'transparent'},pointLabels:{color:'#444',font:{size:11,weight:'600'}},grid:{color:'#eae7e1'},angleLines:{color:'#eae7e1'}}}});
}

/* ── ヒートマップ ヘルパー ── */
function buildHeatmap(tblId){
  var tbl=g(tblId);if(!tbl)return;
  var inverted={per:true,ev_ebitda:true,net_debt_ebitda:true};
  var allKeys=Object.keys(companies[0].summary.kpis);
  var thead=tbl.querySelector('thead'),tbody=tbl.querySelector('tbody');
  thead.innerHTML='<tr><th style="text-align:left;min-width:160px;">KPI</th>'+companies.map(function(c){return'<th style="text-align:center;min-width:80px;">'+c.ticker+'<br><span style="font-size:0.65rem;color:var(--text-light);font-weight:400;">'+shortName(c.name)+'</span></th>';}).join('')+'</tr>';
  tbody.innerHTML=allKeys.map(function(k){
    var meta=companies[0].summary.kpis[k];
    var vals=companies.map(function(c){return c.kpi(k);});
    var present=vals.filter(function(v){return v!=null;});
    if(!present.length){
      return'<tr><td>'+meta.label+'</td>'+vals.map(function(){return'<td><span class="heat-cell heat-na">—</span></td>';}).join('')+'</tr>';
    }
    // ランク付け (反転指標は昇順、それ以外は降順)
    var sorted=present.slice().sort(function(a,b){return inverted[k]?a-b:b-a;});
    function rank(v){if(v==null)return -1;return sorted.indexOf(v);}
    function tierClass(v){
      if(v==null)return'heat-na';
      var r=rank(v),n=sorted.length;
      if(n<=1)return'heat-mid';
      var pct=r/(n-1); // 0=best, 1=worst
      if(pct<=0.15)return'heat-best';
      if(pct<=0.4)return'heat-good';
      if(pct<=0.6)return'heat-mid';
      if(pct<=0.85)return'heat-weak';
      return'heat-worst';
    }
    function fmtVal(v){if(v==null)return'—';if(meta.unit==='%')return v.toFixed(1)+'%';if(meta.unit==='倍')return v.toFixed(2);return v.toString();}
    return'<tr><td style="text-align:left;font-size:0.72rem;color:var(--text-mid);">'+meta.label+(inverted[k]?' <span style="color:var(--text-light);font-size:0.62rem;">↓良</span>':'')+'</td>'+
      vals.map(function(v){return'<td><span class="heat-cell '+tierClass(v)+'">'+fmtVal(v)+'</span></td>';}).join('')+'</tr>';
  }).join('');
}

/* ── rFinancial ── */
function rFinancial(){
  var el=g('sec-financial');
  var aNDE=avg(companies,'net_debt_ebitda');
  var sorted=[].concat(companies).sort(function(a,b){return(a.kpi('net_debt_ebitda')||999)-(b.kpi('net_debt_ebitda')||999);});

  el.innerHTML=
    secH('04','財務健全性','ネット有利子負債/EBITDA・自己資本比率・インタレスト・カバレッジ')+
    '<div class="commentary">'+
      '<strong>レバレッジの分散:</strong> 7社のネット有利子負債/EBITDA は'+fmt(Math.min.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2)+' (豊田通商) から '+fmt(Math.max.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2)+' (双日) まで広く分散。'+
      '伊藤忠 (3.41) と豊田通商 (0.94) は財務余力が大きい。双日 (10.38) は 2025/3 期 M&A の影響と見られるため、決算説明会資料での確認が必要。'+
    '</div>'+
    '<div class="commentary danger">'+
      '<strong>データ取得不能項目:</strong> 本タブで未取得 ― ① 自己資本比率 (Yahoo に総資産フィールドなし)、② インタレスト・カバレッジ (Yahoo に支払利息なし)。EDINET XBRL の追加パースで取得可能 (フェーズ2 予定)。'+
      '<br><br><strong>解釈の注意:</strong> Yahoo の totalDebt 定義に IFRS 16 リース負債が含まれ得る。商社 IR 開示の「ネット有利子負債」とは差が出る可能性あり。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('平均ネット負債/EBITDA',fmt(aNDE,'倍',2),'7社平均','c-red')+
      kpi('最低 (財務余力大)',fmt(Math.min.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2),'豊田通商 (8015)','c-green')+
      kpi('最高 (要注意)',fmt(Math.max.apply(null,companies.filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).map(function(c){return c.kpi('net_debt_ebitda');})),'倍',2),'双日 (2768)','c-red')+
      kpi('自己資本比率','—','データ未取得','c-red')+
      kpi('インタレスト・カバレッジ','—','データ未取得','c-red')+
    '</div>'+
    '<div class="chart-row single">'+
      '<div class="chart-panel"><div class="chart-panel-title">ネット有利子負債 / EBITDA 倍率</div><div class="chart-panel-sub">低いほど財務余力大 / 単位: 倍</div><div class="chart-area tall"><canvas id="fnNDE"></canvas></div></div>'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">財務健全性 一覧 (ネット負債/EBITDA 昇順)</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th>コード</th><th>企業</th><th>区分</th><th>ネット負債/EBITDA (倍)</th><th>自己資本比率 (%)</th><th>インタレスト・カバレッジ (倍)</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+TB[c.tier]+'">'+TIERS[c.tier]+'</span></td>'+
          '<td>'+nv(c.kpi('net_debt_ebitda'),'倍','f2')+'</td>'+
          '<td><span class="na-cell">未取得</span></td><td><span class="na-cell">未取得</span></td></tr>';
      }).join('')+'</tbody></table></div></div>';

  dc(['fnNDE']);
  var nde=[].concat(companies).filter(function(c){return c.kpi('net_debt_ebitda')!=null;}).sort(function(a,b){return a.kpi('net_debt_ebitda')-b.kpi('net_debt_ebitda');});
  mc('fnNDE','bar',{labels:nde.map(function(c){return shortName(c.name);}),datasets:[{data:nde.map(function(c){return c.kpi('net_debt_ebitda');}),backgroundColor:nde.map(function(c){var v=c.kpi('net_debt_ebitda');return v<3?'rgba(45,122,79,0.7)':v<6?'rgba(26,45,79,0.7)':'rgba(181,58,58,0.7)';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(2)+'倍';}}}});
}

/* ── rActivist ── */
function rActivist(){
  var el=g('sec-activist');
  var aCM=avg(companies,'cash_to_mcap');
  var sorted=[].concat(companies).sort(function(a,b){return(b.kpi('cash_to_mcap')||0)-(a.kpi('cash_to_mcap')||0);});

  el.innerHTML=
    secH('05','アクティビスト・シグナル','現預金/時価総額比・大量保有報告件数・政策保有縮減率')+
    '<div class="commentary">'+
      '<strong>「現預金/時価総額」が高い社ほど、アクティビストにとっての"資本配当の効率化余地"が大きい</strong>のがセオリー。'+
      '7社平均は<strong>'+fmt(aCM,'%',1)+'%</strong>で、双日 (19.79%) と豊田通商 (20.49%) が最も高い。一方、5大商社のうち丸紅 (5.49%) と伊藤忠 (5.77%) は現預金水準が比較的タイト。'+
      '5大商社のバークシャー保有 (Berkshire Hathaway) は2020年来の大型エントリーで、長期投資家としての安定保有が継続中 (双日・豊田通商は対象外)。'+
    '</div>'+
    '<div class="commentary danger">'+
      '<strong>データ取得不能項目:</strong> 本タブで未取得 ― ① 大量保有報告件数 (12M) は EDINET API キー未取得のため。'+
      '② 政策保有縮減率 (過去5期) は 5 期分の有報を遡る + タクソノミー差吸収が必要 (中〜高難度)。<br>'+
      '<strong>バフェット保有推移 (5大商社のみ):</strong> 13F filings から SEC EDGAR 経由で取得可能 (フェーズ2 予定)。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('平均 現預金/時価総額',fmt(aCM,'%',1),'7社平均','c-gold')+
      kpi('最高',fmt(Math.max.apply(null,companies.filter(function(c){return c.kpi('cash_to_mcap')!=null;}).map(function(c){return c.kpi('cash_to_mcap');})),'%',1),'豊田通商 (8015)','c-navy')+
      kpi('最低',fmt(Math.min.apply(null,companies.filter(function(c){return c.kpi('cash_to_mcap')!=null;}).map(function(c){return c.kpi('cash_to_mcap');})),'%',1),'丸紅 (8002)','c-navy')+
      kpi('大量保有報告件数','—','EDINET キー必要','c-red')+
      kpi('政策保有縮減率','—','データ未取得','c-red')+
    '</div>'+
    '<div class="chart-row single">'+
      '<div class="chart-panel"><div class="chart-panel-title">現預金 / 時価総額 比率</div><div class="chart-panel-sub">単位: % / 高いほどアクティビストにとっての資本配当余地が大</div><div class="chart-area tall"><canvas id="acCM"></canvas></div></div>'+
    '</div>'+
    '<div class="commentary">'+
      '<strong>関連分析:</strong> バフェット (Berkshire Hathaway) の 5 大商社保有推移、各社の戦略的提携・DX 動向は <strong>「06 戦略・提携動向」タブ</strong> に集約しています。'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">アクティビスト・シグナル 一覧 (現預金/時価総額 降順)</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th>コード</th><th>企業</th><th>区分</th><th>現預金/時価総額 (%)</th><th>大量保有報告件数 (12M)</th><th>政策保有縮減率 (5期)</th><th>バフェット保有</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+TB[c.tier]+'">'+TIERS[c.tier]+'</span></td>'+
          '<td>'+nv(c.kpi('cash_to_mcap'),'%','f1')+'</td>'+
          '<td><span class="na-cell">未取得</span></td><td><span class="na-cell">未取得</span></td>'+
          '<td>'+(c.tier==='big5'?'<span class="badge badge-big5">対象</span>':'<span class="na-cell">対象外</span>')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';

  dc(['acCM']);
  mc('acCM','bar',{labels:sorted.map(function(c){return shortName(c.name);}),datasets:[{data:sorted.map(function(c){return c.kpi('cash_to_mcap');}),backgroundColor:sorted.map(function(c){return TC[c.tier];}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1)+'%';}}}});
}

/* ── rPartnership ── */
function rPartnership(){
  var el=g('sec-partnership');
  if(!buffett||!strategic){
    el.innerHTML=secH('06','戦略・提携動向','バフェット保有推移と各社の DX / 提携施策')+'<div class="commentary danger"><strong>データ未読込:</strong> refs/buffett-holdings.json または refs/strategic-initiatives.json の取得に失敗しました。</div>';
    return;
  }

  var big5=['8058','8031','8001','8053','8002'];
  var cols=buffett._columns;
  var datasets=big5.map(function(t,i){
    var trend=buffett.holdings_pct[t];
    return{label:trend.name,data:trend.trend,borderColor:['#1a2d4f','#2d7a4f','#9b8b6e','#7a6d55','#b53a3a'][i],backgroundColor:['#1a2d4f','#2d7a4f','#9b8b6e','#7a6d55','#b53a3a'][i]+'15',fill:false,tension:0.25,pointRadius:4,pointHoverRadius:6,borderWidth:2};
  });

  var milestoneHtml=buffett.milestones.map(function(m){
    var idx=m.indexOf(':');
    var d=idx>0?m.slice(0,idx):'';
    var t=idx>0?m.slice(idx+1).trim():m;
    return'<div style="display:flex;gap:14px;padding:8px 0;border-bottom:1px dashed var(--border-light);font-size:0.78rem;line-height:1.7;"><span style="font-family:Cormorant Garamond,Georgia,serif;color:var(--gold);font-weight:700;letter-spacing:1px;min-width:90px;">'+d+'</span><span style="color:var(--text-mid);">'+t+'</span></div>';
  }).join('');

  var implHtml=buffett.narrative.implications.map(function(s,i){
    return'<div style="display:flex;gap:10px;padding:6px 0;font-size:0.8rem;line-height:1.8;"><span style="font-family:Cormorant Garamond,Georgia,serif;color:var(--gold);font-weight:700;min-width:18px;">'+(i+1)+'.</span><span style="color:var(--text-mid);">'+s+'</span></div>';
  }).join('');

  var partnerCardsHtml=Object.keys(strategic.companies).map(function(t){
    var c=strategic.companies[t];
    var tierMeta=companies.find(function(x){return x.ticker===t;});
    var tierLabel=tierMeta?TIERS[tierMeta.tier]:'';
    var tierBadgeCls=tierMeta?TB[tierMeta.tier]:'badge-mid';
    var items=c.items.map(function(item){
      return'<div class="partner-item">'+
        '<div class="partner-item-title"><span class="badge-cat badge-cat-'+item.category+'">'+item.category+'</span>'+item.title+'</div>'+
        '<div class="partner-item-summary">'+item.summary+'</div>'+
      '</div>';
    }).join('');
    return'<div class="partner-card">'+
      '<div class="partner-card-header">'+
        '<div><span class="partner-card-ticker">'+t+'</span><span class="partner-card-name">'+c.name+'</span></div>'+
        '<span class="badge '+tierBadgeCls+'">'+tierLabel+'</span>'+
      '</div>'+
      '<div class="partner-card-headline">'+c.headline+'</div>'+
      items+
    '</div>';
  }).join('');

  el.innerHTML=
    secH('06','戦略・提携動向','バフェット (Berkshire Hathaway) 保有推移と各社の DX / 戦略的提携を集約')+
    '<div class="commentary">'+
      '<strong>'+buffett.narrative.thesis+'</strong>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">バフェット (Berkshire Hathaway) 5大商社保有比率推移</div><div class="chart-panel-sub">単位: % / 双日・豊田通商は対象外 / 出典: Berkshire Annual Letter + 各社大量保有報告</div><div class="chart-area tall"><canvas id="ptBuffett"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">バフェット投資のマイルストーン</div><div class="chart-panel-sub">2020 年初開示 〜 直近の動き</div><div style="margin-top:8px;">'+milestoneHtml+'</div></div>'+
    '</div>'+
    '<div class="chart-row single">'+
      '<div class="chart-panel"><div class="chart-panel-title">投資テーゼの含意 (3 観点)</div><div class="chart-panel-sub">5 大商社の長期保有が日本商社セクター全体に与える構造的影響</div><div style="margin-top:12px;">'+implHtml+'</div></div>'+
    '</div>'+
    '<div class="sec-header" style="margin-top:24px;"><div class="sec-num">SECTION 06 — 戦略動向</div><div class="sec-title">各社の DX・戦略的提携</div><div class="sec-desc">公開情報ベースで主要施策を抽出。網羅性は意図せず、業界比較に資する象徴的な施策のみを掲載。</div></div>'+
    '<div class="commentary navy">'+
      '<strong>カテゴリ凡例:</strong>'+
      ' <span class="badge-cat badge-cat-DX">DX</span>デジタル・データ・AI'+
      ' <span class="badge-cat badge-cat-提携">提携</span>戦略的パートナーシップ・合弁'+
      ' <span class="badge-cat badge-cat-海外">海外</span>地域・新興国戦略'+
      ' <span class="badge-cat badge-cat-脱炭素">脱炭素</span>再エネ・水素・CCUS'+
      ' <span class="badge-cat badge-cat-コンスーマー">コンスーマー</span>リテール・B2C・消費者接点'+
    '</div>'+
    '<div class="chart-row tri" style="grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:20px;">'+partnerCardsHtml+'</div>';

  dc(['ptBuffett']);
  mc('ptBuffett','line',{labels:cols,datasets:datasets},{plugins:{legend:{display:true,position:'bottom',labels:{font:{size:11},padding:14}},datalabels:{display:false},tooltip:{callbacks:{label:function(x){return x.dataset.label+': '+x.parsed.y.toFixed(1)+'%';}}}},scales:{y:{title:{display:true,text:'保有比率 (%)'},min:4,max:11},x:{title:{display:true,text:''}}}});
}

/* ── rDetail ── */
function rDetail(){
  var el=g('sec-detail');

  // 5大商社平均 (deviation bar 用) を事前計算
  var big5=byTier('big5');
  var devKeys=['doe','payout_ratio','roe','roic','pbr','per','ev_ebitda','net_debt_ebitda','cash_to_mcap'];
  var avgs={};
  devKeys.forEach(function(k){var vals=big5.map(function(c){return c.kpi(k);}).filter(function(v){return v!=null;});avgs[k]=vals.length?vals.reduce(function(s,v){return s+v;},0)/vals.length:null;});

  // KPI ハイライトの選定 (各社で目立つ数値 4 つ)
  function pickHighlights(c){
    return[
      {label:'時価総額',value:fmtMcap(c.marketCap),sub:'FY24'},
      {label:'ROE',value:fmt(c.kpi('roe'),'%',1),sub:'5大商社平均 '+fmt(avgs.roe,'%',1)},
      {label:'PBR',value:fmt(c.kpi('pbr'),'倍',2),sub:'5大商社平均 '+fmt(avgs.pbr,'倍',2)},
      {label:c.tier==='big5'?'バフェット保有':'バフェット対象',value:c.tier==='big5'?(buffett?(buffett.holdings_pct[c.ticker].trend.slice(-1)[0].toFixed(1)+'%'):'—'):'対象外',sub:c.tier==='big5'?'2025/3 直近':'5大商社のみ'},
    ];
  }

  function bulletList(arr){return arr.map(function(s){return'<div class="narrative-bullet">'+s+'</div>';}).join('');}

  function deviationBars(c){
    var labels={doe:'DOE',payout_ratio:'配当性向',roe:'ROE',roic:'ROIC (簡易)',pbr:'PBR',per:'PER',ev_ebitda:'EV/EBITDA',net_debt_ebitda:'ネット負債/EBITDA',cash_to_mcap:'現預金/時価総額'};
    // ネット負債/EBITDA は低いほど良い指標 → 反転表示
    var inverted={net_debt_ebitda:true,per:true,ev_ebitda:true};
    var rows=devKeys.map(function(k){
      var v=c.kpi(k),avg=avgs[k];
      if(v==null||avg==null||avg===0)return'<div class="deviation-row"><span class="deviation-label">'+labels[k]+'</span><div class="deviation-track"></div><span class="deviation-value" style="color:var(--text-light);font-weight:400;">—</span></div>';
      var rawDev=(v-avg)/Math.abs(avg)*100;
      var dispDev=inverted[k]?-rawDev:rawDev;  // 反転指標: 低い方が "+"
      var clamped=Math.max(-99,Math.min(99,dispDev));
      var widthPct=Math.min(50,Math.abs(clamped)/2); // ±100% で trackの半分埋める
      var sign=dispDev>=0?'pos':'neg';
      var fillStyle=dispDev>=0?'left:50%;width:'+widthPct+'%;':'right:50%;width:'+widthPct+'%;';
      var label=(dispDev>=0?'+':'')+dispDev.toFixed(0)+'%';
      return'<div class="deviation-row"><span class="deviation-label">'+labels[k]+'</span><div class="deviation-track"><div class="deviation-fill '+sign+'" style="'+fillStyle+'"></div></div><span class="deviation-value '+sign+'">'+label+'</span></div>';
    }).join('');
    return'<div class="deviation-section-title">5 大商社平均からの乖離 (反転指標: 低いほど評価良い指標は + 表示)</div>'+rows;
  }

  function narrativeCard(c){
    var n=narratives&&narratives.companies?narratives.companies[c.ticker]:null;
    var hl=pickHighlights(c);
    var hlHtml=hl.map(function(h){return'<span style="margin-right:18px;font-size:0.78rem;color:var(--text-muted);"><strong style="color:var(--navy);font-size:0.95rem;font-weight:700;">'+h.value+'</strong> '+h.label+'<span style="color:var(--text-light);margin-left:6px;font-size:0.7rem;">('+h.sub+')</span></span>';}).join('');
    var body=n?(
      '<div class="narrative-headline">'+n.headline+'</div>'+
      '<div class="narrative-thesis"><span class="narrative-thesis-label">INVESTMENT THESIS</span>'+n.thesis+'</div>'+
      '<div class="narrative-grid">'+
        '<div><div class="narrative-block-title is-strength">強み</div>'+bulletList(n.strengths)+'</div>'+
        '<div><div class="narrative-block-title is-challenge">課題・リスク</div>'+bulletList(n.challenges)+'</div>'+
        '<div><div class="narrative-block-title is-recent">直近トピック</div>'+bulletList(n.recent)+'</div>'+
      '</div>'
    ):'<div class="commentary danger" style="margin-bottom:14px;"><strong>ナラティブ未取得:</strong> data/sogo-shosha/refs/company-narratives.json を確認</div>';

    return'<div class="narrative-card">'+
      '<div class="narrative-card-header">'+
        '<div><span class="narrative-card-id">'+c.ticker+'</span><span class="narrative-card-name">'+c.name+'</span> <span class="badge '+TB[c.tier]+'" style="margin-left:8px;">'+TIERS[c.tier]+'</span></div>'+
        '<div class="narrative-card-meta">'+hlHtml+'</div>'+
      '</div>'+
      body+
      deviationBars(c)+
    '</div>';
  }

  el.innerHTML=
    secH('07','個別企業分析','各社のヘッドライン・投資テーゼ・5 大商社平均からの乖離を 1 枚で')+
    '<div class="commentary">'+
      '<strong>読み方:</strong> 各社のカードは「ヘッドライン → 投資テーゼ → 強み・課題・直近トピック → 5 大商社平均からの乖離バー」で構成。'+
      '乖離バーは <span style="color:var(--green);font-weight:700;">緑 (+)</span> が「他 5 社平均より良い」、<span style="color:var(--red);font-weight:700;">赤 (−)</span> が「悪い」を示す。'+
      'PER ・ EV/EBITDA ・ ネット負債/EBITDA は数値が低いほど評価が高い反転指標のため、表示上も反転して + / − を割当てている。'+
    '</div>'+
    companies.map(narrativeCard).join('');
}

/* ── rSource ── */
function rSource(){
  var el=g('sec-source');
  el.innerHTML=
    secH('A','データソース','取得状況・既知の制約・更新タイミング')+
    '<div class="commentary">'+
      '<strong>本ダッシュボードのデータ:</strong> Yahoo Finance (TTM スナップショット, lastFiscalYearEnd=2026-03-31) を一次ソースとして 9/19 KPI を自動取得。'+
      'EDINET / J-Quants API キー未取得のため、政策保有株式・自己株買い・大量保有報告件数等は未取得。フェーズ2 で補強予定。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('取得日','2026-05-07','','c-navy')+
      kpi('基準期','FY24 通期実績','TTM (2026/3 期末)','c-navy')+
      kpi('対象企業','7社','5大商社+双日+豊田通商','c-navy')+
      kpi('KPI 充足率','9/19','47.4% (KPI のみ)','c-gold')+
      kpi('補足レイヤー','2/3','Buffett + 戦略動向 取込済','c-green')+
      kpi('未取得 KPI','10項目','known-issues.md §2','c-red')+
    '</div>'+
    '<div class="chart-row single">'+
      '<div class="chart-panel">'+
        '<div class="chart-panel-title">データソース別 取得 KPI</div>'+
        '<div class="chart-panel-sub">取得済 / 未取得の内訳 (未取得分はフェーズ2 で補強)</div>'+
        '<div style="margin-top:18px;">'+
          '<div style="margin-bottom:18px;"><div style="font-weight:700;color:var(--navy);font-size:0.85rem;margin-bottom:8px;">Yahoo Finance (yahoo-finance2 npm) <span style="color:var(--green);font-weight:600;font-size:0.7rem;margin-left:6px;">取得済 9 KPI</span></div>'+
            '<div style="font-size:0.78rem;color:var(--text-mid);line-height:1.9;">'+
              'DOE / 配当性向 / ROE / ROIC (簡易) / PBR / PER / EV·EBITDA / ネット有利子負債/EBITDA / 現預金 比率'+
            '</div></div>'+
          '<div style="margin-bottom:18px;"><div style="font-weight:700;color:var(--navy);font-size:0.85rem;margin-bottom:8px;">EDINET API <span style="color:var(--red);font-weight:400;font-size:0.7rem;margin-left:6px;">未取得 4 KPI ― キー未発行</span></div>'+
            '<div style="font-size:0.78rem;color:var(--text-mid);line-height:1.9;">'+
              '政策保有 / 純資産比率 / 大量保有報告件数 (12M) / 自己株買い実施額 / 総還元性向'+
            '</div></div>'+
          '<div style="margin-bottom:18px;"><div style="font-weight:700;color:var(--navy);font-size:0.85rem;margin-bottom:8px;">各社 IR 説明会資料 / 統合報告書 <span style="color:var(--red);font-weight:400;font-size:0.7rem;margin-left:6px;">未取得 6 KPI ― 手動収集</span></div>'+
            '<div style="font-size:0.78rem;color:var(--text-mid);line-height:1.9;">'+
              '資源 営業利益比率 / 海外売上高比率 / セグメント別 ROIC 開示 / 自己資本比率 / インタレスト・カバレッジ / 政策保有縮減率'+
            '</div></div>'+
        '</div>'+
      '</div>'+
    '</div>'+
    '<div class="commentary navy">'+
      '<strong>開示構造の差 ― 5大商社 vs 双日・豊田通商:</strong>'+
      '<br><br><strong>「資源/非資源」分解:</strong> 5大商社は IR で公式分解あり。双日は 5 セグメント分解 (金属・資源 / 化学 / 自動車 / インフラ・ヘルスケア / リテール) で「資源/非資源」直接分解はなし。豊田通商は 7 セグメント (金属 / 部品&ロジ / 自動車 / 機械・エネ / 化学品 / 食料 / アフリカ) で<strong>自動車関連が約 50%</strong>を占め、「資源/非資源」カテゴリ自体が IR で使われない。<br>'+
      '<strong>セグメント別 ROIC:</strong> 伊藤忠・三菱は IR で開示 (詳細パターン)。三井・住友・丸紅は限定的。双日・豊田通商は非開示。<br>'+
      '<strong>バフェット保有:</strong> 5大商社のみ対象。双日・豊田通商は対象外。'+
    '</div>'+
    '<div class="commentary info">'+
      '<strong>更新タイミング ― 株価系 (PER/PBR/時価総額):</strong> 日次でドリフト。デプロイ前に再取得推奨。<br>'+
      '<strong>通期決算系 (ROE/ROIC/DOE/EV·EBITDA):</strong> 2026/3 期決算発表完了後の Yahoo データ更新待ち (タイムラグ 数日〜2週間)。<br>'+
      '<strong>政策保有株式の縮減率:</strong> 統合報告書発行後 (7〜8月) に年次更新。<br>'+
      '<strong>バフェット保有:</strong> 四半期ごと (13F 提出時 / 次回 2026年5月15日)。'+
    '</div>'+
    '<div class="commentary">'+
      '<strong>処理済データ:</strong> processed/comparison-matrix.csv (7 社 × 19 KPI / BOM 付 UTF-8)、'+
      'processed/&lt;ticker&gt;-summary.json (各社の正規化スキーマ)、'+
      'processed/benchmarks.json (5大商社平均 + 7社平均)。<br>'+
      '<strong>制約・補強方法:</strong> data/sogo-shosha/known-issues.md (取得不能項目の詳細とフェーズ2 拡張時の工数見積)。'+
    '</div>';
}

/* ── render dispatch ── */
function render(){
  if(!companies.length)return;
  if(tab==='exec')rExec();
  else if(tab==='shareholder')rShareholder();
  else if(tab==='valuation')rValuation();
  else if(tab==='financial')rFinancial();
  else if(tab==='activist')rActivist();
  else if(tab==='partnership')rPartnership();
  else if(tab==='detail')rDetail();
  else if(tab==='source')rSource();
}

/* ── nav binding ── */
function initNav(){
  var navInner=document.getElementById('mainNav');
  var btnL=document.getElementById('navScrollLeft');
  var btnR=document.getElementById('navScrollRight');
  function updateScrollBtns(){
    if(!navInner||!btnL||!btnR)return;
    btnL.classList.toggle('hidden',navInner.scrollLeft<=4);
    btnR.classList.toggle('hidden',navInner.scrollLeft+navInner.clientWidth>=navInner.scrollWidth-4);
  }
  if(navInner){navInner.addEventListener('scroll',updateScrollBtns);window.addEventListener('resize',updateScrollBtns);setTimeout(updateScrollBtns,100);}
  if(btnL)btnL.addEventListener('click',function(){navInner.scrollBy({left:-200,behavior:'smooth'});});
  if(btnR)btnR.addEventListener('click',function(){navInner.scrollBy({left:200,behavior:'smooth'});});
  document.querySelectorAll('.nav-item').forEach(function(n){n.addEventListener('click',function(){
    document.querySelectorAll('.nav-item').forEach(function(x){x.classList.remove('active');});
    n.classList.add('active');
    tab=n.dataset.tab;
    document.querySelectorAll('.section').forEach(function(s){s.classList.remove('active');});
    var sec=document.getElementById('sec-'+tab);
    if(sec)sec.classList.add('active');
    n.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
    render();
  });});
}

/* ── data loader (static JSON) ── */
async function loadData(){
  try{
    var fetches=TARGETS.map(function(t){
      return fetch('/data/sogo-shosha/processed/'+t.ticker+'-summary.json').then(function(r){
        if(!r.ok)throw new Error('HTTP '+r.status+' for '+t.ticker);
        return r.json();
      });
    });
    var refsFetches=[
      fetch('/data/sogo-shosha/refs/buffett-holdings.json').then(function(r){return r.ok?r.json():null;}).catch(function(){return null;}),
      fetch('/data/sogo-shosha/refs/strategic-initiatives.json').then(function(r){return r.ok?r.json():null;}).catch(function(){return null;}),
      fetch('/data/sogo-shosha/refs/company-narratives.json').then(function(r){return r.ok?r.json():null;}).catch(function(){return null;}),
    ];
    var summaries=await Promise.all(fetches);
    var refs=await Promise.all(refsFetches);
    buffett=refs[0];
    strategic=refs[1];
    narratives=refs[2];
    companies=summaries.map(function(s,i){
      var t=TARGETS[i];
      return{
        ticker:t.ticker,name:s.name||t.name,tier:s.tier||t.tier,
        marketCap:s.marketCap_JPY,
        summary:s,
        kpi:function(key){var k=s.kpis&&s.kpis[key];return k?k.value:null;},
      };
    });
    var countEl=document.getElementById('companyCount');
    if(countEl)countEl.textContent=companies.length;
    initNav();render();
  }catch(e){
    console.error('[sogo-shosha] data load failed:',e);
    var ex=document.getElementById('sec-exec');
    if(ex)ex.innerHTML='<div class="commentary danger"><strong>データ読み込みエラー:</strong> '+e.message+'<br><br>このページは <code>/data/sogo-shosha/processed/&lt;ticker&gt;-summary.json</code> を fetch で読み込みます。<code>file://</code> プロトコルでは CORS エラーになるため、<strong>ローカルでは <code>npx serve</code> または <code>python -m http.server</code> 経由で開いて下さい</strong>。</div>';
  }
}

/* ── CSV download ── */
document.addEventListener('DOMContentLoaded',function(){
  var dlCSV=document.querySelector('[data-action="downloadCSV"]');
  if(dlCSV)dlCSV.addEventListener('click',function(){
    if(!companies.length)return;
    var NL=String.fromCharCode(10);
    var keys=['doe','total_payout_ratio','buyback_amount','payout_ratio','crossheld_to_eq','roe','roic','pbr','per','ev_ebitda','resource_ratio','overseas_revenue','segment_roic','net_debt_ebitda','equity_ratio','interest_coverage','cash_to_mcap','large_holder_count','crossheld_reduction'];
    var labels=keys.map(function(k){return companies[0].summary.kpis[k].label;});
    var h=String.fromCharCode(0xFEFF)+'Code,Name,Tier,'+labels.join(',')+NL;
    companies.forEach(function(c){
      var row=[c.ticker,c.name,TIERS[c.tier]];
      keys.forEach(function(k){var v=c.kpi(k);row.push(v==null?'':v);});
      h+=row.join(',')+NL;
    });
    var b=new Blob([h],{type:'text/csv;charset=utf-8'});
    var a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='sogo-shosha-data.csv';a.click();
  });
  var dlPDF=document.querySelector('[data-action="downloadPDF"]');
  if(dlPDF)dlPDF.addEventListener('click',function(){window.print();});

  loadData();
});
})();
