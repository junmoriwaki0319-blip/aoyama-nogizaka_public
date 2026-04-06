/**
 * デジタルメディアセクター ダッシュボード
 * 7タブ: Executive Summary / サブセクター分析 / バリュエーション / アクティビスト / 成長とリスク / 個別企業 / データソース
 */
(function(){
'use strict';

const CATEGORIES={broadcasting:'放送',publishing:'出版・デジタルコンテンツ',platform:'デジタルプラットフォーム',video:'動画・映像制作',music:'音楽',adtech:'広告テック・メディアテック',print_dx:'印刷・DX',news:'新聞・ニュース'};
const CC={broadcasting:'#1a2d4f',publishing:'#2d7a4f',platform:'#b53a3a',video:'#9b8b6e',music:'#5a7fa8',adtech:'#c8946e',print_dx:'#5555aa',news:'#777'};
const CB={broadcasting:'badge-bizapp',publishing:'badge-fintech',platform:'badge-security',video:'badge-cx',music:'badge-vertical',adtech:'badge-other',print_dx:'badge-bizapp',news:'badge-cx'};

if(typeof companies==='undefined')window.companies=[];
var companies=window.companies;
var tab='exec',selComp=null;
const CH={};

/* ── helpers ── */
function g(id){return document.getElementById(id);}
function fmtM(v){if(v==null)return'-';if(v>=10000)return(v/10000).toFixed(1)+'兆円';return v.toLocaleString()+'億円';}
function shortName(n){return n.replace(/ホールディングス|HD|グループ/g,'').trim();}
function avg(a,k){const v=a.filter(c=>c[k]!=null);return v.length?(v.reduce((s,c)=>s+c[k],0)/v.length).toFixed(1):'-';}
function topN(a,k,n){return[...a].filter(c=>c[k]!=null).sort((x,y)=>y[k]-x[k]).slice(0,n);}
function nv(v,s='',f){if(v==null)return'-';if(f==='f1')return v.toFixed(1)+s;if(f==='f2')return v.toFixed(2)+s;if(f==='loc')return v.toLocaleString()+s;return v+s;}
function pn(v,s=''){if(v==null)return'-';return'<span class="'+(v>=0?'pos':'neg')+'">'+(v>0?'+':'')+v.toFixed(1)+s+'</span>';}
function mc(id,type,data,opts){var ctx=g(id);if(!ctx)return;var base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};if(['bar','line','scatter','bubble'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};CH[id]=new Chart(ctx,{type:type,data:data,options:dm(base,opts||{})});}
function dc(ids){ids.forEach(function(id){if(CH[id]){CH[id].destroy();delete CH[id];}});}
function dm(t,s){var o=Object.assign({},t);Object.keys(s).forEach(function(k){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]);else o[k]=s[k];});return o;}
function secH(n,t,d){return'<div class="sec-header"><div class="sec-num">SECTION '+n+'</div><div class="sec-title">'+t+'</div><div class="sec-desc">'+d+'</div></div>';}
function kpi(l,v,s,cls){return'<div class="kpi-card '+cls+'"><div class="kpi-label">'+l+'</div><div class="kpi-value">'+v+'</div>'+(s?'<div class="kpi-sub">'+s+'</div>':'')+'</div>';}
function rankCard(t,items,fn){return'<div class="ranking-card"><div class="ranking-title">'+t+'</div>'+items.map(function(c,i){return'<div class="ranking-row"><span><span class="ranking-num">'+(i+1)+'</span>'+shortName(c.name)+'</span><span style="font-weight:600">'+fn(c)+'</span></div>';}).join('')+'</div>';}
function catOf(c){return CATEGORIES[c.category]||c.category;}
function byCat(cat){return companies.filter(function(c){return c.category===cat;});}


function enableTableSort(tbl,data,cols,rowFn,cb){
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
        var va=a[def.key],vb=b[def.key];
        if(va==null)va=def.type==='num'?-Infinity:'';
        if(vb==null)vb=def.type==='num'?-Infinity:'';
        if(def.type==='num')return sortD==='desc'?vb-va:va-vb;
        return sortD==='desc'?String(vb).localeCompare(String(va)):String(va).localeCompare(String(vb));
      });
      tbl.querySelector('tbody').innerHTML=sorted.map(rowFn).join('');
      if(cb)cb(tbl);
    });
  });
}

/* ── rExec ── */
function rExec(){
  var el=g('sec-exec');
  var tm=companies.reduce(function(s,c){return s+(c.marketCap||0);},0);
  var aOP=avg(companies,'operatingMargin'),aROE=avg(companies,'roe'),aPER=avg(companies,'per'),aPBR=avg(companies,'pbr');
  var topMC=topN(companies,'marketCap',3),topOP=topN(companies,'operatingMargin',3),topROE=topN(companies,'roe',3);

  el.innerHTML=
    secH('01','Executive Summary','日本上場デジタルメディア企業の全体概況と投資機会')+
    '<div class="commentary">'+
      '<strong>セクター概況 (2026年4月基準):</strong> 対象<strong>'+companies.length+'社</strong>の合計時価総額は<strong>'+fmtM(tm)+'</strong>。'+
      '放送・出版・デジタルプラットフォーム・映像制作・音楽・印刷DXの8カテゴリを横断的にカバー。'+
      '平均営業利益率<strong>'+aOP+'%</strong>、平均ROE<strong>'+aROE+'%</strong>。<br><br>'+
      '<strong>構造変化の3つの潮流:</strong><br>'+
      '<strong>(1) 放送→配信シフト:</strong> TBS HDの配信広告収入は前年比+45.4%増（82.4億円）、有料配信+36.5%（121.5億円）で高成長。一方で地上波広告収入は逓減傾向。在京キー局は放送外収入比率の拡大が経営課題。<br>'+
      '<strong>(2) 出版→デジタルコンテンツ:</strong> KADOKAWA（フロム・ソフトウェア保有）のようなIP総合企業化が進行。電子書籍市場は年率6%成長。DNP・凸版の印刷大手もDX転換を加速。<br>'+
      '<strong>(3) 検索広告の減速とAIシフト:</strong> LINEヤフーの検索広告は前年比-13.2%と減速。GoogleのAI Overviews拡大でオーガニック検索CTRが最大58%低下。クリエイターエコノミー（note等）やLINE公式アカウント広告が新たな成長軸。<br><br>'+
      '<strong>マネジメントへの5つの提言 ― メディア産業の構造転換を乗り越える経営設計:</strong><br><br>'+
      '<strong>(1) 放送外収入比率のKPI化と配信事業への経営資源シフト:</strong> '+
      'TBS HDの配信広告+45.4%増は、地上波依存からの脱却が可能であることを実証した。'+
      '各局は<strong>放送外収入比率を取締役会KPIに設定</strong>し、配信・EC・IP展開への投資配分を明確化すべき。'+
      '3年後に放送外収入比率50%超を目指す中計の策定を推奨。<br>'+
      '<strong>(2) 不動産含み益の戦略的活用 ― スピンオフ or 証券化:</strong> '+
      'フジ・メディアHD（お台場）に限らず、在京キー局は都心一等地に本社・スタジオを保有。'+
      'ダルトンの取締役12名選任提案は否決されたが、<strong>不動産事業分離の議論は継続</strong>する。'+
      '不動産のREIT化・J-REIT組成による含み益の顕在化と放送事業への再投資は、PBR改善と成長投資を両立するシナリオ。'+
      '「攻められる前に自ら動く」先手のアセットマネジメントが経営陣に求められる。<br>'+
      '<strong>(3) AI時代の「メディア信頼性」のマネタイズ:</strong> '+
      '生成AIの普及でフェイクニュース・コンテンツ汚染が加速する中、'+
      '既存メディアの<strong>「取材力・編集力に裏付けられた信頼性」はむしろ希少資源</strong>となる。'+
      'ユーザベース(NewsPicks)のような有料ビジネスメディアの成功、noteのクリエイターエコノミーの台頭は、'+
      '「信頼×専門性」への課金モデルが成立することを示す。AIが代替できない領域への集中投資を。<br>'+
      '<strong>(4) 印刷大手のDX転換を加速する「選択と集中」:</strong> '+
      'DNP・凸版は売上10兆円規模の巨大企業だが、印刷市場の構造的縮小は不可逆。'+
      'DNPのデジタル教科書・半導体フォトマスク、凸版のIoTパッケージング等の<strong>非印刷事業のカーブアウト・IPO</strong>も選択肢。'+
      '「印刷会社」から「情報コミュニケーション企業」への名実ともの転換を取締役会が主導すべき。<br>'+
      '<strong>(5) クロスボーダーIP展開の座組設計:</strong> '+
      'KADOKAWAのシンガポールSOZO社子会社化、ソニー出資によるフロム・ソフトウェア間接保有は、'+
      '日本のIP（アニメ・ゲーム・出版）のグローバル展開における資本連携の型を示す。'+
      'メディア企業は<strong>「コンテンツを作る側」から「IP資産を運用する側」</strong>へポジションを転換し、'+
      '海外パートナーとの共同製作・ライセンス・M&Aの座組を経営戦略の中核に据えるべき。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('対象企業数',companies.length+'社','8カテゴリ','c-navy')+
      kpi('時価総額合計',fmtM(tm),'','c-navy')+
      kpi('平均営業利益率',aOP+'%','','c-gold')+
      kpi('平均ROE',aROE+'%','','c-green')+
      kpi('平均PER',aPER+'倍','','c-navy')+
      kpi('平均PBR',aPBR+'倍','','c-gold')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 億円</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PBRランキング TOP20</div><div class="chart-panel-sub">赤線=PBR 1.0倍（解散価値）/ 1.0倍未満はアクティビスト注目領域</div><div class="chart-area tall"><canvas id="exPBR"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">カテゴリ別 時価総額構成比</div><div class="chart-area"><canvas id="exPie"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">主要ランキング</div>'+
        '<div style="display:grid;gap:10px;padding-top:8px;">'+
          rankCard('時価総額',topMC,function(c){return fmtM(c.marketCap);})+
          rankCard('営業利益率',topOP,function(c){return nv(c.operatingMargin,'%','f1');})+
          rankCard('ROE',topROE,function(c){return nv(c.roe,'%','f1');})+
        '</div></div>'+
    '</div>';

  dc(['exMC','exPBR','exPie']);
  var s15=[].concat(companies).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);}).slice(0,15);
  mc('exMC','bar',{labels:s15.map(function(c){return shortName(c.name);}),datasets:[{data:s15.map(function(c){return c.marketCap;}),backgroundColor:s15.map(function(c){return CC[c.category]||'#777';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return fmtM(v);}}}});

  var segMC={};companies.forEach(function(c){segMC[c.category]=(segMC[c.category]||0)+(c.marketCap||0);});
  mc('exPie','doughnut',{labels:Object.keys(segMC).map(function(k){return CATEGORIES[k];}),datasets:[{data:Object.values(segMC),backgroundColor:Object.keys(segMC).map(function(k){return CC[k];}),borderWidth:0}]},{plugins:{legend:{position:'right'},datalabels:{display:true,color:'#fff',font:{size:10,weight:600},formatter:function(v,ctx){var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);return(v/t*100).toFixed(1)+'%';}}}});

  var pbrTop=[].concat(companies).filter(function(c){return c.pbr!=null&&c.pbr>0&&c.pbr<15;}).sort(function(a,b){return a.pbr-b.pbr;}).slice(0,20);
  mc('exPBR','bar',{labels:pbrTop.map(function(c){return shortName(c.name);}),datasets:[{data:pbrTop.map(function(c){return c.pbr;}),backgroundColor:pbrTop.map(function(c){return c.pbr<1.0?'rgba(181,58,58,0.7)':'rgba(26,45,79,0.5)';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(2)+'倍';}},annotation:{annotations:{pbr1:{type:'line',xMin:1.0,xMax:1.0,borderColor:'#b53a3a',borderWidth:2,borderDash:[4,4],label:{display:true,content:'PBR 1.0倍',position:'start',backgroundColor:'rgba(181,58,58,0.85)',color:'#fff',font:{size:9}}}}}},scales:{x:{min:0,title:{display:true,text:'PBR (倍)'}}}});
}

/* ── rSubsector ── */
function rSubsector(){
  var el=g('sec-subsector');
  var catKeys=Object.keys(CATEGORIES);
  var catStats=catKeys.map(function(cat){
    var cs=byCat(cat);
    return{cat:cat,label:CATEGORIES[cat],count:cs.length,
      totalMcap:cs.reduce(function(s,c){return s+(c.marketCap||0);},0),
      avgOPM:avg(cs,'operatingMargin'),avgROE:avg(cs,'roe'),avgPER:avg(cs,'per'),avgPBR:avg(cs,'pbr')};
  });

  el.innerHTML=
    secH('02','サブセクター分析','8カテゴリ別の市場構成と財務指標比較')+
    '<div class="commentary">'+
      '<strong>カテゴリ別概況:</strong> 放送局（在京キー局+BS/CS）は安定した広告収入基盤を持つが成長は鈍化。'+
      'デジタルプラットフォーム（LINEヤフー等）は時価総額でセクターを牽引。'+
      '印刷・DX（DNP/凸版）は伝統企業のDX転換事例として注目。'+
      '広告テックは成長性は高いが利益率の企業間格差が顕著。'+
    '</div>'+
    '<div class="kpi-grid">'+
      catStats.map(function(s){return kpi(s.label,s.count+'社','時価総額計 '+fmtM(s.totalMcap),'c-navy');}).join('')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">カテゴリ別 平均指標比較</div><div class="chart-panel-sub">営業利益率・ROE・PER・PBR</div><div class="chart-area tall"><canvas id="ssBar"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">カテゴリ別 時価総額分布</div><div class="chart-panel-sub">企業ごとの時価総額（対数スケール）</div><div class="chart-area tall"><canvas id="ssDist"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">カテゴリ別 企業数と時価総額</div><div class="chart-area"><canvas id="ssMcap"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PER vs PBR（カテゴリ色分け）</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area"><canvas id="ssPERPBR"></canvas></div></div>'+
    '</div>'+
    catKeys.map(function(cat){
      var cs=byCat(cat).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);});
      if(!cs.length)return'';
      return '<div class="table-panel"><div class="table-header"><div class="table-header-title">'+CATEGORIES[cat]+' ('+cs.length+'社)</div></div><div class="table-scroll"><table>'+
        '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
        '<tbody>'+cs.map(function(c){
          return '<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
            '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
            '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
        }).join('')+'</tbody></table></div></div>';
    }).join('');

  bindRows(el);
  dc(['ssBar','ssDist','ssMcap','ssPERPBR']);

  mc('ssBar','bar',{labels:catStats.map(function(s){return s.label;}),datasets:[
    {label:'営業利益率(%)',data:catStats.map(function(s){return parseFloat(s.avgOPM)||0;}),backgroundColor:'#1a2d4f'},
    {label:'ROE(%)',data:catStats.map(function(s){return parseFloat(s.avgROE)||0;}),backgroundColor:'#2d7a4f'},
    {label:'PBR(倍)',data:catStats.map(function(s){return parseFloat(s.avgPBR)||0;}),backgroundColor:'#9b8b6e'}
  ]},{plugins:{datalabels:{display:true,anchor:'end',align:'top',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1);}}}});

  mc('ssMcap','bar',{labels:catStats.map(function(s){return s.label;}),datasets:[{label:'時価総額合計(億円)',data:catStats.map(function(s){return s.totalMcap;}),backgroundColor:catKeys.map(function(k){return CC[k];}),borderWidth:0}]},{plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'top',color:'#999',font:{size:9},formatter:function(v){return fmtM(v);}}}});

  var distData=catKeys.map(function(cat){return{label:CATEGORIES[cat],data:byCat(cat).map(function(c){return{x:CATEGORIES[cat],y:c.marketCap||0,name:c.name};}),backgroundColor:CC[cat]+'88',borderColor:CC[cat],borderWidth:1,pointRadius:6};});
  mc('ssDist','scatter',{datasets:distData.map(function(ds,i){return{label:ds.label,data:ds.data.map(function(d){return{x:i,y:d.y};}),backgroundColor:ds.backgroundColor,borderColor:ds.borderColor,borderWidth:1,pointRadius:6};})},{scales:{x:{type:'category',labels:catKeys.map(function(k){return CATEGORIES[k];}),ticks:{font:{size:9}}},y:{type:'logarithmic',title:{display:true,text:'時価総額 (億円)'}}},plugins:{tooltip:{callbacks:{label:function(ctx){var cat=catKeys[ctx.dataIndex!==undefined?ctx.datasetIndex:0];var cs=byCat(cat).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);});var c=cs[ctx.dataIndex];return c?c.name+': '+fmtM(c.marketCap):'';}}}}});

  var mx2=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));
  mc('ssPERPBR','bubble',{datasets:Object.keys(CATEGORIES).map(function(cat){return{label:CATEGORIES[cat],data:byCat(cat).filter(function(c){return c.per!=null&&c.pbr!=null;}).map(function(c){return{x:c.per,y:c.pbr,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx2)*22),name:c.name};}),backgroundColor:CC[cat]+'77',borderColor:CC[cat],borderWidth:1};})},{scales:{x:{title:{display:true,text:'PER (倍)'},min:0,max:80},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': PER'+x.raw.x+'倍 / PBR'+x.raw.y+'倍';}}}}});
}

/* ── rValuation ── */
function rValuation(){
  var el=g('sec-valuation');
  var sorted=[].concat(companies).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);});
  el.innerHTML=
    secH('03','財務指標・バリュエーション','PER/PBR/ROE/営業利益率の横断比較')+
    '<div class="commentary">'+
      '<strong>バリュエーション分析:</strong> セクター平均PERは<strong>'+avg(companies,'per')+'倍</strong>、PBRは<strong>'+avg(companies,'pbr')+'倍</strong>。'+
      'PBR-ROE間の正の相関が確認され、<strong>ROE改善が株価評価に直結</strong>する構造。'+
      'ANYCOLOR（PBR '+nv(companies.find(function(c){return c.ticker==='5032';})?.pbr,'倍','f2')+'）やカプコン等の高ROE企業は高PBR評価。'+
      '一方、PBR1倍割れの中小型株にはアクティビストの介入余地がある。'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">PBR vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 右上=高収益×高評価</div><div class="chart-area tall"><canvas id="vlPBR"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PER vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="vlPER"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">営業利益率分布</div><div class="chart-area"><canvas id="vlHist"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">ROEランキング TOP15</div><div class="chart-area"><canvas id="vlROE"></canvas></div></div>'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">バリュエーション一覧（時価総額順）</div></div><div class="table-scroll"><table id="tblVal">'+
      '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>カテゴリ</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+(CB[c.category]||'')+'">'+catOf(c)+'</span></td>'+
          '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
          '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';

  bindRows(el);var tblVal=g("tblVal");if(tblVal)enableTableSort(tblVal,sorted,[{key:'ticker',type:'str'},{key:'name',type:'str'},{key:'category',type:'str'},{key:'marketCap',type:'num'},{key:'operatingMargin',type:'num'},{key:'roe',type:'num'},{key:'per',type:'num'},{key:'pbr',type:'num'},{key:'price',type:'num'}],function(c){return'<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td><td><span class="badge ">'+(CATEGORIES[c.category]||c.category)+'</span></td><td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td><td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';},function(t){bindRows(el);});dc(["vlPBR",'vlPER','vlHist','vlROE']);
  var mx=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));

  mc('vlPBR','bubble',{datasets:Object.keys(CATEGORIES).map(function(cat){return{label:CATEGORIES[cat],data:byCat(cat).filter(function(c){return c.roe!=null&&c.pbr!=null&&c.roe>-50&&c.roe<100&&c.pbr>0&&c.pbr<20;}).map(function(c){return{x:c.roe,y:c.pbr,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx)*25),name:c.name};}),backgroundColor:CC[cat]+'77',borderColor:CC[cat],borderWidth:1};})},{scales:{x:{title:{display:true,text:'ROE (%)'},min:-50,max:100},y:{title:{display:true,text:'PBR (倍)'},min:0,max:20}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': ROE'+x.raw.x+'% / PBR'+x.raw.y+'倍';}}}}});

  mc('vlPER','bubble',{datasets:Object.keys(CATEGORIES).map(function(cat){return{label:CATEGORIES[cat],data:byCat(cat).filter(function(c){return c.per!=null&&c.operatingMargin!=null&&c.operatingMargin>-50&&c.operatingMargin<60&&c.per>0&&c.per<80;}).map(function(c){return{x:c.operatingMargin,y:c.per,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx)*25),name:c.name};}),backgroundColor:CC[cat]+'77',borderColor:CC[cat],borderWidth:1};})},{scales:{x:{title:{display:true,text:'営業利益率 (%)'},min:-50,max:60},y:{title:{display:true,text:'PER (倍)'},min:0,max:80}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': OPM'+x.raw.x+'% / PER'+x.raw.y+'倍';}}}}});

  var bins=[-20,-5,0,5,10,15,20,30,50];var hist=bins.slice(0,-1).map(function(_,i){return companies.filter(function(c){return c.operatingMargin!=null&&c.operatingMargin>=bins[i]&&c.operatingMargin<bins[i+1];}).length;});
  mc('vlHist','bar',{labels:bins.slice(0,-1).map(function(b,i){return b+'~'+bins[i+1]+'%';}),datasets:[{data:hist,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});

  var roeTop=topN(companies,'roe',15);
  mc('vlROE','bar',{labels:roeTop.map(function(c){return shortName(c.name);}),datasets:[{data:roeTop.map(function(c){return c.roe;}),backgroundColor:roeTop.map(function(c){return CC[c.category]||'#777';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1)+'%';}}}});
}

/* ── rActivist ── */
function rActivist(){
  var el=g('sec-activist');
  el.innerHTML=
    secH('04','アクティビスト・M&A','メディア企業への株主提案とM&A・業界再編')+
    '<div class="commentary" style="border-left-color:var(--gold)">'+
      '<div class="commentary-title">アクティビスト動向 — メディア企業への介入</div>'+
      '<p>2025年6月総会: 全体111社に株主提案（過去最多）。メディアセクターでは不動産含み益を持つ放送局が標的に。</p>'+
      '<p><strong>ダルトン・インベストメンツ → フジ・メディアHD(4676):</strong> 2025年6月に取締役12名選任の株主提案。全員否決も会社側全11名は8割超の賛成で可決。'+
      '不動産事業分離（お台場エリア）要求が継続論点。放送外収入比率の拡大と不動産スピンオフが焦点。</p>'+
      '<p><strong>放送局の不動産含み益:</strong> 在京キー局は放送法に基づく免許事業であり、都心一等地に本社・スタジオを保有。'+
      'PBR1倍割れの放送局は不動産価値だけでも時価総額を上回るケースがあり、アクティビストの関心を集める。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--navy)">'+
      '<div class="commentary-title">M&A・業界再編</div>'+
      '<p><strong>テレビ東京HD(9413):</strong> Q-nine社を子会社化（2026年1月）。コンテンツ制作力強化。</p>'+
      '<p><strong>KADOKAWA(9468):</strong> シンガポールSOZO社を子会社化（2025年11月）。海外IP展開を加速。ソニーが出資しフロム・ソフトウェア（エルデンリング）を間接保有。</p>'+
      '<p><strong>TBS HD(9401):</strong> 配信事業強化のためのM&A・提携を積極化。配信広告収入+45.4%増。</p>'+
      '<p><strong>印刷大手のDX転換:</strong> DNP・凸版がデジタル教科書、電子書籍基盤、IoTパッケージング等へ事業転換。伝統的印刷市場の縮小を補う成長投資。</p>'+
    '</div>';
}

/* ── rGrowth ── */
function rGrowth(){
  var el=g('sec-growth');
  el.innerHTML=
    secH('05','成長ドライバーとリスク','配信シフト・AI・規制環境の分析')+
    '<div class="commentary" style="border-left-color:var(--green)">'+
      '<div class="commentary-title">成長ドライバー</div>'+
      '<p><strong>(1) 放送→配信シフト:</strong> TBS HDの配信広告収入+45.4%増、有料配信+36.5%増が象徴的。U-NEXT HOLDINGSは国内有料動画配信トップクラス。放送局は「テレビ×配信×EC」の三位一体モデルへ転換中。</p>'+
      '<p><strong>(2) クリエイターエコノミー:</strong> note（クリエイターの有料記事販売）やLINEヤフー（LINE公式アカウント広告の成長）など、個人/中小事業者のコンテンツ収益化プラットフォームが拡大。</p>'+
      '<p><strong>(3) 印刷大手のDX転換:</strong> DNPはデジタル教科書・honto運営・半導体フォトマスクに注力。凸版はデジタルコンテンツ基盤・IoTパッケージング。伝統企業のDX転換が新たな成長エンジンに。</p>'+
      '<p><strong>(4) IP総合企業の台頭:</strong> KADOKAWA（出版×アニメ×ゲーム/フロム・ソフトウェア）やサイバーエージェント（ABEMA×広告×Cygames）など、メディアの垣根を越えたIP総合戦略が高評価。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">構造変化と課題</div>'+
      '<p><strong>検索広告の減速:</strong> LINEヤフーの検索広告は前年比-13.2%。GoogleのAI Overviews拡大でオーガニック検索CTRが最大58%低下。検索依存型メディアの収益モデルが根本から問われている。</p>'+
      '<p><strong>放送収入の構造的減少:</strong> 地上波広告収入は年率2-3%減少トレンド。配信収入が補完するが、規模では未だ放送収入の1/10程度。転換期の利益率悪化に注意。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">リスク要因</div>'+
      '<p><strong>(1) AIによるコンテンツ自動生成:</strong> 生成AIがニュース記事・広告クリエイティブ・映像制作の自動化を加速。メディア制作の付加価値が問われる局面。</p>'+
      '<p><strong>(2) 放送法規制と外資規制:</strong> 放送局は外国人議決権制限（20%未満）等の規制下にあり、M&Aや資本政策に制約。</p>'+
      '<p><strong>(3) 広告市場の二極化:</strong> デジタル広告はGoogle/Meta/Amazon等グローバルプラットフォーマーに集中。日本のメディア企業はデータ基盤の差で劣後するリスク。</p>'+
      '<p><strong>(4) 出版市場の縮小:</strong> 紙媒体の販売部数は年率5-8%減少。電子書籍は成長するが紙の減少を完全には補えない。</p>'+
    '</div>';
}

/* ── rDetail ── */
function rDetail(){
  var el=g('sec-detail');
  if(!selComp)selComp=companies[0];
  var c=selComp;
  var peers=companies.filter(function(x){return x.category===c.category&&x.ticker!==c.ticker;});
  el.innerHTML=
    secH('06','個別企業分析','選択企業の詳細指標と同カテゴリ比較')+
    '<div class="inline-filters"><span class="f-label">企業選択</span>'+
      '<select id="dtSel">'+companies.map(function(x){return'<option value="'+x.ticker+'"'+(x.ticker===c.ticker?' selected':'')+'>'+x.ticker+' '+x.name+'</option>';}).join('')+'</select>'+
    '</div>'+
    '<div style="font-family:\'Noto Serif JP\',serif;font-size:1.4rem;font-weight:700;color:var(--navy);margin-bottom:4px;">'+c.name+'</div>'+
    '<div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:12px;">'+c.ticker+' / '+catOf(c)+'</div>'+
    '<div class="kpi-grid">'+
      kpi('株価',nv(c.price,'円','loc'),'','c-navy')+
      kpi('時価総額',fmtM(c.marketCap),'','c-navy')+
      kpi('営業利益率',nv(c.operatingMargin,'%','f1'),'','c-gold')+
      kpi('ROE',nv(c.roe,'%','f1'),'','c-green')+
      kpi('PER',nv(c.per,'倍','f1'),'','c-navy')+
      kpi('PBR',nv(c.pbr,'倍','f2'),'','c-gold')+
    '</div>'+
    (peers.length?
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">同カテゴリ比較 — '+catOf(c)+'</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th style="text-align:left">企業名</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th></tr></thead>'+
      '<tbody><tr class="row-hl"><td><strong>'+shortName(c.name)+'</strong></td><td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td><td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td></tr>'+
      peers.sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);}).map(function(p){
        return'<tr><td>'+shortName(p.name)+'</td><td>'+nv(p.marketCap,'','loc')+'</td><td>'+nv(p.operatingMargin,'%','f1')+'</td><td>'+nv(p.roe,'%','f1')+'</td><td>'+nv(p.per,'倍','f1')+'</td><td>'+nv(p.pbr,'倍','f2')+'</td></tr>';
      }).join('')+
      '</tbody></table></div></div>':'');

  g('dtSel').addEventListener('change',function(e){selComp=companies.find(function(c){return c.ticker===e.target.value;});rDetail();});
}

/* ── rSource ── */
function rSource(){
  var el=g('sec-source');
  el.innerHTML=
    secH('A','データソース・方法論','データ取得元と更新方針')+
    '<div class="commentary">'+
      '<strong>データソース:</strong> 株価・時価総額・営業利益率・ROE・PER・PBRはYahoo Finance (yahoo-finance2)から取得した実データ。'+
      '基準日: 2026年4月取得時点の直近終値。<br>'+
      '<strong>カテゴリ分類:</strong> 各社の主力事業に基づき8カテゴリに分類（放送 / 出版 / プラットフォーム / 動画・映像 / 音楽 / 広告テック / 印刷・DX / ニュース）。一部企業は他セクターDB（広告・ゲーム）にも重複掲載。<br>'+
      '<strong>更新頻度:</strong> 四半期ごとにyahoo-finance2でデータ更新 → Firestoreへアップロード。<br>'+
      '<strong>注意事項:</strong> 本レポートは情報提供を目的としたものであり、特定の金融商品の売買を推奨するものではありません。投資判断はご自身の責任においてお願いいたします。'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">全企業データ一覧</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>カテゴリ</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
      '<tbody>'+companies.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td>'+c.name+'</td><td><span class="badge '+(CB[c.category]||'')+'">'+catOf(c)+'</span></td>'+
          '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
          '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';
}

/* ── navigation & routing ── */
function render(){
  var fns={exec:rExec,subsector:rSubsector,valuation:rValuation,activist:rActivist,growth:rGrowth,detail:rDetail,source:rSource};
  if(fns[tab])fns[tab]();
}

function bindRows(el){
  el.querySelectorAll('.clickable-row').forEach(function(tr){
    tr.addEventListener('click',function(){
      selComp=companies.find(function(c){return c.ticker===tr.dataset.code;});
      document.querySelector('.nav-item[data-tab="detail"]').click();
    });
  });
}

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

/* ── Firestore integration ── */
window.loadPremiumData=async function(){
  if(!window.firebaseDb)return;
  try{
    var m=await import('https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js');
    var db=window.firebaseDb;
    var snap=await m.getDoc(m.doc(db,'premiumContent','digital-media-companies'));
    if(snap.exists()){companies=snap.data().companies||[];window.companies=companies;}
  }catch(e){console.error('Premium data load failed:',e);return;}
  var countEl=document.getElementById('companyCount');
  if(countEl)countEl.textContent=companies.length;
  initNav();render();
};


// CSV/PDF download
var dlCSV=document.querySelector('[data-action="downloadCSV"]');
if(dlCSV)dlCSV.addEventListener('click',function(){
  if(!companies.length)return;
  var tierOrCat=companies[0].tier?'Tier':'カテゴリ';
  var csv='﻿'+'コード,企業名,'+tierOrCat+',時価総額(億円),営業利益率(%),ROE(%),PER,PBR,株価
';
  companies.forEach(function(c){csv+=c.ticker+','+c.name+','+(c.tier||c.category)+','+(c.marketCap||'')+','+(c.operatingMargin||'')+','+(c.roe||'')+','+(c.per||'')+','+(c.pbr||'')+','+(c.price||'')+'
';});
  var blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='sector-data.csv';a.click();
});
var dlPDF=document.querySelector('[data-action="downloadPDF"]');
if(dlPDF)dlPDF.addEventListener('click',function(){window.print();});

document.addEventListener('DOMContentLoaded',function(){initNav();});
})();
