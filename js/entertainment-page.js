/**
 * ゲーム・コンテンツセクター ダッシュボード
 * 7タブ: Executive Summary / サブセクター分析 / バリュエーション / アクティビスト / 成長とリスク / 個別企業 / データソース
 */
(function(){
'use strict';

const CATEGORIES={console:'コンソール・PC',mobile:'モバイルゲーム','anime-ip':'アニメ・映像・IP',vtuber:'VTuber・配信',support:'ゲーム支援'};
const CC={console:'#1a2d4f',mobile:'#2d7a4f','anime-ip':'#b53a3a',vtuber:'#9b8b6e',support:'#5a7fa8'};
const CB={console:'badge-publisher',mobile:'badge-mobile','anime-ip':'badge-anime',vtuber:'badge-vtuber',support:'badge-tools'};

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

/* ── rExec ── */
function rExec(){
  var el=g('sec-exec');
  var tm=companies.reduce(function(s,c){return s+(c.marketCap||0);},0);
  var aOP=avg(companies,'operatingMargin'),aROE=avg(companies,'roe'),aPER=avg(companies,'per'),aPBR=avg(companies,'pbr');
  var topMC=topN(companies,'marketCap',3),topOP=topN(companies,'operatingMargin',3),topROE=topN(companies,'roe',3);

  el.innerHTML=
    secH('01','Executive Summary','日本上場ゲーム・コンテンツ企業の全体概況と投資機会')+
    '<div class="commentary">'+
      '<strong>セクター概況 (2026年4月基準):</strong> 対象<strong>'+companies.length+'社</strong>の合計時価総額は<strong>'+fmtM(tm)+'</strong>。'+
      '日本ゲーム市場は2025年約2.6兆円（ファミ通ゲーム白書）、モバイルIAP収益が1.6兆円超と市場中核。'+
      'ユニークユーザー推定5,000万人超、グローバルiOS収益でアジア2位。'+
      '平均営業利益率<strong>'+aOP+'%</strong>、平均ROE<strong>'+aROE+'%</strong>。<br><br>'+
      '<strong>マネジメントへの5つの提言 ― ゲーム・コンテンツ産業の次の10年を設計する:</strong><br><br>'+
      '<strong>(1) IP価値の「可視化」と資本市場への訴求:</strong> '+
      'ゲーム企業の株価は開発パイプラインの成否に大きく左右されるが、IP価値そのものが適切に評価されていないケースが多い。'+
      'バンダイナムコのIP別売上開示（ドラゴンボール4,000億円超、ガンダム1,700億円超）は好事例。'+
      '自社IP群の<strong>LTV（生涯価値）推計とロイヤリティ収入の中長期予測</strong>を開示し、「開発会社」ではなく「IP資産運用会社」としての再評価を促すべき。<br>'+
      '<strong>(2) アクティビスト対応の先手設計:</strong> '+
      '3D Investment Partners（スクエニHD 14.36%）、ストラテジックキャピタル（ガンホー 5.4%）、サウジAyyal（カプコン・バンナム・コナミ等一斉取得）の動きは、'+
      '豊富なIP資産×安定CF×PBR割安の組み合わせがアクティビストの標的となることを示す。'+
      '取締役会は<strong>「攻められる前に自ら変わる」プロアクティブなガバナンス改革</strong>（業績連動報酬、政策保有株の縮減、IR強化）を推進すべき。<br>'+
      '<strong>(3) 海外売上比率50%超の戦略的目標設定:</strong> '+
      'カプコン（海外80%超）の高バリュエーションは、グローバル展開力がプレミアムの源泉であることを証明している。'+
      'モバイルゲーム企業（MIXI、DeNA等）は国内IAPモデルの成熟に直面しており、'+
      '<strong>「ポケポケ」（DeNA/ポケモン社、海外DL比率70%超）</strong>のような海外先行型タイトル戦略への転換が生存条件。<br>'+
      '<strong>(4) 開発費高騰への構造対応 ― AI活用とミドルタイトル戦略:</strong> '+
      'AAA級ゲーム1タイトルの開発費が100〜300億円に達する中、全てを大作に賭けるモデルは持続困難。'+
      '生成AIによる開発コスト20〜30%削減の試算を前提に、<strong>ミドルバジェット×高回転のポートフォリオ戦略</strong>と、'+
      'AI活用による開発プロセスのリエンジニアリングを同時に推進すべき。<br>'+
      '<strong>(5) VTuber・クリエイターエコノミーとの協業設計:</strong> '+
      'ANYCOLOR（営業利益率37.2%）やカバーの急成長は、従来のゲーム企業が取り込めていなかった「コンテンツ消費のソーシャル化」を体現。'+
      'ゲーム×VTuberのコラボマーケティングを超え、<strong>IPの共創・ファンコミュニティの資産化</strong>まで踏み込んだパートナーシップ設計が中期的な競争優位となる。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('対象企業数',companies.length+'社','5カテゴリ','c-navy')+
      kpi('時価総額合計',fmtM(tm),'','c-navy')+
      kpi('平均営業利益率',aOP+'%','','c-gold')+
      kpi('平均ROE',aROE+'%','','c-green')+
      kpi('平均PER',aPER+'倍','','c-navy')+
      kpi('平均PBR',aPBR+'倍','','c-gold')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 億円</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PBRランキング TOP20</div><div class="chart-panel-sub">赤線=PBR 1.0x（解散価値）/ 1.0x未満はアクティビスト注目領域</div><div class="chart-area tall"><canvas id="exPBR"></canvas></div></div>'+
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
  mc('exPBR','bar',{labels:pbrTop.map(function(c){return shortName(c.name);}),datasets:[{data:pbrTop.map(function(c){return c.pbr;}),backgroundColor:pbrTop.map(function(c){return c.pbr<1.0?'rgba(181,58,58,0.7)':'rgba(26,45,79,0.5)';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(2)+'x';}},annotation:{annotations:{pbr1:{type:'line',xMin:1.0,xMax:1.0,borderColor:'#b53a3a',borderWidth:2,borderDash:[4,4],label:{display:true,content:'PBR 1.0x',position:'start',backgroundColor:'rgba(181,58,58,0.85)',color:'#fff',font:{size:9}}}}}},scales:{x:{min:0,title:{display:true,text:'PBR (倍)'}}}});
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
    secH('02','サブセクター分析','5カテゴリ別の市場構成と財務指標比較')+
    '<div class="commentary">'+
      '<strong>カテゴリ別概況:</strong> コンソール・PCパブリッシャーが時価総額の大半を占めるが、'+
      'VTuber・配信は高い営業利益率（ANYCOLOR 37.2%）と成長率を誇る新興カテゴリ。'+
      'モバイルゲームはIAPモデルの成熟化に伴い選別が進行中。アニメ・IP企業は映像化・テーマパーク展開によるIP多角化が評価指標。'+
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
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">バリュエーション一覧（時価総額順）</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>カテゴリ</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+(CB[c.category]||'')+'">'+catOf(c)+'</span></td>'+
          '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
          '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';

  bindRows(el);dc(['vlPBR','vlPER','vlHist','vlROE']);
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
    secH('04','アクティビスト動向・IP戦略','株主提案の急増とIP多角化による企業価値向上')+
    '<div class="commentary" style="border-left-color:var(--gold)">'+
      '<div class="commentary-title">アクティビスト動向 — ゲーム企業への介入が本格化</div>'+
      '<p>2025年6月総会: 全体111社に株主提案（過去最多）。ゲーム・コンテンツ企業への介入が顕著に増加。</p>'+
      '<p><strong>3D Investment Partners → スクウェア・エニックスHD(9684):</strong> 株式14.36%買い増し、2025年12月に100ページ超の経営課題プレゼンテーションを公表。'+
      '開発パイプラインの選択と集中、非中核IP売却、株主還元強化を要求。</p>'+
      '<p><strong>ストラテジックキャピタル → ガンホー(3765):</strong> 株式5.4%保有。CEO報酬の業績連動化、自己株式全数消却、タイトル別売上開示等を提案。特設サイト「再起の処方箋」を開設。</p>'+
      '<p><strong>ValueAct Capital → 任天堂(7974):</strong> 1.1億ドル以上のポジション構築。デジタル化・事業多角化を推進。</p>'+
      '<p><strong>サウジAyyal First Investment:</strong> カプコン、スクエニHD、バンナムHD、ネクソン、東映、コーエーテクモHDに一斉大量取得。'+
      '中東マネーがゲームIPに注目するグローバルトレンド。</p>'+
    '</div>'+
    '<div class="commentary">'+
      '<div class="commentary-title">IP多角化戦略 — ゲーム→映画→アニメ→テーマパークの収益連鎖</div>'+
      '<p><strong>マリオ映画:</strong> 全世界興行収入13億ドル超。ゲームIPの映像化が業界標準戦略に。</p>'+
      '<p><strong>バンダイナムコ:</strong> ドラゴンボール・ワンピース等IP多角化で2025年3月期DL版ゲーム収益が任天堂を上回る成果。</p>'+
      '<p><strong>カプコン:</strong> 「モンスターハンターワイルズ」発売後即1,000万本超突破。海外売上80%超の高い国際競争力。</p>'+
      '<p><strong>ポケモン:</strong> ゲーム・アニメ・映画・カードゲーム・テーマパーク全方位展開で年間100億ドル超のIP収益を維持。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--navy)">'+
      '<div class="commentary-title">M&A・業界再編</div>'+
      '<p><strong>Microsoft:</strong> Activision Blizzard買収690億ドル（2023年完了）— 業界史上最大M&A。</p>'+
      '<p><strong>ソニー:</strong> Bungie買収36億ドル（2022年）でライブサービス型ゲーム強化。</p>'+
      '<p><strong>ブシロード:</strong> フロントウイング全株式をグッドスマイルカンパニーグループへ譲渡（2024年9月）。</p>'+
      '<p>中小ゲーム企業のM&Aターゲット化が進行。独自IPを持つ企業のバリュエーション上昇が見込まれる。</p>'+
    '</div>';
}

/* ── rGrowth ── */
function rGrowth(){
  var el=g('sec-growth');
  el.innerHTML=
    secH('05','成長ドライバーとリスク','AI活用・プラットフォーム競争・規制リスクの分析')+
    '<div class="commentary" style="border-left-color:var(--green)">'+
      '<div class="commentary-title">成長ドライバー</div>'+
      '<p><strong>(1) 生成AI活用:</strong> NPC会話自動生成、テクスチャ・3Dモデル自動作成、QAテスト自動化で開発コスト20〜30%削減の試算。'+
      'AI活用は脅威ではなく機会として作用する見込み。</p>'+
      '<p><strong>(2) 海外売上比率の拡大:</strong> カプコン80%超、バンダイナムコ50%超、コナミ40%超。'+
      'グローバル展開力がバリュエーションプレミアムの源泉。</p>'+
      '<p><strong>(3) モバイルIAPモデルの高度化:</strong> 「ウマ娘」「モンスト」「ポケポケ」等の大型タイトルが月次数十億円規模の安定収益。'+
      'DeNAは「ポケポケ」効果で2025年4-6月営業利益が前年比1,061%増。</p>'+
      '<p><strong>(4) VTuber・ライブ配信:</strong> ANYCOLORがグロース→プライムへ昇格。カバーはIPO初値430%上昇。'+
      'ゲーム企業とのコラボで新たなIP創出・消費形として成長。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">プラットフォーム競争</div>'+
      '<p><strong>Switch 2:</strong> 2026年3月期1,900万台販売見込み（任天堂上方修正）。AI半導体需要による部品コスト上昇がリスク。</p>'+
      '<p><strong>PS5 + クラウド:</strong> Microsoft、Apple、Googleがクラウドゲーミングに参入。「ハード販売台数」から「エコシステム囲い込み」へ構造転換が進行。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">リスク要因</div>'+
      '<p><strong>(1) ガチャ規制:</strong> 有料ランダム型アイテム販売の規制議論が国内外で継続中。景表法改正の動きに注意。</p>'+
      '<p><strong>(2) 中国市場:</strong> 2023年12月の突然のゲーム規制案でテンセント等の時価総額11兆円超が消失。日本企業の中国依存度は重要監視項目。</p>'+
      '<p><strong>(3) 開発費高騰:</strong> AAA級ゲーム1タイトル100〜300億円（一部大作500億円超）。ヒット確率の低下と開発期間の長期化で中小企業の存続リスクが上昇。</p>'+
      '<p><strong>(4) 円高転換:</strong> 海外売上比率の高い企業は円高局面で収益目減りリスク。</p>'+
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
      '<strong>カテゴリ分類:</strong> 各社の主力事業に基づき5カテゴリに分類（コンソール・PC / モバイルゲーム / アニメ・映像・IP / VTuber・配信 / ゲーム支援）。<br>'+
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
    var snap=await m.getDoc(m.doc(db,'premiumContent','entertainment-companies'));
    if(snap.exists()){companies=snap.data().companies||[];window.companies=companies;}
  }catch(e){console.error('Premium data load failed:',e);return;}
  var countEl=document.getElementById('companyCount');
  if(countEl)countEl.textContent=companies.length;
  initNav();render();
};

document.addEventListener('DOMContentLoaded',function(){initNav();});
})();
