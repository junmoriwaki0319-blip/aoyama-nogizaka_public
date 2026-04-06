/**
 * 広告代理店セクター ダッシュボード
 * 7タブ: Executive Summary / バリュエーション / Tier分析 / M&A・再編 / 成長とリスク / 個別企業 / データソース
 */
(function(){
'use strict';

const TIERS={Tier1:'大手総合',Tier2:'デジタル中堅',Tier3:'グロース・スタンダード'};
const TC={Tier1:'#1a2d4f',Tier2:'#2d7a4f',Tier3:'#9b8b6e'};
const TB={Tier1:'badge-publisher',Tier2:'badge-mobile',Tier3:'badge-anime'};

// companies をグローバルスコープに公開（firebase-auth-ad-agency.js が参照するため）
if(typeof companies==='undefined')window.companies=[];
var companies=window.companies;
var tab='exec',selComp=null;
const CH={};

/* ── helpers ── */
function g(id){return document.getElementById(id);}
function fmtM(v){if(v==null)return'-';if(v>=10000)return(v/10000).toFixed(1)+'兆円';return v.toLocaleString()+'億円';}
function shortName(n){return n.replace(/ホールディングス|HD|グループ/g,'').trim();}
function avg(a,k){var v=a.filter(function(c){return c[k]!=null;});return v.length?(v.reduce(function(s,c){return s+c[k];},0)/v.length).toFixed(1):'-';}
function topN(a,k,n){return[].concat(a).filter(function(c){return c[k]!=null;}).sort(function(x,y){return y[k]-x[k];}).slice(0,n);}
function nv(v,s,f){if(v==null)return'-';if(f==='f1')return v.toFixed(1)+(s||'');if(f==='f2')return v.toFixed(2)+(s||'');if(f==='loc')return v.toLocaleString()+(s||'');return v+(s||'');}
function pn(v,s){if(v==null)return'-';return'<span class="'+(v>=0?'pos':'neg')+'">'+(v>0?'+':'')+v.toFixed(1)+(s||'')+'</span>';}
function mc(id,type,data,opts){var ctx=g(id);if(!ctx)return;var base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};if(['bar','line','scatter','bubble'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};CH[id]=new Chart(ctx,{type:type,data:data,options:dm(base,opts||{})});}
function dc(ids){ids.forEach(function(id){if(CH[id]){CH[id].destroy();delete CH[id];}});}
function dm(t,s){var o=Object.assign({},t);Object.keys(s).forEach(function(k){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]);else o[k]=s[k];});return o;}
function secH(n,t,d){return'<div class="sec-header"><div class="sec-num">SECTION '+n+'</div><div class="sec-title">'+t+'</div><div class="sec-desc">'+d+'</div></div>';}
function kpi(l,v,s,cls){return'<div class="kpi-card '+cls+'"><div class="kpi-label">'+l+'</div><div class="kpi-value">'+v+'</div>'+(s?'<div class="kpi-sub">'+s+'</div>':'')+'</div>';}
function rankCard(t,items,fn){return'<div class="ranking-card"><div class="ranking-title">'+t+'</div>'+items.map(function(c,i){return'<div class="ranking-row"><span><span class="ranking-num">'+(i+1)+'</span>'+shortName(c.name)+'</span><span style="font-weight:600">'+fn(c)+'</span></div>';}).join('')+'</div>';}
function byTier(t){return companies.filter(function(c){return c.tier===t;});}

/* ── rExec ── */
function rExec(){
  var el=g('sec-exec');
  var tm=companies.reduce(function(s,c){return s+(c.marketCap||0);},0);
  var aOP=avg(companies,'operatingMargin'),aROE=avg(companies,'roe'),aPER=avg(companies,'per'),aPBR=avg(companies,'pbr');
  var topMC=topN(companies,'marketCap',3),topOP=topN(companies,'operatingMargin',3),topROE=topN(companies,'roe',3);
  var t1=byTier('Tier1').length,t2=byTier('Tier2').length,t3=byTier('Tier3').length;

  el.innerHTML=
    secH('01','Executive Summary','日本の広告代理店'+companies.length+'社の時価総額・財務指標・業界動向を網羅的に分析')+
    '<div class="commentary">'+
      '<strong>セクター概況 (2026年4月基準):</strong> 対象<strong>'+companies.length+'社</strong>の合計時価総額は<strong>'+fmtM(tm)+'</strong>。'+
      '日本の広告市場2025年7.7兆円、インターネット広告が約48%を占め、従来型マス広告からデジタルへの不可逆的シフトが続く。'+
      'Tier1（大手'+t1+'社）が時価総額の大半を占めるが、Tier3（グロース'+t3+'社）は企業数で過半。'+
      '平均営業利益率<strong>'+aOP+'%</strong>、平均ROE<strong>'+aROE+'%</strong>。<br><br>'+
      '<strong>企業選別の提言 — 3軸で評価:</strong><br>'+
      '<strong>(1) M&A・グループ再編によるバリューアップ余地:</strong> 博報堂DYのデジタルHD完全子会社化、NTTドコモのCARTA買収等、親子上場解消・事業統合が価値創造の起点。<br>'+
      '<strong>(2) AI・テクノロジーを活用したビジネスモデル転換力:</strong> サイバーエージェントのAIクリエイティブ生成、Appier GroupのAI広告最適化等、テクノロジー主導の成長企業に注目。<br>'+
      '<strong>(3) プラットフォーマー依存脱却と独自データ資産の構築:</strong> Google/Meta依存度の高い中小代理店はAI自動化による存在意義希薄化リスク。'+
      '独自データ資産を持つ企業が中期的に優位。<br><br>'+
      '特に<strong>時価総額100億円以下のグロース市場銘柄</strong>にロールアップ型M&Aターゲットとなるリスクおよびチャンスあり。'+
    '</div>'+
    '<div class="kpi-grid">'+
      kpi('対象企業数',companies.length+'社','Tier1:'+t1+' / Tier2:'+t2+' / Tier3:'+t3,'c-navy')+
      kpi('時価総額合計',fmtM(tm),'','c-navy')+
      kpi('平均営業利益率',aOP+'%','','c-gold')+
      kpi('平均ROE',aROE+'%','','c-green')+
      kpi('平均PER',aPER+'倍','','c-navy')+
      kpi('平均PBR',aPBR+'倍','','c-gold')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 億円</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">営業利益率 vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / Tier色分け</div><div class="chart-area tall"><canvas id="exBub"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">Tier別 時価総額構成比</div><div class="chart-area"><canvas id="exPie"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">主要ランキング</div>'+
        '<div style="display:grid;gap:10px;padding-top:8px;">'+
          rankCard('時価総額',topMC,function(c){return fmtM(c.marketCap);})+
          rankCard('営業利益率',topOP,function(c){return nv(c.operatingMargin,'%','f1');})+
          rankCard('ROE',topROE,function(c){return nv(c.roe,'%','f1');})+
        '</div></div>'+
    '</div>';

  dc(['exMC','exBub','exPie']);
  var s15=[].concat(companies).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);}).slice(0,15);
  mc('exMC','bar',{labels:s15.map(function(c){return shortName(c.name);}),datasets:[{data:s15.map(function(c){return c.marketCap;}),backgroundColor:s15.map(function(c){return TC[c.tier]||'#777';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return fmtM(v);}}}});

  var tierMC={};companies.forEach(function(c){tierMC[c.tier]=(tierMC[c.tier]||0)+(c.marketCap||0);});
  mc('exPie','doughnut',{labels:Object.keys(tierMC).map(function(k){return TIERS[k];}),datasets:[{data:Object.values(tierMC),backgroundColor:Object.keys(tierMC).map(function(k){return TC[k];}),borderWidth:0}]},{plugins:{legend:{position:'right'},datalabels:{display:true,color:'#fff',font:{size:10,weight:600},formatter:function(v,ctx){var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);return(v/t*100).toFixed(1)+'%';}}}});

  var mx=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));
  mc('exBub','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.operatingMargin!=null&&c.roe!=null;}).map(function(c){return{x:c.operatingMargin,y:c.roe,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx)*30),name:c.name};}),backgroundColor:TC[tier]+'88',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'営業利益率 (%)'}},y:{title:{display:true,text:'ROE (%)'}}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': OPM'+x.raw.x+'% / ROE'+x.raw.y+'%';}}}}});
}

/* ── rValuation ── */
function rValuation(){
  var el=g('sec-valuation');
  var sorted=[].concat(companies).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);});
  el.innerHTML=
    secH('02','財務指標・バリュエーション','PER/PBR/ROE/営業利益率の横断比較')+
    '<div class="commentary">'+
      '<strong>バリュエーション分析:</strong> セクター平均PERは<strong>'+avg(companies,'per')+'倍</strong>、PBRは<strong>'+avg(companies,'pbr')+'倍</strong>。'+
      '広告セクターはSaaSと比較して低PER・低PBRの成熟型バリュエーション。PBR-ROE間の正の相関が確認され、<strong>ROE改善が株価評価に直結</strong>する構造。'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">PBR vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="vlPBR"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">PER vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="vlPER"></canvas></div></div>'+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">営業利益率分布</div><div class="chart-area"><canvas id="vlHist"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">ROEランキング TOP15</div><div class="chart-area"><canvas id="vlROE"></canvas></div></div>'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">バリュエーション一覧（時価総額順）</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>Tier</th><th>市場</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
      '<tbody>'+sorted.map(function(c){
        return'<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
          '<td><span class="badge '+(TB[c.tier]||'')+'">'+c.tier+'</span></td><td>'+nv(c.market)+'</td>'+
          '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
          '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';

  bindRows(el);dc(['vlPBR','vlPER','vlHist','vlROE']);
  var mx=Math.max.apply(null,companies.map(function(c){return c.marketCap||1;}));
  mc('vlPBR','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.roe!=null&&c.pbr!=null;}).map(function(c){return{x:c.roe,y:c.pbr,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx)*25),name:c.name};}),backgroundColor:TC[tier]+'77',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': ROE'+x.raw.x+'% / PBR'+x.raw.y+'倍';}}}}});
  mc('vlPER','bubble',{datasets:Object.keys(TIERS).map(function(tier){return{label:TIERS[tier],data:byTier(tier).filter(function(c){return c.per!=null&&c.operatingMargin!=null;}).map(function(c){return{x:c.operatingMargin,y:c.per,r:Math.max(4,Math.sqrt((c.marketCap||1)/mx)*25),name:c.name};}),backgroundColor:TC[tier]+'77',borderColor:TC[tier],borderWidth:1};})},{scales:{x:{title:{display:true,text:'営業利益率 (%)'}},y:{title:{display:true,text:'PER (倍)'},max:80}},plugins:{tooltip:{callbacks:{label:function(x){return x.raw.name+': OPM'+x.raw.x+'% / PER'+x.raw.y+'倍';}}}}});
  var bins=[-20,-5,0,5,10,15,20,30];var hist=bins.slice(0,-1).map(function(_,i){return companies.filter(function(c){return c.operatingMargin!=null&&c.operatingMargin>=bins[i]&&c.operatingMargin<bins[i+1];}).length;});
  mc('vlHist','bar',{labels:bins.slice(0,-1).map(function(b,i){return b+'~'+bins[i+1]+'%';}),datasets:[{data:hist,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});
  var roeTop=topN(companies,'roe',15);
  mc('vlROE','bar',{labels:roeTop.map(function(c){return shortName(c.name);}),datasets:[{data:roeTop.map(function(c){return c.roe;}),backgroundColor:roeTop.map(function(c){return TC[c.tier]||'#777';}),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1)+'%';}}}});
}

/* ── rTier ── */
function rTier(){
  var el=g('sec-tier');
  var tierKeys=Object.keys(TIERS);
  var tierStats=tierKeys.map(function(t){
    var cs=byTier(t);
    return{tier:t,label:TIERS[t],count:cs.length,
      totalMcap:cs.reduce(function(s,c){return s+(c.marketCap||0);},0),
      avgOPM:avg(cs,'operatingMargin'),avgROE:avg(cs,'roe'),avgPER:avg(cs,'per'),avgPBR:avg(cs,'pbr')};
  });

  el.innerHTML=
    secH('03','Tier分析','大手総合・デジタル中堅・グロースの3階層比較')+
    '<div class="commentary">'+
      '<strong>Tier構成:</strong> '+
      'Tier1（大手'+tierStats[0].count+'社）が時価総額の'+((tierStats[0].totalMcap/companies.reduce(function(s,c){return s+(c.marketCap||0);},0))*100).toFixed(0)+'%を占める寡占構造。'+
      'Tier3（グロース'+tierStats[2].count+'社）は企業数で過半を占めるが、時価総額100億円未満が多く、'+
      'グロース市場上場基準引き上げ（5年経過後に時価総額100億円以上）により約7割が基準未達リスク。<br>'+
      'ロールアップ型M&A（Macbee Planet、ジーニー等）による業界再編が加速する見込み。'+
    '</div>'+
    '<div class="kpi-grid">'+
      tierStats.map(function(s){return kpi(s.label,s.count+'社','時価総額計 '+fmtM(s.totalMcap),'c-navy');}).join('')+
    '</div>'+
    '<div class="chart-row">'+
      '<div class="chart-panel"><div class="chart-panel-title">Tier別 平均指標比較</div><div class="chart-area tall"><canvas id="trBar"></canvas></div></div>'+
      '<div class="chart-panel"><div class="chart-panel-title">Tier別 時価総額構成</div><div class="chart-area tall"><canvas id="trMcap"></canvas></div></div>'+
    '</div>'+
    tierKeys.map(function(t){
      var cs=byTier(t).sort(function(a,b){return(b.marketCap||0)-(a.marketCap||0);});
      return'<div class="table-panel"><div class="table-header"><div class="table-header-title">'+TIERS[t]+' ('+cs.length+'社)</div></div><div class="table-scroll"><table>'+
        '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>市場</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th></tr></thead>'+
        '<tbody>'+cs.map(function(c){
          return'<tr class="clickable-row" data-code="'+c.ticker+'"><td>'+c.ticker+'</td><td><strong>'+shortName(c.name)+'</strong></td>'+
            '<td>'+nv(c.market)+'</td><td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td>'+
            '<td>'+nv(c.roe,'%','f1')+'</td><td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td></tr>';
        }).join('')+'</tbody></table></div></div>';
    }).join('');

  bindRows(el);dc(['trBar','trMcap']);
  mc('trBar','bar',{labels:tierStats.map(function(s){return s.label;}),datasets:[
    {label:'営業利益率(%)',data:tierStats.map(function(s){return parseFloat(s.avgOPM)||0;}),backgroundColor:'#1a2d4f'},
    {label:'ROE(%)',data:tierStats.map(function(s){return parseFloat(s.avgROE)||0;}),backgroundColor:'#2d7a4f'},
    {label:'PBR(倍)',data:tierStats.map(function(s){return parseFloat(s.avgPBR)||0;}),backgroundColor:'#9b8b6e'}
  ]},{plugins:{datalabels:{display:true,anchor:'end',align:'top',color:'#999',font:{size:9},formatter:function(v){return v.toFixed(1);}}}});
  mc('trMcap','bar',{labels:tierStats.map(function(s){return s.label;}),datasets:[{label:'時価総額合計(億円)',data:tierStats.map(function(s){return s.totalMcap;}),backgroundColor:tierKeys.map(function(k){return TC[k];}),borderWidth:0}]},{plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'top',color:'#999',font:{size:9},formatter:function(v){return fmtM(v);}}}});
}

/* ── rMna ── */
function rMna(){
  var el=g('sec-mna');
  el.innerHTML=
    secH('04','M&A・業界再編','大型M&A・グループ再編・垂直統合の動向分析')+
    '<div class="commentary" style="border-left-color:var(--gold)">'+
      '<div class="commentary-title">M&Aによる業界再編の加速</div>'+
      '<p><strong>電通グループ(4324):</strong> 2025年12月期最終赤字3,276億円。海外事業で約3,000億円の減損損失を計上。約3,400人削減。一方、国内事業の営業利益率は20%台半ばと堅調。</p>'+
      '<p><strong>博報堂DYホールディングス(2433):</strong> デジタルホールディングスを約270億円でTOB、2026年3月に非上場化完了。マス広告×デジタル専業の同業統合。</p>'+
      '<p><strong>NTTドコモ:</strong> CARTA HOLDINGS 約249億円でTOB。通信×広告テクノロジーの水平統合。</p>'+
      '<p><strong>ADKホールディングス:</strong> クラフトン（PUBG開発元）へ約750億円で売却。ゲームIP×広告×アニメの垂直統合。</p>'+
    '</div>'+
    '<div class="commentary">'+
      '<div class="commentary-title">グループ再編 — 取り込みと切り離しの対照的な動き</div>'+
      '<p><strong>博報堂DY:</strong> 親子上場を積極整理。DAC+アイレップ完全子会社化(2018年)、ソウルドアウトTOB約107億円(2022年)、デジタルHD完全子会社化約270億円(2025年)。Hakuhodo DY ONE設立。</p>'+
      '<p><strong>ユナイテッド(2497):</strong> 2025年5月に親子上場関係解消。</p>'+
      '<p><strong>電通グループ未解消:</strong> セプテーニHD(4293)、電通総研(4812)、ドリームインキュベータ(4310)が親子上場のまま。解消なら株価カタリストに。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--navy)">'+
      '<div class="commentary-title">垂直・水平統合 — 広告の枠を超える動き</div>'+
      '<p>電通グループ: ドリームインキュベータ(4310)を持分法適用化（約24%）、BX（ビジネストランスフォーメーション）領域を強化。</p>'+
      '<p>クラフトン×ADK: ゲームIP×広告×アニメの垂直統合モデル。</p>'+
      '<p>NTTドコモ×CARTA: 通信データ×広告テクノロジーの水平統合。9,000万契約のデータ資産を広告最適化に活用。</p>'+
    '</div>';
}

/* ── rGrowth ── */
function rGrowth(){
  var el=g('sec-growth');
  el.innerHTML=
    secH('05','成長ドライバーとリスク','AI活用・デジタルシフト・規制リスクの分析')+
    '<div class="commentary" style="border-left-color:var(--green)">'+
      '<div class="commentary-title">中小型株の注目動向 — ロールアップの新興勢力</div>'+
      '<p><strong>Macbee Planet(7095):</strong> LTVマーケティング軸で成果報酬型市場のリーディングカンパニー。ネットマーケティング完全子会社化「All Ads」改称、MOJA買収。成果報酬型市場は現在約3,000億円→2030年9,000億円成長見込み。</p>'+
      '<p><strong>ジーニー(6562):</strong> SSP/DSP軸に積極的M&A。米国Zelto 67億円子会社化、CATS買収、ソーシャルワイヤー子会社化。グロース市場では異例の大型買収を連発。</p>'+
      '<p><strong>AViC(9554):</strong> 大手ネット専業代理店出身者が2018年設立、2022年上場。売上高26.8億円（前期比+38.6%）。ADK Marketing Solutionsとの合弁設立。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--green)">'+
      '<div class="commentary-title">成長ドライバー</div>'+
      '<p><strong>(1) AI活用:</strong> 博報堂DYのAI自動化、サイバーエージェントのAIクリエイティブ生成、Appier Group(4180)のAI広告最適化プラットフォーム。</p>'+
      '<p><strong>(2) デジタルネイティブ企業:</strong> サイバーエージェントのABEMA黒字化、メディア&IP事業大幅増益。</p>'+
      '<p><strong>(3) 専門特化型:</strong> インフルエンサーマーケティング（トリドリ、THECOO、サイバー・バズ）、リテールメディア（ウネリー）、位置情報広告（ジオロケーションテクノロジー）。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">大手の海外展開は苦戦</div>'+
      '<p><strong>電通:</strong> 海外M&A失敗（イージス買収以降の減損累計5,000億円超）。米州事業で約3,000億円減損。「国内回帰」路線鮮明。</p>'+
      '<p><strong>博報堂DY:</strong> 海外19ヶ国・100超オフィス展開も上期5.5%減収。グローバル市場ではWPP、Publicis、Omnicom等の欧米メガエージェンシーが圧倒的。</p>'+
    '</div>'+
    '<div class="commentary" style="border-left-color:var(--red)">'+
      '<div class="commentary-title">リスク要因</div>'+
      '<p><strong>(1) AI Overviews:</strong> GoogleのAI Overviews表示時オーガニック検索CTR最大58%低下、リスティング広告CTR 78.4%減少（2025年Q3）。LLMO対応が急務。</p>'+
      '<p><strong>(2) プラットフォーマー依存とMCP影響:</strong> Google/Meta依存度の高い中小代理店はAIによる自動化で代理店モデルの価値が薄れるリスク。</p>'+
      '<p><strong>(3) 広告主のインハウス化:</strong> 大手広告主がデジタル広告運用を内製化する動きが加速。</p>'+
      '<p><strong>(4) AI自動化による中小代理店の存在意義希薄化:</strong> 特にリスティング広告の自動入札・自動クリエイティブ生成により、運用代行の付加価値が低下。</p>'+
    '</div>';
}

/* ── rDetail ── */
function rDetail(){
  var el=g('sec-detail');
  if(!selComp)selComp=companies[0];
  var c=selComp;
  var peers=companies.filter(function(x){return x.tier===c.tier&&x.ticker!==c.ticker;});
  el.innerHTML=
    secH('06','個別企業分析','選択企業の詳細指標と同Tier比較')+
    '<div class="inline-filters"><span class="f-label">企業選択</span>'+
      '<select id="dtSel">'+companies.map(function(x){return'<option value="'+x.ticker+'"'+(x.ticker===c.ticker?' selected':'')+'>'+x.ticker+' '+x.name+'</option>';}).join('')+'</select>'+
    '</div>'+
    '<div style="font-family:\'Noto Serif JP\',serif;font-size:1.4rem;font-weight:700;color:var(--navy);margin-bottom:4px;">'+c.name+'</div>'+
    '<div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:12px;">'+c.ticker+' / '+nv(c.market)+' / '+(TIERS[c.tier]||c.tier)+'</div>'+
    '<div class="kpi-grid">'+
      kpi('株価',nv(c.price,'円','loc'),'','c-navy')+
      kpi('時価総額',fmtM(c.marketCap),'','c-navy')+
      kpi('営業利益率',nv(c.operatingMargin,'%','f1'),'','c-gold')+
      kpi('ROE',nv(c.roe,'%','f1'),'','c-green')+
      kpi('PER',nv(c.per,'倍','f1'),'','c-navy')+
      kpi('PBR',nv(c.pbr,'倍','f2'),'','c-gold')+
    '</div>'+
    (peers.length?
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">同Tier比較 — '+(TIERS[c.tier]||c.tier)+'</div></div><div class="table-scroll"><table>'+
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
      '<strong>データソース:</strong> 株価・時価総額・営業利益率・ROE・PER・PBRはYahoo Finance (yahoo-finance2)から取得した実データ。基準日: 2026年4月取得時点の直近終値。<br>'+
      '<strong>Tier分類:</strong> Tier1=大手総合代理店、Tier2=デジタル広告・ネット広告中堅、Tier3=グロース・スタンダード市場の中小型株。<br>'+
      '<strong>更新頻度:</strong> 四半期ごとにyahoo-finance2でデータ更新 → Firestoreへアップロード。<br>'+
      '<strong>注意事項:</strong> 本レポートは情報提供を目的としたものであり、特定の金融商品の売買を推奨するものではありません。投資判断はご自身の責任においてお願いいたします。'+
    '</div>'+
    '<div class="table-panel"><div class="table-header"><div class="table-header-title">全企業データ一覧</div></div><div class="table-scroll"><table>'+
      '<thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>Tier</th><th>市場</th><th>時価総額(億)</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>株価</th></tr></thead>'+
      '<tbody>'+companies.map(function(c){
        return'<tr><td>'+c.ticker+'</td><td>'+c.name+'</td><td><span class="badge '+(TB[c.tier]||'')+'">'+c.tier+'</span></td><td>'+nv(c.market)+'</td>'+
          '<td>'+nv(c.marketCap,'','loc')+'</td><td>'+nv(c.operatingMargin,'%','f1')+'</td><td>'+nv(c.roe,'%','f1')+'</td>'+
          '<td>'+nv(c.per,'倍','f1')+'</td><td>'+nv(c.pbr,'倍','f2')+'</td><td>'+nv(c.price,'円','loc')+'</td></tr>';
      }).join('')+'</tbody></table></div></div>';
}

/* ── navigation & routing ── */
function render(){
  var fns={exec:rExec,valuation:rValuation,tier:rTier,mna:rMna,growth:rGrowth,detail:rDetail,source:rSource};
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
    var snap=await m.getDoc(m.doc(db,'premiumContent','ad-agency-companies'));
    if(snap.exists()){companies=snap.data().companies||[];window.companies=companies;}
  }catch(e){console.error('Premium data load failed:',e);return;}
  var countEl=document.getElementById('companyCount');
  if(countEl)countEl.textContent=companies.length;
  initNav();render();
};

document.addEventListener('DOMContentLoaded',function(){initNav();});
})();
