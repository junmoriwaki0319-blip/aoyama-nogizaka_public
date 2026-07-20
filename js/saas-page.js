// ===== SCRIPT 1: Utility Functions =====
function fmtM(v){if(v>=10000)return(v/10000).toFixed(1)+'兆円';return v.toLocaleString()+'億円';}
function fmtB(v){if(v>=100000)return(v/100000).toFixed(1)+'兆円';if(v>=100)return Math.round(v/100).toLocaleString()+'億円';return v.toLocaleString()+'百万円';}
function shortName(n){return n.replace(/ホールディングス|HD|グループ/g,'').trim();}
function avg(arr,key){const v=arr.filter(c=>c[key]!=null);return v.length?(v.reduce((s,c)=>s+c[key],0)/v.length).toFixed(1):'-';}
function topN(arr,key,n){return[...arr].filter(c=>c[key]!=null).sort((a,b)=>b[key]-a[key]).slice(0,n);}
function nv(v,suf='',fmt){if(v==null)return'-';if(fmt==='loc')return v.toLocaleString()+suf;if(fmt==='f1')return v.toFixed(1)+suf;if(fmt==='f2')return v.toFixed(2)+suf;return v+suf;}
function hasData(arr,key){return arr.some(c=>c[key]!=null);}
const DATA_AS_OF = {
  // stockPrice は loadPremiumData で実データ(updatedAt=データ取得日)から動的設定する
  stockPrice: '各社最新株価',
  financials: '各社直近本決算(kabutan.jp)',
  saasKPI: '各社直近決算説明資料',
};
// ===== SCRIPT 2: Data Definition =====
// SaaS Dashboard - Company Data (loaded from Firestore after authentication)
// データは認証後にFirestoreから動的に取得されます
const SEGMENTS = {
  BizApp: 'ビジネスアプリ', FinTech: 'FinTech', Security: 'セキュリティ',
  CX: 'CX / マーケティング', Vertical: 'バーティカルSaaS', Other: 'その他',
};
const SEG_BADGE = { BizApp:'badge-bizapp', FinTech:'badge-fintech', Security:'badge-security', CX:'badge-cx', Vertical:'badge-vertical', Other:'badge-other' };
const SC = { BizApp:'#1a2d4f', FinTech:'#2d7a4f', Security:'#b53a3a', CX:'#9b8b6e', Vertical:'#5555aa', Other:'#777' };
const QUARTERS = ['FY24Q1','FY24Q2','FY24Q3','FY24Q4','FY25Q1','FY25Q2','FY25Q3','FY25Q4'];
const P = ['#1a2d4f','#9b8b6e','#2d7a4f','#b53a3a','#5a7fa8','#c8946e','#6b8e5e','#8b6b8e','#4a8b8b','#a89b5a','#7a5a3a','#5a6b8e'];
// マーケットデータ: Firestoreから動的取得（認証後）
let TOPIX_RETURN_1Y, TOPIX_MONTHLY, SAAS_INDEX_MONTHLY, INDEX_MONTHS, SAAS_EVENTS, QUARTERS_DATA;
// 企業データはFirestoreから動的取得（認証後にloadPremiumData()で取得）
let companies = [];
// ===== SCRIPT 3: Dashboard Logic =====
(()=>{
  'use strict';
  Chart.defaults.color='#777777';
  Chart.defaults.borderColor='#e0ddd6';
  Chart.defaults.font.family="'Noto Sans JP', system-ui";
  Chart.defaults.font.size=11;
  Chart.register(ChartDataLabels);
  Chart.defaults.plugins.datalabels={display:false};
  let tab='exec', selComp=null, mCodes=[], mSeg='ALL';
  const C={};
  // Firestoreからプレミアムデータを取得してダッシュボード初期化
  window.loadPremiumData = async function() {
    if (!window.firebaseDb) return;
    try {
      const { collection, doc, getDoc } = await import('https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js');
      const db = window.firebaseDb;
      // マーケットデータ取得
      const marketSnap = await getDoc(doc(db, 'premiumContent', 'saas-market'));
      if (marketSnap.exists()) {
        const m = marketSnap.data();
        TOPIX_RETURN_1Y = m.TOPIX_RETURN_1Y;
        TOPIX_MONTHLY = m.TOPIX_MONTHLY;
        SAAS_INDEX_MONTHLY = m.SAAS_INDEX_MONTHLY;
        INDEX_MONTHS = m.INDEX_MONTHS;
        SAAS_EVENTS = m.SAAS_EVENTS;
        if (m.QUARTERS) QUARTERS_DATA = m.QUARTERS;
      }
      // 企業データ取得
      const compSnap = await getDoc(doc(db, 'premiumContent', 'saas-companies'));
      if (compSnap.exists()) {
        const cd = compSnap.data();
        companies = cd.companies || [];
        // 株価基準日ラベルを doc の updatedAt(=データ取得日)から動的設定
        if (cd.updatedAt) {
          const d = new Date(cd.updatedAt);
          if (!isNaN(d)) DATA_AS_OF.stockPrice = `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日時点`;
        }
      }
      const asOfEl = document.getElementById('stockAsOf');
      if (asOfEl) asOfEl.textContent = DATA_AS_OF.stockPrice;
      document.getElementById('companyCount').textContent = companies.length;
      initNav(); render();
    } catch (e) {
      console.error('Premium data load failed:', e);
    }
  };
  // 未認証時はナビだけ初期化
  initNav();
  function initNav(){
    const navInner=document.getElementById('mainNav');
    const btnL=document.getElementById('navScrollLeft');
    const btnR=document.getElementById('navScrollRight');
    // スクロールボタンの表示制御
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
    // タブクリック
    document.querySelectorAll('.nav-item').forEach(n=>n.addEventListener('click',()=>{
      document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
      n.classList.add('active');
      tab=n.dataset.tab;
      document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
      document.getElementById('sec-'+tab).classList.add('active');
      // アクティブタブを表示領域にスクロール
      n.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
      render();
    }));
  }
  function render(){
    const fns={exec:rExec,market:rMarket,financial:rFinancial,dupont:rDupont,invest:rInvest,matrix:rMatrix,unit:rUnit,detail:rDetail,source:rSource};
    if(fns[tab])fns[tab]();
  }
  // ============ HELPERS ============
  function g(id){return document.getElementById(id);}
  function mc(id,type,data,opts={}){
    const ctx=g(id);if(!ctx)return;
    const base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};
    if(['bar','line','scatter','bubble'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};
    C[id]=new Chart(ctx,{type,data,options:dm(base,opts)});
  }
  function dc(ids){ids.forEach(id=>{if(C[id]){C[id].destroy();delete C[id];}});}
  function dm(t,s){const o={...t};for(const k of Object.keys(s)){if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]);else o[k]=s[k];}return o;}
  function secH(n,t,d){return`<div class="sec-header"><div class="sec-num">SECTION ${n}</div><div class="sec-title">${t}</div><div class="sec-desc">${d}</div></div>`;}
  function kpi(l,v,s,cls){return`<div class="kpi-card ${cls}"><div class="kpi-label">${l}</div><div class="kpi-value">${v}</div>${s?`<div class="kpi-sub">${s}</div>`:''}</div>`;}
  function rankCard(t,items,fn){return`<div class="ranking-card"><div class="ranking-title">${t}</div>${items.map((c,i)=>`<div class="ranking-row"><span><span class="ranking-num">${i+1}</span>${shortName(c.name)}</span><span style="font-weight:600">${fn(c)}</span></div>`).join('')}</div>`;}
  function bindRows(el){el.querySelectorAll('.clickable-row').forEach(tr=>tr.addEventListener('click',()=>{selComp=companies.find(c=>c.code===tr.dataset.code);document.querySelector('.nav-item[data-tab="detail"]').click();}));}
  function pn(v,suf=''){return`<span class="${v>=0?'pos':'neg'}">${v>0?'+':''}${typeof v==='number'?v.toFixed(1):v}${suf}</span>`;}
  // ============ TABLE SORT ============
  function enableTableSort(tableEl, dataArr, colDefs, renderRow, onAfter){
    // colDefs: [{key, type:'num'|'str'},...] matching th order
    // renderRow: (item) => '<tr>...</tr>'
    let sortKey=null, sortDir='desc';
    const ths=tableEl.querySelectorAll('thead th');
    ths.forEach((th,i)=>{
      if(i>=colDefs.length||!colDefs[i])return;
      th.addEventListener('click',()=>{
        const def=colDefs[i];
        if(sortKey===def.key){sortDir=sortDir==='desc'?'asc':'desc';}
        else{sortKey=def.key;sortDir='desc';}
        ths.forEach(h=>{h.classList.remove('sort-asc','sort-desc');});
        th.classList.add(sortDir==='asc'?'sort-asc':'sort-desc');
        const sorted=[...dataArr].sort((a,b)=>{
          let va=typeof def.key==='function'?def.key(a):a[def.key];
          let vb=typeof def.key==='function'?def.key(b):b[def.key];
          if(va==null)va=def.type==='num'?-Infinity:'';
          if(vb==null)vb=def.type==='num'?-Infinity:'';
          if(def.type==='num')return sortDir==='desc'?vb-va:va-vb;
          return sortDir==='desc'?String(vb).localeCompare(String(va)):String(va).localeCompare(String(vb));
        });
        tableEl.querySelector('tbody').innerHTML=sorted.map(renderRow).join('');
        if(onAfter)onAfter(tableEl);
      });
    });
  }
  // ============ 01 EXECUTIVE SUMMARY ============
  function rExec(){
    const el=g('sec-exec');
    const tm=companies.reduce((s,c)=>s+c.marketCap,0);
    const tr=companies.reduce((s,c)=>s+c.revenue,0);
    const aOP=avg(companies,'opMargin'), aROE=avg(companies,'roe');
    const aGrowth=avg(companies,'arrGrowth'), aR40=avg(companies,'ruleOf40');
    const topMC=topN(companies,'marketCap',3), topGrowth=topN(companies,'arrGrowth',3), topR40=topN(companies,'ruleOf40',3);
    const above=companies.filter(c=>c.relativeReturn>0).length;
    const hasARR=hasData(companies,'arr'), hasR40=hasData(companies,'ruleOf40');
    // SaaSインデックス・TOPIX値を動的に算出
    const saasIdx=SAAS_INDEX_MONTHLY&&SAAS_INDEX_MONTHLY.length?SAAS_INDEX_MONTHLY[SAAS_INDEX_MONTHLY.length-1].toFixed(1):'—';
    const topixIdx=TOPIX_MONTHLY&&TOPIX_MONTHLY.length?TOPIX_MONTHLY[TOPIX_MONTHLY.length-1].toFixed(1):'—';
    const idxDiff=SAAS_INDEX_MONTHLY&&TOPIX_MONTHLY?(SAAS_INDEX_MONTHLY[SAAS_INDEX_MONTHLY.length-1]-TOPIX_MONTHLY[TOPIX_MONTHLY.length-1]).toFixed(1):null;
    const idxDiffStr=idxDiff?(Number(idxDiff)<0?'▲'+Math.abs(Number(idxDiff)).toFixed(1):('+'+idxDiff)):'—';
    el.innerHTML=`
      ${secH('01','Executive Summary','国内SaaS セクター全体概況と主要指標ハイライト')}
      <div class="commentary">
        <strong>セクター概況 (${DATA_AS_OF.stockPrice}基準):</strong> 対象${companies.length}社の合計時価総額は<strong>${fmtM(tm)}</strong>、売上高合計<strong>${fmtB(tr)}</strong>。
        SaaSインデックスは直近15ヶ月(2025年1月起点)で<strong>${saasIdx}pt</strong>(TOPIX:${topixIdx}pt)と<strong>${idxDiffStr}pt</strong>のアンダーパフォーム。
        TOPIX超過は${above}/${companies.length}社。平均営業利益率${aOP}%、平均ROE${aROE}%。
        <strong>株価の低迷にもかかわらず、業績ファンダメンタルズは堅調</strong>であり、AI代替懸念による過度なディスカウントが示唆される。<br><br>
        <strong>足元のリスク要因:</strong>
        (1) <strong>DeepSeekショック(2025年1月)</strong>を契機とした生成AIによるSaaS代替懸念。SaaS指数は一時▲15.5%の急落を記録したが、
        その後各社の売上高・営業利益は堅調に推移しており、株価下落は<strong>「バリュエーション調整」であり「ファンダメンタルズ毀損」ではなかった</strong>ことが確認された。
        (2) <strong>AI Agent競争の激化</strong>(2025年6月、OpenAI・Google等のAgent製品相次ぎ発表)により、ワークフロー自動化系SaaSの中長期的な代替リスクが意識される局面。
        ただし、業務特化型SaaS(バーティカル)やミッションクリティカルなBizAppは<strong>スイッチングコストの壁</strong>が依然高い。
        (3) <strong>Anthropicショック(2026年2月)</strong>のClaude Code発表でコーディング支援・開発ツール系SaaSに一時的な売り圧力。
        <strong>株価は▲9.1%急落(117.9→107.2pt)、しかし業績は+8.3%成長</strong> ― この対比がSaaSセクターの本質を物語る。
        FY25下半期で<strong>全${companies.length}社が増収</strong>、営業利益率も21.5%→22.4%へ改善。
        AI関連の売り圧力が最も強かったCX領域でもAppier(OPM+0.6pt)、プレイド(OPM+2.7pt)と利益率が拡大しており、
        国内SaaS企業の多くはAI機能を<strong>自社プロダクトに統合する側</strong>であり、脅威よりも付加価値向上の機会として活用している。<br><br>
        <strong>マネジメントへの5つの提言 ― 株価低迷を打破するアクションプラン:</strong><br>
        <strong>(1) AIカニバリゼーションへの先制的ポジショニング:</strong>
        Anthropicショック後、ラクス▲13.5%、Sansan▲12.5%、freee▲9.0%の急落が示す通り、投資家は「SaaSのUI/UX層がAIエージェントに置換される」リスクを織り込んでいる。
        しかし業界固有のワークフロー・規制対応・データネットワーク効果を持つSaaSは代替困難。自社プロダクトの「AI代替可能層」と「代替困難層」を分解し、後者への投資集中とAI統合戦略をIRで明示することがバリュエーション回復の起点となる。<br>
        <strong>(2) 「Rule of 40超」を目指す利益構造転換:</strong>
        freee・マネーフォワードの黒字化達成は「成長一辺倒」から「成長+収益性」への転換を象徴。Rule of 40を社内KPIとして採用し、成長率と利益率のバランスを四半期ごとに取締役会レベルで議論すべき。Gross Margin80%超の企業は販管費効率化だけで大幅改善が可能。<br>
        <strong>(3) ARR・アカウント数・成長率の開示徹底による不透明性ディスカウント解消:</strong>
        国内SaaS30社のうちARR開示は${companies.filter(c=>c.arr!=null).length}社に留まる。ARR・有料顧客数・ARR成長率は国際的にSaaS企業評価の標準指標であり、これらの四半期開示が機関投資家カバレッジ取得と適正バリュエーション獲得への最短経路。NRR・チャーンレートは補足指標として有用だが、プロダクト特性や季節要因による変動が大きく、単体での開示は「SaaSの死」シナリオの補強材料に利用されるリスクもあるため、開示戦略には慎重な設計が求められる。<br>
        <strong>(4) キャッシュの戦略的配分 ― M&Aと成長投資の好機:</strong>
        黒字化とバリュエーション圧縮が同時に進行する現在、潤沢な手元資金の投資先が問われている。金利上昇局面では資金調達コストが漸増するため、<strong>今のうちにキャッシュを成長に振り向ける</strong>経営判断の巧拙が中期的な企業価値を左右する。
        マネーフォワードの事例が示唆的 ― (a) スマートキャンプ(SaaS比較サイト「BOXIL」)を傘下で売上10億→40億円へ成長させた後、丸の内キャピタル(PEファンド)へ売却し特別利益63億円を計上。非中核事業のバリューアップ→売却というPE的手法でポートフォリオを最適化。(b) 一方で2026年2月にソニービズネットワークスからクラウド勤怠管理「AKASHI」を吸収分割で取得。中堅・エンタープライズ向け実績を持つバーティカルSaaSを割安に取得し、クラウド勤怠Plusとしてプロダクト統合を進める。サイバーセキュリティクラウドのDataSign連結(売上+33%)やSansanの非中核事業(EventHub)売却→Bill One集中も同様の「選択と集中」型M&A。<br>
        <strong>(5) IR戦略のリデザイン:</strong>
        決算説明資料へのAI戦略セクション新設、SaaS KPIダッシュボード充実(ARR・顧客数・成長率の標準開示)、Rule of 40のセクター内ベンチマーク提示、機関投資家向け「AI×SaaS」テーマIR Day開催を推奨。
      </div>
      <div class="kpi-grid">
        ${kpi('対象企業数',companies.length+'社','','c-navy')}
        ${kpi('時価総額合計',fmtM(tm),'','c-navy')}
        ${kpi('売上高合計',fmtB(tr),'','c-gold')}
        ${kpi('平均営業利益率',aOP+'%','','c-gold')}
        ${kpi('平均ROE',aROE+'%','','c-navy')}
        ${kpi('平均Gross Margin',avg(companies,'grossMargin')+'%','','c-green')}
        ${kpi('TOPIX超過銘柄',above+'/'+companies.length+'社','','c-navy')}
        ${kpi('vs TOPIX (1Y)',idxDiffStr+'pt','SaaSインデックス','c-red')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">SaaS Index vs TOPIX 推移</div><div class="chart-panel-sub">月次指数 (25/01=100) / TOPIX実績値・SaaS指数は30社加重平均推計</div><div class="chart-area tall"><canvas id="exIdx"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 億円</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">売上高 vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額 / セグメント色分け</div><div class="chart-area tall"><canvas id="exBub"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">セグメント別 時価総額構成比</div><div class="chart-area"><canvas id="exPie"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">主要ランキング</div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;padding-top:8px;">
            ${rankCard('時価総額',topMC,c=>c.marketCap.toLocaleString()+'億円')}
            ${rankCard('営業利益率',topN(companies,'opMargin',3),c=>c.opMargin.toFixed(1)+'%')}
            ${rankCard('ROE',topN(companies,'roe',3),c=>c.roe.toFixed(1)+'%')}
          </div>
        </div>
      </div>`;
    dc(['exIdx','exMC','exPie','exBub']);
    // SaaS Index vs TOPIX with event annotations
    const evAnnot={};
    SAAS_EVENTS.forEach((ev,i)=>{evAnnot['ev'+i]={type:'line',xMin:ev.idx,xMax:ev.idx,borderColor:'rgba(181,58,58,0.5)',borderWidth:1.5,borderDash:[4,3],label:{display:true,content:ev.label,position:'start',backgroundColor:'rgba(181,58,58,0.85)',color:'#fff',font:{size:9,weight:600},padding:{x:4,y:2}}};});
    mc('exIdx','line',{labels:INDEX_MONTHS,datasets:[{label:'SaaS Index',data:SAAS_INDEX_MONTHLY,borderColor:'#1a2d4f',backgroundColor:'rgba(26,45,79,0.08)',fill:true,tension:0.3,pointRadius:3,borderWidth:2},{label:'TOPIX',data:TOPIX_MONTHLY,borderColor:'#9b8b6e',borderDash:[5,3],fill:false,tension:0.3,pointRadius:2,borderWidth:2}]},{plugins:{annotation:{annotations:evAnnot}}});
    const s15=[...companies].sort((a,b)=>b.marketCap-a.marketCap).slice(0,15);
    mc('exMC','bar',{labels:s15.map(c=>shortName(c.name)),datasets:[{data:s15.map(c=>c.marketCap),backgroundColor:s15.map(c=>SC[c.segment]||'#777'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:v=>fmtM(v)}}});
    const segMC={};companies.forEach(c=>{segMC[c.segment]=(segMC[c.segment]||0)+c.marketCap;});
    mc('exPie','doughnut',{labels:Object.keys(segMC).map(k=>SEGMENTS[k]),datasets:[{data:Object.values(segMC),backgroundColor:Object.keys(segMC).map(k=>SC[k]),borderWidth:0}]},{plugins:{legend:{position:'right'},datalabels:{display:true,color:'#fff',font:{size:10,weight:600},formatter:(v,ctx)=>{const t=ctx.dataset.data.reduce((a,b)=>a+b,0);return(v/t*100).toFixed(1)+'%';}}}});
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('exBub','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.revenue/1000,y:c.opMargin,r:Math.max(4,Math.sqrt(c.marketCap/mx)*30),name:c.name})),backgroundColor:SC[seg]+'88',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'売上高 (十億円)'}},y:{title:{display:true,text:'営業利益率 (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 売上${(x.raw.x*1000).toLocaleString()}百万 / OPM${x.raw.y}%`}},annotation:{annotations:{opm15:{type:'line',yMin:15,yMax:15,borderColor:'rgba(155,139,110,0.4)',borderWidth:1.5,borderDash:[4,3],label:{display:true,content:'OPM 15%',position:'end',backgroundColor:'rgba(155,139,110,0.6)',color:'#fff',font:{size:9}}}}}}});
  }
  // ============ 02 STOCK MARKET PERFORMANCE ============
  function rMarket(){
    const el=g('sec-market');
    const sorted=[...companies].sort((a,b)=>b.relativeReturn-a.relativeReturn);
    const above=companies.filter(c=>c.relativeReturn>0).length;
    const mkRow=c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td><span class="badge ${SEG_BADGE[c.segment]||''}">${SEGMENTS[c.segment]}</span></td><td>${c.stockPrice.toLocaleString()}</td><td>${c.marketCap.toLocaleString()}</td><td class="${c.stockReturn1Y>=0?'pos':'neg'}">${c.stockReturn1Y>0?'+':''}${c.stockReturn1Y.toFixed(1)}%</td><td class="${c.relativeReturn>=0?'pos':'neg'}" style="font-weight:600">${c.relativeReturn>0?'+':''}${c.relativeReturn.toFixed(1)}%</td><td>${c.per.toFixed(1)}</td><td>${c.pbr.toFixed(1)}</td><td>${c.psr.toFixed(1)}</td></tr>`;
    const saasIdx2=SAAS_INDEX_MONTHLY&&SAAS_INDEX_MONTHLY.length?SAAS_INDEX_MONTHLY[SAAS_INDEX_MONTHLY.length-1].toFixed(1):'—';
    const topixIdx2=TOPIX_MONTHLY&&TOPIX_MONTHLY.length?TOPIX_MONTHLY[TOPIX_MONTHLY.length-1].toFixed(1):'—';
    const idxDiff2=SAAS_INDEX_MONTHLY&&TOPIX_MONTHLY?(SAAS_INDEX_MONTHLY[SAAS_INDEX_MONTHLY.length-1]-TOPIX_MONTHLY[TOPIX_MONTHLY.length-1]).toFixed(1):null;
    const idxDiffStr2=idxDiff2?(Number(idxDiff2)<0?'▲'+Math.abs(Number(idxDiff2)).toFixed(1):('+'+idxDiff2)):'—';
    const top3=sorted.slice(0,3).map(c=>shortName(c.name)+'('+pn(c.stockReturn1Y,'%')+')').join('、');
    const bot3=sorted.slice(-3).reverse().map(c=>shortName(c.name)+'('+pn(c.stockReturn1Y,'%')+')').join('、');
    el.innerHTML=`
      ${secH('02','株式市場パフォーマンス','TOPIX対比の相対株価推移と銘柄別騰落率')}
      <div class="commentary">
        <strong>市場動向 (2025年1月〜2026年3月):</strong> SaaSインデックスは<strong>${saasIdx2}pt</strong>(TOPIX:${topixIdx2}pt)で<strong>${idxDiffStr2}pt</strong>のアンダーパフォーム。
        TOPIX超過は${above}/${companies.length}社。${top3}が牽引。
        一方、${bot3}等は大幅下落。<strong>業績堅調にもかかわらず株価は低迷</strong>しており、AI代替懸念によるディスカウントが続く。<br><br>
        <strong>主要イベントと株価への影響:</strong>
        2025年1月末の<strong>DeepSeekショック</strong>でSaaS指数は84.5pt(▲15.5%)まで急落。「生成AIがSaaSを代替する」との懸念が急速に広まった。
        各社の四半期決算は売上高成長を継続していたが、<strong>回復は緩やかで86.7pt(25/03)に留まった</strong>。
        6月の<strong>AI Agent競争激化</strong>(OpenAI・Google Agent)で再び足踏み。その後は業績堅調を背景に徐々に回復し、
        2026年1月には<strong>117.9pt</strong>まで戻したが、2月の<strong>Anthropicショック</strong>(Claude Code発表)で再び<strong>107.2ptへ急落(▲9.1%)</strong>。
        結果として「AIショック=SaaS企業の業績毀損」ではなく、
        <strong>「AI懸念→バリュエーション圧縮→業績堅調で下値限定」</strong>のパターンが繰り返されている。<br><br>
        <strong>Anthropicショック後の業績検証 ― 株価▲9.1% vs 業績+8.3%:</strong>
        Claude Code発表(2026年2月)でSaaS指数は117.9→107.2pt(<strong>▲9.1%</strong>)へ急落。
        しかしFY25下半期の実績は対照的で、全${companies.length}社が増収、セクター売上高は<strong>+8.3%成長</strong>(280,445→303,600百万円)、営業利益率も21.5%→<strong>22.4%</strong>に改善。
        AI関連の売り圧力が最も強かったCX領域でもAppier(OPM+0.6pt)、プレイド(OPM+2.7pt)と利益率拡大。
        <strong>「株価の調整幅」と「業績の成長幅」の乖離は、市場が生成AIのSaaS代替リスクを過大評価している</strong>ことを示唆する。
      </div>
      <div class="kpi-grid">
        ${kpi('SaaSインデックス',saasIdx2+'pt','直近15ヶ月','c-navy')}
        ${kpi('TOPIX',topixIdx2+'pt','同期間','c-navy')}
        ${kpi('vs TOPIX',idxDiffStr2+'pt','アンダーパフォーム','c-red')}
        ${kpi('TOPIX超過銘柄',above+'/'+companies.length+'社','','c-gold')}
        ${kpi('Anthropic後 株価','▲9.1%','SaaS指数 26/02','c-red')}
        ${kpi('Anthropic後 業績','+8.3%','H2売上成長','c-green')}
        ${kpi('H2営業利益率','22.4%','H1: 21.5%','c-gold')}
        ${kpi('FY25下期 増収企業',companies.length+'/'+companies.length+'社','全社増収','c-green')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">SaaS Index vs TOPIX 推移</div><div class="chart-panel-sub">月次指数 (25/01=100) / TOPIX実績値・SaaS指数は30社加重平均推計</div><div class="chart-area tall"><canvas id="mkIdx"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">銘柄別 株価騰落率 (1年)</div><div class="chart-panel-sub">赤線=TOPIX (${TOPIX_RETURN_1Y}%)</div><div class="chart-area tall"><canvas id="mkStocks"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">銘柄別パフォーマンス一覧</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">株価: ${DATA_AS_OF.stockPrice} / PER・PBR・ROE: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table id="tblMarket">
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>セグメント</th><th>株価</th><th>時価総額(億)</th><th>1Y騰落率</th><th>vs TOPIX</th><th>PER</th><th>PBR</th><th>PSR</th></tr></thead>
        <tbody>${sorted.map(c=>mkRow(c)).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['mkIdx','mkStocks']);
    // テーブルソート有効化
    const mkColDefs=[
      {key:'code',type:'str'},{key:'name',type:'str'},{key:'segment',type:'str'},
      {key:'stockPrice',type:'num'},{key:'marketCap',type:'num'},
      {key:'stockReturn1Y',type:'num'},{key:'relativeReturn',type:'num'},
      {key:'per',type:'num'},{key:'pbr',type:'num'},{key:'psr',type:'num'}
    ];
    enableTableSort(g('tblMarket'),companies,mkColDefs,mkRow,t=>bindRows(el));
    // 初期ソート状態のインジケータ（vs TOPIX降順）
    g('tblMarket').querySelectorAll('thead th')[6].classList.add('sort-desc');
    const mkAnnot={};
    SAAS_EVENTS.forEach((ev,i)=>{mkAnnot['ev'+i]={type:'line',xMin:ev.idx,xMax:ev.idx,borderColor:'rgba(181,58,58,0.5)',borderWidth:1.5,borderDash:[4,3],label:{display:true,content:ev.label,position:'start',backgroundColor:'rgba(181,58,58,0.85)',color:'#fff',font:{size:9,weight:600},padding:{x:4,y:2}}};});
    mc('mkIdx','line',{labels:INDEX_MONTHS,datasets:[{label:'SaaS Index',data:SAAS_INDEX_MONTHLY,borderColor:'#1a2d4f',backgroundColor:'rgba(26,45,79,0.08)',fill:true,tension:0.3,pointRadius:4,borderWidth:2},{label:'TOPIX',data:TOPIX_MONTHLY,borderColor:'#9b8b6e',borderDash:[5,3],fill:false,tension:0.3,pointRadius:3,borderWidth:2}]},{plugins:{annotation:{annotations:mkAnnot}}});
    mc('mkStocks','bar',{labels:sorted.map(c=>shortName(c.name)),datasets:[{data:sorted.map(c=>c.stockReturn1Y),backgroundColor:sorted.map(c=>c.stockReturn1Y>=TOPIX_RETURN_1Y?'#1a2d4f':'rgba(181,58,58,0.6)'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},annotation:{annotations:{topix:{type:'line',xMin:TOPIX_RETURN_1Y,xMax:TOPIX_RETURN_1Y,borderColor:'#b53a3a',borderWidth:2,borderDash:[4,4],label:{display:true,content:'TOPIX '+TOPIX_RETURN_1Y+'%',position:'start'}}}}}});
  }
  // ============ 03 FINANCIAL ============
  function rFinancial(){
    const el=g('sec-financial');
    const sorted=[...companies].sort((a,b)=>b.revenue-a.revenue);
    el.innerHTML=`
      ${secH('03','財務指標・バリュエーション','売上高・利益率・PER/PBR/PSRの横断比較')}
      <div class="commentary">
        <strong>バリュエーション分析 (${DATA_AS_OF.stockPrice}基準):</strong> セクター平均PERは${avg(companies,'per')}倍、PBRは${avg(companies,'pbr')}倍、PSRは${avg(companies,'psr')}倍。
        SaaS企業は高PSR・高PBRのグロース評価が主流。PBR-ROE間の正の相関が確認され、<strong>ROE改善が株価評価に直結</strong>する構造。
        赤字企業はPSR(売上高倍率)での評価が適切。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">売上高 vs 営業利益</div><div class="chart-panel-sub">TOP15 / 百万円</div><div class="chart-area tall"><canvas id="fnRP"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">PSR vs 売上高</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 右上=大型高評価 左下=小型割安</div><div class="chart-area tall"><canvas id="fnPSR"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">純利益率 vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 右上=高収益×高資本効率の優良企業</div><div class="chart-area"><canvas id="fnPBR"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">営業利益率分布</div><div class="chart-area"><canvas id="fnHist"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">財務指標一覧 (売上高順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">財務データ: ${DATA_AS_OF.financials} / 株価: ${DATA_AS_OF.stockPrice}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>売上高(百万)</th><th>営業利益(百万)</th><th>営業利益率</th><th>純利益率</th><th>PER</th><th>PBR</th><th>PSR</th><th>ROE</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td>${c.revenue.toLocaleString()}</td><td>${c.opProfit.toLocaleString()}</td><td class="${c.opMargin>=15?'pos':c.opMargin<0?'neg':''}">${c.opMargin.toFixed(1)}%</td><td>${c.netMargin.toFixed(1)}%</td><td>${c.per.toFixed(1)}</td><td>${c.pbr.toFixed(1)}</td><td>${c.psr.toFixed(1)}</td><td class="${c.roe>=15?'pos':''}">${c.roe.toFixed(1)}%</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['fnRP','fnPSR','fnPBR','fnHist']);
    const t15=sorted.slice(0,15);
    mc('fnRP','bar',{labels:t15.map(c=>shortName(c.name)),datasets:[{label:'売上高',data:t15.map(c=>c.revenue),backgroundColor:'rgba(26,45,79,0.6)',borderWidth:0},{label:'営業利益',data:t15.map(c=>c.opProfit),backgroundColor:'rgba(45,122,79,0.6)',borderWidth:0}]},{indexAxis:'y'});
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('fnPSR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.revenue/1000,y:c.psr,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'売上高 (十億円)'}},y:{title:{display:true,text:'PSR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 売上${(x.raw.x*1000).toLocaleString()}百万 / PSR${x.raw.y}倍`}}}});
    mc('fnPBR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.roe,y:c.pbr,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ROE${x.raw.x}% / PBR${x.raw.y}倍`}}}});
    const bins=[-10,0,5,10,15,20,30,55];const hist=bins.slice(0,-1).map((_,i)=>companies.filter(c=>c.opMargin>=bins[i]&&c.opMargin<bins[i+1]).length);
    mc('fnHist','bar',{labels:bins.slice(0,-1).map((b,i)=>`${b}~${bins[i+1]}%`),datasets:[{data:hist,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});
  }
  // ============ 04 DUPONT ============
  function rDupont(){
    const el=g('sec-dupont');
    const sorted=[...companies].sort((a,b)=>b.roe-a.roe);
    const hasDupont=hasData(companies,'assetTurnover');
    el.innerHTML=`
      ${secH('04','DuPont分解分析','ROE = 売上高純利益率 x 総資産回転率 x 財務レバレッジ')}
      <div class="commentary">
        <strong>DuPont Analysis:</strong> ROE平均<strong>${avg(companies,'roe')}%</strong>、純利益率平均<strong>${avg(companies,'netMargin')}%</strong>。
        ${hasDupont?`回転率平均<strong>${avg(companies,'assetTurnover')}x</strong>、レバレッジ平均<strong>${avg(companies,'leverage')}x</strong>。`:'総資産回転率・財務レバレッジは現在未取得です。'}
        SaaS企業は低レバレッジ・高利益率型が多く、ROEドライバーは主に純利益率。OBC(ROE21.2%)は高利益率型、マークラインズ(ROE28.0%)は高回転率型の好例。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">ROE vs 純利益率</div><div class="chart-panel-sub">ROE順</div><div class="chart-area tall"><canvas id="dpBar"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">純利益率 vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額</div><div class="chart-area tall"><canvas id="dpScatter"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">収益性指標一覧 (ROE順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">財務データ: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>ROE</th><th>純利益率</th><th>営業利益率</th><th>Gross Margin</th><th>PBR</th><th>PER</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td class="${c.roe>=15?'pos':''}" style="font-weight:700">${c.roe.toFixed(1)}%</td><td>${c.netMargin.toFixed(1)}%</td><td>${c.opMargin.toFixed(1)}%</td><td>${c.grossMargin.toFixed(1)}%</td><td>${c.pbr.toFixed(1)}</td><td>${c.per.toFixed(1)}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['dpBar','dpScatter']);
    mc('dpBar','bar',{labels:sorted.map(c=>shortName(c.name)),datasets:[{label:'純利益率(%)',data:sorted.map(c=>c.netMargin),backgroundColor:'#1a2d4f'},{label:'営業利益率(%)',data:sorted.map(c=>c.opMargin),backgroundColor:'#9b8b6e'}]},{indexAxis:'y',scales:{x:{stacked:false}},plugins:{legend:{position:'top'}}});
    const maxMC=Math.max(...companies.map(c=>c.marketCap));
    mc('dpScatter','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.roe>0).map(c=>({x:c.netMargin,y:c.roe,r:Math.max(4,Math.sqrt(c.marketCap/maxMC)*30),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'売上高純利益率 (%)'}},y:{title:{display:true,text:'ROE (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 純利益率${x.raw.x}% / ROE${x.raw.y}%`}},annotation:{annotations:{roe15:{type:'line',yMin:15,yMax:15,borderColor:'rgba(45,122,79,0.4)',borderWidth:1.5,borderDash:[4,3],label:{display:true,content:'ROE 15%',position:'end',backgroundColor:'rgba(45,122,79,0.7)',color:'#fff',font:{size:9}}}}}}});
  }
  // ============ 05 INVESTMENT METRICS ============
  function rInvest(){
    const el=g('sec-invest');
    const sorted=[...companies].sort((a,b)=>b.psr-a.psr);
    el.innerHTML=`
      ${secH('05','投資指標・バリュエーション','PSR・PER・PBR・ROE等の投資判断指標')}
      <div class="commentary">
        <strong>バリュエーション分析 (${DATA_AS_OF.stockPrice}基準):</strong> SaaS企業は配当・株主優待が少ないため、<strong>PSR・PBR・PER</strong>が主要なバリュエーション指標。
        平均PSRは${avg(companies,'psr')}倍、平均PERは${avg(companies,'per')}倍、平均PBRは${avg(companies,'pbr')}倍。<br>
        <em>※ EV/ARR・FCF利回り・ROIC・キャッシュポジション等の詳細投資指標は現在データ未取得のため非表示です。各社IR開示データの拡充に合わせて順次追加予定。</em>
      </div>
      <div class="kpi-grid">
        ${kpi('平均PSR',avg(companies,'psr')+'倍','','c-navy')}
        ${kpi('平均PER',avg(companies,'per')+'倍','','c-gold')}
        ${kpi('平均PBR',avg(companies,'pbr')+'倍','','c-green')}
        ${kpi('平均ROE',avg(companies,'roe')+'%','','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">PSR vs 売上高成長率</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 高成長×高PSR=市場の期待が大きい銘柄</div><div class="chart-area tall"><canvas id="ivPSR"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">PER vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 右上=高収益で高評価 左上=割高リスク</div><div class="chart-area tall"><canvas id="ivPER"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">バリュエーション一覧 (PSR順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">株価: ${DATA_AS_OF.stockPrice} / 財務データ: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>時価総額(億)</th><th>売上高(百万)</th><th>PSR</th><th>PER</th><th>PBR</th><th>ROE</th><th>営業利益率</th><th>vs TOPIX</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td>${c.marketCap.toLocaleString()}</td><td>${c.revenue.toLocaleString()}</td><td style="font-weight:700">${c.psr.toFixed(1)}</td><td>${c.per.toFixed(1)}</td><td>${c.pbr.toFixed(1)}</td><td class="${c.roe>=15?'pos':''}">${c.roe.toFixed(1)}%</td><td class="${c.opMargin>=15?'pos':c.opMargin<0?'neg':''}">${c.opMargin.toFixed(1)}%</td><td class="${c.relativeReturn>=0?'pos':'neg'}">${c.relativeReturn>0?'+':''}${c.relativeReturn.toFixed(1)}%</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['ivPSR','ivPER']);
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('ivPSR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.stockReturn1Y!=null).map(c=>({x:c.stockReturn1Y,y:c.psr,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'株価騰落率 1Y (%)'}},y:{title:{display:true,text:'PSR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 騰落率${x.raw.x}% / PSR${x.raw.y}倍`}}}});
    mc('ivPER','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.per>0&&c.per<200).map(c=>({x:c.opMargin,y:c.per,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'営業利益率 (%)'}},y:{title:{display:true,text:'PER (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: OPM${x.raw.x}% / PER${x.raw.y}倍`}}}});
  }
  // ============ 06 SaaS KPI MATRIX ============
  function rMatrix(){
    const el=g('sec-matrix');
    const sorted=[...companies].sort((a,b)=>b.revenue-a.revenue);
    const hasR40m=hasData(companies,'ruleOf40');
    const hasChurnm=hasData(companies,'grossChurn');
    const r40Top3=hasR40m?topN(companies,'ruleOf40',3):[];
    const churnTop3=hasChurnm?[...companies].filter(c=>c.grossChurn!=null).sort((a,b)=>a.grossChurn-b.grossChurn).slice(0,3):[];
    el.innerHTML=`
      ${secH('06','SaaS KPIマトリクス','ARR・成長率・Rule of 40の横断比較')}
      <div class="commentary">
        <strong>KPIマトリクス:</strong> 各社決算説明資料からSaaS固有KPIを取得。ARR開示は<strong>${companies.filter(c=>c.arr!=null).length}/${companies.length}社</strong>。
        ${hasData(companies,'arr')?`ARR開示企業の合計は<strong>${fmtB(companies.filter(c=>c.arr!=null).reduce((s,c)=>s+c.arr,0))}</strong>。`:''}
        ARR・有料アカウント数・ARR成長率はSaaS企業の健全性を測る標準指標であり、ほぼすべての上場SaaS企業がIRで開示している。
        非開示企業は売上高・営業利益率等の公開財務データで代替表示。<br><br>
        <strong>Rule of 40:</strong> ARR成長率+営業利益率で算出。${hasR40m?r40Top3.map(c=>shortName(c.name)+'('+c.ruleOf40.toFixed(1)+')').join('、')+'が40超え':'データ開示企業のみ表示'}。<br>
        <em>※ Churn Rate・NRRは補足指標として各社開示値を収録(Churn開示${companies.filter(c=>c.grossChurn!=null).length}社、NRR開示${companies.filter(c=>c.nrr!=null).length}社)。ただしプロダクト特性・季節要因・算出基準の差異が大きいため、企業間の単純比較には注意が必要。</em>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">売上高 vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 右上=規模と収益性を両立する企業</div><div class="chart-area tall"><canvas id="mtRev"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">Gross Margin vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=時価総額 / 対角線から離れるほど販管費が重い</div><div class="chart-area tall"><canvas id="mtGM"></canvas></div></div>
      </div>
      <div class="chart-row single">
        <div class="chart-panel"><div class="chart-panel-title">四半期売上高推移 (TOP10)</div><div class="chart-panel-sub">※四半期データは通期実績からの推計値です</div><div class="chart-area tall"><canvas id="mtQRev"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">SaaS KPI一覧 (売上高順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">SaaS KPI: ${DATA_AS_OF.saasKPI} / 財務データ: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>売上高(百万)</th><th>ARR(百万)</th><th>ARR成長率</th><th>Rule of 40</th><th>営業利益率</th><th>Gross Margin</th><th style="color:#999">Churn※</th><th style="color:#999">NRR※</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td>${c.revenue.toLocaleString()}</td><td>${nv(c.arr,'','loc')}</td><td>${nv(c.arrGrowth,'%','f1')}</td><td class="${c.ruleOf40!=null&&c.ruleOf40>=40?'pos':''}">${nv(c.ruleOf40,'','f1')}</td><td class="${c.opMargin>=15?'pos':c.opMargin<0?'neg':''}">${c.opMargin.toFixed(1)}%</td><td>${c.grossMargin.toFixed(1)}%</td><td style="color:#999">${nv(c.grossChurn,'%','f2')}</td><td style="color:#999">${nv(c.nrr,'%','f1')}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['mtRev','mtGM','mtQRev']);
    const mx=Math.max(...companies.map(c=>c.marketCap));
    mc('mtRev','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.revenue/1000,y:c.opMargin,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'売上高 (十億円)'}},y:{title:{display:true,text:'営業利益率 (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: 売上${(x.raw.x*1000).toLocaleString()}百万 / OPM${x.raw.y}%`}},annotation:{annotations:{opm15:{type:'line',yMin:15,yMax:15,borderColor:'rgba(155,139,110,0.5)',borderWidth:1.5,borderDash:[4,3],label:{display:true,content:'営業利益率15%',position:'end',backgroundColor:'rgba(155,139,110,0.7)',color:'#fff',font:{size:9}}}}}}});
    mc('mtGM','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.grossMargin,y:c.opMargin,r:Math.max(4,Math.sqrt(c.marketCap/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'Gross Margin (%)'}},y:{title:{display:true,text:'営業利益率 (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: GM${x.raw.x}% / OPM${x.raw.y}%`}}}});
    const top10=sorted.slice(0,10);
    mc('mtQRev','line',{labels:QUARTERS,datasets:top10.map((c,i)=>({label:shortName(c.name),data:c.quarterlyRevenue,borderColor:P[i%P.length],fill:false,tension:0.3,pointRadius:3,borderWidth:2}))},{});
  }
  // ============ 07 UNIT ECONOMICS ============
  function rUnit(){
    const el=g('sec-unit');
    const sorted=[...companies].sort((a,b)=>b.grossMargin-a.grossMargin);
    el.innerHTML=`
      ${secH('07','Unit Economics & 収益構造','Gross Margin・営業利益率・コスト構造分析')}
      <div class="commentary">
        <strong>Unit Economics:</strong> LTV/CAC・Payback Period・S&M/R&D/G&A比率等のUnit Economics指標は、各社が個別にIR資料で開示しているもの以外は外部から算出困難なため、現在データ未取得です。<br>
        代替として、Gross Margin・営業利益率等の公開データを用いた収益構造分析を掲載します。<br><br>
        <strong>参考:</strong> Gross Margin平均<strong>${avg(companies,'grossMargin')}%</strong>、営業利益率平均<strong>${avg(companies,'opMargin')}%</strong>。
        高Gross Margin企業(${topN(companies,'grossMargin',3).map(c=>shortName(c.name)+c.grossMargin.toFixed(0)+'%').join('、')})はSaaSモデルの典型。
      </div>
      <div class="kpi-grid">
        ${kpi('平均Gross Margin',avg(companies,'grossMargin')+'%','','c-green')}
        ${kpi('平均営業利益率',avg(companies,'opMargin')+'%','','c-gold')}
        ${kpi('平均純利益率',avg(companies,'netMargin')+'%','','c-navy')}
        ${kpi('平均ROE',avg(companies,'roe')+'%','','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">Gross Margin vs 営業利益率</div><div class="chart-panel-sub">バブルサイズ=売上高</div><div class="chart-area tall"><canvas id="utGM"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">Gross Margin ランキング</div><div class="chart-panel-sub">セグメント別</div><div class="chart-area tall"><canvas id="utGMBar"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">収益構造一覧 (Gross Margin順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">財務データ: ${DATA_AS_OF.financials}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>Gross Margin</th><th>営業利益率</th><th>純利益率</th><th>ROE</th><th>売上高(百万)</th><th>営業利益(百万)</th><th>セグメント</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td style="font-weight:700">${c.grossMargin.toFixed(1)}%</td><td class="${c.opMargin>=15?'pos':c.opMargin<0?'neg':''}">${c.opMargin.toFixed(1)}%</td><td>${c.netMargin.toFixed(1)}%</td><td class="${c.roe>=15?'pos':''}">${c.roe.toFixed(1)}%</td><td>${c.revenue.toLocaleString()}</td><td>${c.opProfit.toLocaleString()}</td><td><span class="badge ${SEG_BADGE[c.segment]||''}">${SEGMENTS[c.segment]}</span></td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el);dc(['utGM','utGMBar']);
    const mxRev=Math.max(...companies.map(c=>c.revenue));
    mc('utGM','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg).map(c=>({x:c.grossMargin,y:c.opMargin,r:Math.max(4,Math.sqrt(c.revenue/mxRev)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'Gross Margin (%)'}},y:{title:{display:true,text:'営業利益率 (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: GM${x.raw.x}% / OPM${x.raw.y}%`}}}});
    mc('utGMBar','bar',{labels:sorted.map(c=>shortName(c.name)),datasets:[{data:sorted.map(c=>c.grossMargin),backgroundColor:sorted.map(c=>SC[c.segment]||'#777'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:v=>v.toFixed(1)+'%'}}});
  }
  // ============ 08 DETAIL ============
  function rDetail(){
    const el=g('sec-detail');
    if(!selComp)selComp=companies[0];
    const c=selComp;
    const peers=companies.filter(x=>x.segment===c.segment&&x.code!==c.code);
    el.innerHTML=`
      ${secH('08','個別企業分析','選択企業の詳細指標と同業比較')}
      <div class="inline-filters"><span class="f-label">企業選択</span>
        <select id="dtSel">${companies.map(x=>`<option value="${x.code}" ${x.code===c.code?'selected':''}>${x.code} ${x.name}</option>`).join('')}</select>
      </div>
      <div style="font-family:'Noto Serif JP',serif;font-size:1.4rem;font-weight:700;color:var(--navy);margin-bottom:4px;">${c.name}</div>
      <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:8px;">${c.code} / ${SEGMENTS[c.segment]} / ${c.products.join(', ')}</div>
      <div class="kpi-grid">
        ${kpi('株価',c.stockPrice.toLocaleString()+'円','','c-navy')}
        ${kpi('時価総額',c.marketCap.toLocaleString()+'億円','','c-navy')}
        ${kpi('売上高',c.revenue.toLocaleString()+'百万円','','c-gold')}
        ${kpi('営業利益率',c.opMargin.toFixed(1)+'%','','c-gold')}
        ${kpi('Gross Margin',c.grossMargin.toFixed(1)+'%','','c-green')}
        ${kpi('ROE',c.roe.toFixed(1)+'%','','c-navy')}
        ${kpi('PSR / PER',c.psr.toFixed(1)+' / '+c.per.toFixed(1),'','c-navy')}
        ${kpi('vs TOPIX (1Y)',(c.relativeReturn>0?'+':'')+c.relativeReturn.toFixed(1)+'%','','c-gold')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">四半期売上高推移</div><div class="chart-panel-sub">※推計値 (通期実績ベース)</div><div class="chart-area short"><canvas id="dtRev"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">四半期営業利益率推移</div><div class="chart-panel-sub">※推計値 (通期実績ベース)</div><div class="chart-area short"><canvas id="dtOPM"></canvas></div></div>
      </div>
      ${peers.length?`<div class="table-panel"><div class="table-header"><div class="table-header-title">同業比較 - ${SEGMENTS[c.segment]}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">企業名</th><th>売上高(百万)</th><th>営業利益率</th><th>Gross Margin</th><th>ROE</th><th>PSR</th><th>PER</th><th>vs TOPIX</th></tr></thead>
        <tbody><tr class="row-hl"><td><strong>${shortName(c.name)}</strong></td><td>${c.revenue.toLocaleString()}</td><td>${c.opMargin.toFixed(1)}%</td><td>${c.grossMargin.toFixed(1)}%</td><td>${c.roe.toFixed(1)}%</td><td>${c.psr.toFixed(1)}</td><td>${c.per.toFixed(1)}</td><td>${c.relativeReturn>0?'+':''}${c.relativeReturn.toFixed(1)}%</td></tr>
        ${peers.map(p=>`<tr><td>${shortName(p.name)}</td><td>${p.revenue.toLocaleString()}</td><td>${p.opMargin.toFixed(1)}%</td><td>${p.grossMargin.toFixed(1)}%</td><td>${p.roe.toFixed(1)}%</td><td>${p.psr.toFixed(1)}</td><td>${p.per.toFixed(1)}</td><td>${p.relativeReturn>0?'+':''}${p.relativeReturn.toFixed(1)}%</td></tr>`).join('')}
        </tbody></table></div></div>`:''}
      <div style="margin-top:16px;"><a href="${c.irUrl}" target="_blank" style="color:var(--navy);font-size:0.82rem;">IR情報ページを開く &rarr;</a></div>`;
    g('dtSel').addEventListener('change',e=>{selComp=companies.find(c=>c.code===e.target.value);rDetail();});
    dc(['dtRev','dtOPM']);
    if(c.quarterlyRevenue)mc('dtRev','bar',{labels:QUARTERS,datasets:[{data:c.quarterlyRevenue,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'top',color:'#999',font:{size:9},formatter:v=>v.toLocaleString()}}});
    if(c.quarterlyOpMargin)mc('dtOPM','line',{labels:QUARTERS,datasets:[{data:c.quarterlyOpMargin,borderColor:'#2d7a4f',fill:true,backgroundColor:'rgba(45,122,79,0.08)',tension:0.3,pointRadius:4,borderWidth:2}]},{plugins:{legend:{display:false}}});
  }
  // ============ A DATA SOURCE ============
  function rSource(){
    const el=g('sec-source');
    el.innerHTML=`
      ${secH('A','データソース・方法論','データ取得元と更新方針')}
      <div class="commentary">
        <strong>データソース:</strong> 各社の決算短信・決算説明会資料・有価証券報告書から取得。株価データはYahoo Finance Japan準拠。<br>
        <strong>更新頻度:</strong> 四半期決算発表後に順次更新(目安: 決算発表後2週間以内)。<br>
        <strong>注意事項:</strong> 株価・時価総額・売上高・営業利益・純利益等の基本財務データはkabutan.jpの実績値を使用。
        ARR・NRR・Churn等のSaaS固有KPIは各社決算説明資料から取得(マネーフォワード・freee・Sansan・HENNGE・サイボウズ・サイバーセキュリティクラウド・スマレジ・トヨクモ等)。<strong>非開示企業は「-」(null)表示</strong>。推計値は一切含まれておらず、各社IR開示値のみ使用。<br>
        <strong>※四半期データについて:</strong> 四半期売上高・四半期営業利益率は、通期実績値を基に四半期按分・季節性調整を行った<strong>推計値</strong>です。
        各社が個別に開示する四半期決算短信の実績値とは異なる場合があります。投資判断にあたっては、各社IRの一次情報をご確認ください。<br>
        <strong>※SaaSインデックスについて:</strong> 対象30社の時価総額加重平均株価騰落率を月次配分して算出した推計指数です。
        TOPIX月次データはstooq.com実績値(2025年1月=100で正規化)を使用。
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">各社IR情報リンク</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th style="text-align:left">セグメント</th><th style="text-align:left">主要プロダクト</th><th style="text-align:left">IR URL</th></tr></thead>
        <tbody>${companies.map(c=>`<tr><td>${c.code}</td><td><strong>${c.name}</strong></td><td><span class="badge ${SEG_BADGE[c.segment]||''}">${SEGMENTS[c.segment]}</span></td><td style="font-size:0.72rem">${c.products.join(', ')}</td><td><a href="${c.irUrl}" target="_blank" style="color:var(--navy);font-size:0.72rem;">${c.irUrl}</a></td></tr>`).join('')}</tbody>
      </table></div></div>
      <div class="commentary" style="margin-top:24px;font-size:0.78rem;">
        <strong>SaaS KPI定義:</strong><br>
        ・<strong>ARR (Annual Recurring Revenue):</strong> 年間経常収益 = MRR x 12<br>
        ・<strong>NRR (Net Revenue Retention):</strong> 既存顧客からの売上維持率(アップセル含む)。100%超 = 既存顧客からの成長<br>
        ・<strong>Gross Churn Rate:</strong> 月次解約率(ダウングレード含む)<br>
        ・<strong>LTV (Lifetime Value):</strong> 顧客生涯価値 = ARPU x 粗利率 / Churn Rate<br>
        ・<strong>CAC (Customer Acquisition Cost):</strong> 顧客獲得コスト = S&M費用 / 新規獲得顧客数<br>
        ・<strong>Rule of 40:</strong> ARR成長率 + 営業利益率。40以上が健全<br>
        ・<strong>EV/ARR:</strong> 企業価値 / ARR。SaaS企業のバリュエーション指標<br>
        ・<strong>FCF Yield:</strong> フリーキャッシュフロー / 時価総額<br>
        ・<strong>ROIC:</strong> 投下資本利益率 = NOPAT / 投下資本
      </div>`;
  }
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
      'コード','企業名','セグメント','主要プロダクト','IR URL',
      '株価(円)','時価総額(億円)','発行済株式数(千株)',
      '株価騰落率1Y(%)','vs TOPIX(pp)',
      '売上高(百万円)','営業利益(百万円)','純利益(百万円)',
      '営業利益率(%)','純利益率(%)','Gross Margin(%)',
      'PER(倍)','PBR(倍)','PSR(倍)','ROE(%)',
      '四半期売上Q1','四半期売上Q2','四半期売上Q3','四半期売上Q4',
      '四半期売上Q5','四半期売上Q6','四半期売上Q7','四半期売上Q8',
      '四半期OPM Q1','四半期OPM Q2','四半期OPM Q3','四半期OPM Q4',
      '四半期OPM Q5','四半期OPM Q6','四半期OPM Q7','四半期OPM Q8'
    ];
    const rows = companies.map(c => [
      c.code, c.name, SEGMENTS[c.segment]||c.segment, c.products.join(' / '), c.irUrl,
      c.stockPrice, c.marketCap, c.shares,
      c.stockReturn1Y, c.relativeReturn,
      c.revenue, c.opProfit, c.netProfit,
      c.opMargin, c.netMargin, c.grossMargin,
      c.per, c.pbr, c.psr, c.roe,
      ...(c.quarterlyRevenue||[]),
      ...(c.quarterlyOpMargin||[])
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
    a.download = 'SaaSセクター分析_データ一覧_' + new Date().toISOString().slice(0,10) + '.csv';
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
