// ============================================================
// 小売セクター分析レポート 共通レンダラー (ドラフト版)
// ============================================================
// 各ページJS ({slug}-page.js) で window.SECTOR_CONFIG と
// グローバル companies (var) を定義した後に読み込むこと。
// 構成タブ: 01 Executive Summary / 02 財務指標・バリュエーション /
//           03 月次KPI横断比較 / 04 個別企業分析 / A 業界統計・マクロ付録
// SPEEDA等の実データ受領後は各ページJSのデータ定義のみ差し替える。
// ============================================================
(() => {
  'use strict';
  const CFG = window.SECTOR_CONFIG;
  const companies = window.companies || [];
  if (!CFG) { console.error('SECTOR_CONFIG が未定義です'); return; }

  Chart.defaults.color = '#777777';
  Chart.defaults.borderColor = '#e0ddd6';
  Chart.defaults.font.family = "'Noto Sans JP', system-ui";
  Chart.defaults.font.size = 11;
  Chart.register(ChartDataLabels);
  Chart.defaults.plugins.datalabels = { display: false };

  const P = ['#1a2d4f','#9b8b6e','#2d7a4f','#b53a3a','#5a7fa8','#c8946e','#6b8e5e','#8b6b8e','#4a8b8b','#a89b5a','#7a5a3a','#5a6b8e'];
  const SEGMENTS = CFG.segments;
  const SC = CFG.segColors;
  const MONTHS = CFG.months;
  const SERIES = CFG.monthlySeries; // [{field,label,sub,industryAvg?,industryAvgLabel?}]

  let tab = 'exec', selComp = null, mCodes = [], mSeg = 'ALL';
  const C = {};

  // ---------- helpers ----------
  function g(id){ return document.getElementById(id); }
  function fmtBil(v){ if(v==null)return'-'; if(v>=1000000)return (v/1000000).toFixed(2)+'兆円'; if(v>=10000)return (v/10000).toFixed(1)+'百億円'; return v.toLocaleString()+'百万円'; }
  function shortName(n){ return n.replace(/ホールディングス|HD|グループ/g,'').trim(); }
  function avg(arr,key){ const vals=arr.map(c=>c[key]).filter(v=>v!=null); return vals.length?(vals.reduce((s,v)=>s+v,0)/vals.length).toFixed(1):'-'; }
  function topN(arr,key,n){ return [...arr].filter(c=>c[key]!=null).sort((a,b)=>b[key]-a[key]).slice(0,n); }
  function nv(v,suf='',fmt){ if(v==null)return'-'; if(fmt==='loc')return v.toLocaleString()+suf; if(fmt==='f1')return v.toFixed(1)+suf; if(fmt==='f2')return v.toFixed(2)+suf; return v+suf; }
  function sv(v,d=1){ return v==null?'-':(v>0?'+':'')+v.toFixed(d)+'%'; }
  function hasMonthly(c){ return SERIES.some(s=>Array.isArray(c[s.field])&&c[s.field].some(v=>v!=null)); }
  function mc(id,type,data,opts={}){
    const ctx=g(id); if(!ctx)return;
    const base={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#777',font:{size:10}}},datalabels:{display:false}}};
    if(['bar','line','scatter','bubble'].includes(type))base.scales={x:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}},y:{ticks:{color:'#999',font:{size:9}},grid:{color:'#eae7e1'}}};
    C[id]=new Chart(ctx,{type,data,options:dm(base,opts)});
  }
  function dc(ids){ ids.forEach(id=>{ if(C[id]){ C[id].destroy(); delete C[id]; } }); }
  function dm(t,s){ const o={...t}; for(const k of Object.keys(s)){ if(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))o[k]=dm(o[k]||{},s[k]); else o[k]=s[k]; } return o; }
  function secH(n,t,d){ return `<div class="sec-header"><div class="sec-num">SECTION ${n}</div><div class="sec-title">${t}</div><div class="sec-desc">${d}</div></div>`; }
  function kpi(l,v,s,cls){ return `<div class="kpi-card ${cls}"><div class="kpi-label">${l}</div><div class="kpi-value">${v}</div>${s?`<div class="kpi-sub">${s}</div>`:''}</div>`; }
  function rankCard(t,items,fn){ return `<div class="ranking-card"><div class="ranking-title">${t}</div>${items.map((c,i)=>`<div class="ranking-row"><span><span class="ranking-num">${i+1}</span>${shortName(c.name)}</span><span style="font-weight:600">${fn(c)}</span></div>`).join('')}</div>`; }
  function bindRows(el){ el.querySelectorAll('.clickable-row').forEach(tr=>tr.addEventListener('click',()=>{ selComp=companies.find(c=>c.code===tr.dataset.code); document.querySelector('.nav-item[data-tab="detail"]').click(); })); }
  const DRAFT_NOTE = '<div class="commentary" style="border-left-color:#c2912e;background:#fdf9ef;"><strong>ドラフト版:</strong> 本ページの数値はレポート構成確認用の<strong>仮置き値</strong>です。SPEEDA・各社IR月次開示等の実データ受領後に全数値を差し替えます。</div>';

  // ---------- nav ----------
  initNav();
  function initNav(){
    const navInner=g('mainNav'), btnL=g('navScrollLeft'), btnR=g('navScrollRight');
    function updateScrollBtns(){
      if(!navInner||!btnL||!btnR)return;
      btnL.classList.toggle('hidden',navInner.scrollLeft<=4);
      btnR.classList.toggle('hidden',navInner.scrollLeft+navInner.clientWidth>=navInner.scrollWidth-4);
    }
    if(navInner){ navInner.addEventListener('scroll',updateScrollBtns); window.addEventListener('resize',updateScrollBtns); setTimeout(updateScrollBtns,100); }
    if(btnL)btnL.addEventListener('click',()=>{navInner.scrollBy({left:-200,behavior:'smooth'});});
    if(btnR)btnR.addEventListener('click',()=>{navInner.scrollBy({left:200,behavior:'smooth'});});
    document.querySelectorAll('.nav-item').forEach(n=>n.addEventListener('click',()=>{
      document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
      n.classList.add('active');
      tab=n.dataset.tab;
      document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
      g('sec-'+tab).classList.add('active');
      n.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
      render();
    }));
  }
  function render(){
    const fns={ exec:rExec, financial:rFinancial, monthly:rMonthly, detail:rDetail, macro:rMacro };
    if(fns[tab])fns[tab]();
  }

  // ============ 01 EXECUTIVE SUMMARY ============
  function rExec(){
    const el=g('sec-exec');
    const tm=companies.reduce((s,c)=>s+(c.marketCap||0),0);
    const tr=companies.reduce((s,c)=>s+(c.revenue||0),0);
    const ts=companies.reduce((s,c)=>s+(c.stores||0),0);
    const aOP=avg(companies,'opMargin'), aROE=avg(companies,'roe'), aPER=avg(companies,'per'), aDIV=avg(companies,'dividendYield');
    const topMC=topN(companies,'marketCap',3), topOP=topN(companies,'opMargin',3), topROE=topN(companies,'roe',3);
    el.innerHTML=`
      ${secH('01','Executive Summary',CFG.sectorName+'セクター全体概況と主要指標ハイライト')}
      ${DRAFT_NOTE}
      <div class="commentary">${CFG.execCommentary}</div>
      <div class="kpi-grid">
        ${kpi('対象企業数',companies.length+'社','ドラフト選定','c-navy')}
        ${kpi('時価総額合計',fmtBil(tm),'仮置き値','c-navy')}
        ${kpi('売上高合計',fmtBil(tr),'仮置き値','c-gold')}
        ${kpi('平均営業利益率',aOP+'%','','c-green')}
        ${kpi('平均ROE',aROE+'%','セクター平均','c-gold')}
        ${kpi('平均PER',aPER+'倍','','c-navy')}
        ${kpi('平均配当利回り',aDIV+'%','','c-green')}
        ${kpi('総'+ (CFG.storesLabel||'店舗数'), ts?ts.toLocaleString():'-','','c-navy')}
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">時価総額ランキング TOP15</div><div class="chart-panel-sub">単位: 百億円 / 仮置き値</div><div class="chart-area tall"><canvas id="exMC"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">営業利益率 vs ROE</div><div class="chart-panel-sub">バブルサイズ=時価総額 / セグメント色分け</div><div class="chart-area tall"><canvas id="exBub"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">セグメント別 時価総額構成比</div><div class="chart-area"><canvas id="exPie"></canvas></div></div>
        <div class="chart-panel">
          <div class="chart-panel-title">主要ランキング</div>
          <div style="display:grid;gap:10px;padding-top:8px;">
            ${rankCard('時価総額',topMC,c=>fmtBil(c.marketCap))}
            ${rankCard('営業利益率',topOP,c=>nv(c.opMargin,'%','f1'))}
            ${rankCard('ROE',topROE,c=>nv(c.roe,'%','f1'))}
          </div>
        </div>
      </div>`;
    dc(['exMC','exPie','exBub']);
    const s15=[...companies].sort((a,b)=>(b.marketCap||0)-(a.marketCap||0)).slice(0,15);
    mc('exMC','bar',{labels:s15.map(c=>shortName(c.name)),datasets:[{data:s15.map(c=>Math.round(c.marketCap/10000*10)/10),backgroundColor:s15.map(c=>SC[c.segment]||'#777'),borderWidth:0}]},{indexAxis:'y',plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',color:'#999',font:{size:9},formatter:v=>v.toLocaleString()+'百億'}}});
    const segMC={}; companies.forEach(c=>{segMC[c.segment]=(segMC[c.segment]||0)+(c.marketCap||0);});
    mc('exPie','doughnut',{labels:Object.keys(segMC).map(k=>SEGMENTS[k]||k),datasets:[{data:Object.values(segMC),backgroundColor:Object.keys(segMC).map(k=>SC[k]||'#777'),borderWidth:0}]},{plugins:{legend:{position:'right'},datalabels:{display:true,color:'#fff',font:{size:10,weight:600},formatter:(v,ctx)=>{const t=ctx.dataset.data.reduce((a,b)=>a+b,0);return(v/t*100).toFixed(1)+'%';}}}});
    const mx=Math.max(...companies.map(c=>c.marketCap||0));
    mc('exBub','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.opMargin!=null&&c.roe!=null).map(c=>({x:c.opMargin,y:c.roe,r:Math.max(4,Math.sqrt((c.marketCap||0)/mx)*30),name:c.name})),backgroundColor:SC[seg]+'88',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'営業利益率 (%)'}},y:{title:{display:true,text:'ROE (%)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ${x.raw.x}% / ROE${x.raw.y}%`}}}});
  }

  // ============ 02 FINANCIAL ============
  function rFinancial(){
    const el=g('sec-financial');
    const sorted=[...companies].sort((a,b)=>(b.revenue||0)-(a.revenue||0));
    el.innerHTML=`
      ${secH('02','財務指標・バリュエーション','売上高・利益率・PER/PBRの横断比較 (仮置き値)')}
      ${DRAFT_NOTE}
      <div class="commentary">
        <strong>バリュエーション分析 (ドラフト):</strong> セクター平均PERは${avg(companies,'per')}倍、PBRは${avg(companies,'pbr')}倍(いずれも仮置き値ベース)。
        実データ反映後に、PBR-ROE相関・PBR1倍割れ企業のスクリーニング・アクティビスト注目領域の分析コメントを追記予定。
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">売上高 vs 営業利益</div><div class="chart-panel-sub">TOP15 / 百万円</div><div class="chart-area tall"><canvas id="fnRP"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">PBR vs ROE マトリクス</div><div class="chart-panel-sub">4象限分析 / バブルサイズ=時価総額 / 赤点線=PBR1.0倍</div><div class="chart-area tall"><canvas id="fnPBR"></canvas></div></div>
      </div>
      <div class="chart-row">
        <div class="chart-panel"><div class="chart-panel-title">PER vs PBR</div><div class="chart-panel-sub">左下=割安ゾーン</div><div class="chart-area"><canvas id="fnPERPBR"></canvas></div></div>
        <div class="chart-panel"><div class="chart-panel-title">営業利益率分布</div><div class="chart-area"><canvas id="fnHist"></canvas></div></div>
      </div>
      <div class="table-panel"><div class="table-header"><div class="table-header-title">財務指標一覧 (売上高順)</div><div style="font-size:0.72rem;color:#999;margin-top:2px;">全数値: ドラフト仮置き値</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">コード</th><th style="text-align:left">企業名</th><th>決算期</th><th>セグメント</th><th>売上高</th><th>営業利益</th><th>純利益</th><th>営業利益率</th><th>純利益率</th><th>PER</th><th>PBR</th><th>ROE</th><th>配当利回り</th></tr></thead>
        <tbody>${sorted.map(c=>`<tr class="clickable-row" data-code="${c.code}"><td>${c.code}</td><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem;white-space:nowrap">${c.fiscalYear||'-'}</td><td><span class="badge">${SEGMENTS[c.segment]||''}</span></td><td>${nv(c.revenue,'','loc')}</td><td>${nv(c.opProfit,'','loc')}</td><td>${nv(c.netProfit,'','loc')}</td><td class="${c.opMargin!=null&&c.opMargin>=8?'pos':c.opMargin!=null&&c.opMargin<2?'neg':''}">${nv(c.opMargin,'%')}</td><td>${nv(c.netMargin,'%')}</td><td>${nv(c.per,'','f1')}</td><td>${nv(c.pbr,'','f1')}</td><td class="${c.roe!=null&&c.roe>=12?'pos':''}">${nv(c.roe,'%')}</td><td>${nv(c.dividendYield,'%','f1')}</td></tr>`).join('')}</tbody>
      </table></div></div>`;
    bindRows(el); dc(['fnRP','fnPBR','fnPERPBR','fnHist']);
    const t15=sorted.slice(0,15);
    mc('fnRP','bar',{labels:t15.map(c=>shortName(c.name)),datasets:[{label:'売上高',data:t15.map(c=>c.revenue),backgroundColor:'rgba(26,45,79,0.6)',borderWidth:0},{label:'営業利益',data:t15.map(c=>c.opProfit),backgroundColor:'rgba(45,122,79,0.6)',borderWidth:0}]},{indexAxis:'y'});
    const mx=Math.max(...companies.map(c=>c.marketCap||0));
    mc('fnPBR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.roe!=null&&c.pbr!=null).map(c=>({x:c.roe,y:c.pbr,r:Math.max(4,Math.sqrt((c.marketCap||0)/mx)*28),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'ROE (%)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: ROE${x.raw.x}% / PBR${x.raw.y}倍`}}}});
    mc('fnPERPBR','bubble',{datasets:Object.keys(SEGMENTS).map(seg=>({label:SEGMENTS[seg],data:companies.filter(c=>c.segment===seg&&c.per!=null&&c.pbr!=null).map(c=>({x:c.per,y:c.pbr,r:Math.max(4,Math.sqrt((c.marketCap||0)/mx)*25),name:c.name})),backgroundColor:SC[seg]+'77',borderColor:SC[seg],borderWidth:1}))},{scales:{x:{title:{display:true,text:'PER (倍)'}},y:{title:{display:true,text:'PBR (倍)'}}},plugins:{tooltip:{callbacks:{label:x=>`${x.raw.name}: PER${x.raw.x}倍 / PBR${x.raw.y}倍`}}}});
    const bins=[0,2,4,6,8,10,15,60]; const hist=bins.slice(0,-1).map((_,i)=>companies.filter(c=>c.opMargin!=null&&c.opMargin>=bins[i]&&c.opMargin<bins[i+1]).length);
    mc('fnHist','bar',{labels:bins.slice(0,-1).map((b,i)=>`${b}-${bins[i+1]}%`),datasets:[{data:hist,backgroundColor:'rgba(26,45,79,0.5)',borderWidth:0}]},{plugins:{legend:{display:false}}});
  }

  // ============ 03 MONTHLY KPI ============
  function rMonthly(){
    const el=g('sec-monthly');
    const withData=companies.filter(hasMonthly);
    if(!mCodes.length)mCodes=withData.slice(0,5).map(c=>c.code);
    const vis=(mSeg==='ALL'?companies:companies.filter(c=>c.segment===mSeg)).filter(hasMonthly);
    const panels=SERIES.map((s,i)=>`<div class="chart-panel"><div class="chart-panel-title">${s.label}</div>${s.sub?`<div class="chart-panel-sub">${s.sub}</div>`:''}<div class="chart-area"><canvas id="mS${i}"></canvas></div></div>`);
    const rows=[]; for(let i=0;i<panels.length;i+=2)rows.push(`<div class="chart-row${i+1>=panels.length?' single':''}">${panels.slice(i,i+2).join('')}</div>`);
    el.innerHTML=`
      ${secH('03','月次KPI横断比較','複数企業・セグメント横断で'+SERIES.map(s=>s.label).join('・')+'を比較')}
      ${DRAFT_NOTE}
      <div class="commentary">${CFG.monthlyCommentary}</div>
      <div class="chip-wrap"><div class="chip-label">セグメントフィルタ</div>
        <div class="chip-select" id="mSegC">
          <div class="chip ${mSeg==='ALL'?'selected':''}" data-seg="ALL">全セグメント</div>
          ${Object.entries(SEGMENTS).map(([k,v])=>`<div class="chip ${mSeg===k?'selected':''}" data-seg="${k}">${v}</div>`).join('')}
        </div>
      </div>
      <div class="chip-wrap"><div class="chip-label">比較企業を選択 (複数可 / 月次開示企業のみ)</div>
        <div class="chip-select" id="mCompC">${vis.map(c=>`<div class="chip ${mCodes.includes(c.code)?'selected':''}" data-code="${c.code}">${shortName(c.name)}</div>`).join('')}</div>
      </div>
      ${rows.join('')}
      <div class="table-panel"><div class="table-header"><div class="table-header-title">${SERIES[0].label} ヒートマップ</div></div><div class="table-scroll"><table class="heatmap-table" id="hmKPI"></table></div></div>
      <div class="commentary" style="margin-top:24px;font-size:0.78rem;">${CFG.monthlyMethodology}</div>`;
    g('mSegC').querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{mSeg=ch.dataset.seg;mCodes=[];rMonthly();}));
    g('mCompC').querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{const c=ch.dataset.code;mCodes=mCodes.includes(c)?mCodes.filter(x=>x!==c):[...mCodes,c];updateMonthlyCharts();}));
    updateMonthlyCharts();
  }
  function updateMonthlyCharts(){
    const sel=companies.filter(c=>mCodes.includes(c.code));
    g('mCompC')?.querySelectorAll('.chip').forEach(ch=>ch.classList.toggle('selected',mCodes.includes(ch.dataset.code)));
    if(!sel.length)return;
    dc(SERIES.map((_,i)=>'mS'+i));
    SERIES.forEach((s,i)=>{
      const datasets=sel.map((c,j)=>({label:shortName(c.name),data:c[s.field]||[],borderColor:P[j%P.length],fill:false,tension:0.3,pointRadius:3,borderWidth:2}));
      if(Array.isArray(s.industryAvg))datasets.push({label:s.industryAvgLabel||'業界平均',data:s.industryAvg,borderColor:'#999',borderDash:[6,3],fill:false,tension:0.3,pointRadius:0,borderWidth:2});
      mc('mS'+i,'line',{labels:MONTHS,datasets},{});
    });
    // ヒートマップ (第1系列)
    const f=SERIES[0].field, t=g('hmKPI');
    if(t){
      const comps=sel.filter(c=>Array.isArray(c[f])&&c[f].some(v=>v!=null));
      let h=`<thead><tr><th style="text-align:left">企業名</th>${MONTHS.map(m=>`<th>${m}</th>`).join('')}<th>平均</th></tr></thead><tbody>`;
      comps.forEach(c=>{
        const v=c[f], valid=v.filter(x=>x!=null), a=valid.length?valid.reduce((s,x)=>s+x,0)/valid.length:null;
        h+=`<tr><td>${shortName(c.name)}</td>${v.map(x=>`<td class="${x==null?'':x>=CFG.heatmapPosThreshold?'pos':x<CFG.heatmapNegThreshold?'neg':''}">${x==null?'-':x.toFixed(1)}</td>`).join('')}<td style="font-weight:700">${a==null?'-':a.toFixed(1)}</td></tr>`;
      });
      h+='</tbody>'; t.innerHTML=h;
    }
  }

  // ============ 04 DETAIL ============
  function rDetail(){
    const el=g('sec-detail');
    if(!selComp)selComp=companies[0];
    const c=selComp;
    const peers=companies.filter(x=>x.segment===c.segment&&x.code!==c.code);
    el.innerHTML=`
      ${secH('04','個別企業分析','選択企業の詳細指標と同業比較 (仮置き値)')}
      <div class="inline-filters"><span class="f-label">企業選択</span>
        <select id="dtSel">${companies.map(x=>`<option value="${x.code}" ${x.code===c.code?'selected':''}>${x.code} ${x.name}</option>`).join('')}</select>
      </div>
      <div style="font-family:'Noto Serif JP',serif;font-size:1.4rem;font-weight:700;color:var(--navy);margin-bottom:4px;">${c.name}</div>
      <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:8px;">${c.code} / ${SEGMENTS[c.segment]||''}${c.stores?` / ${c.stores.toLocaleString()}${CFG.storesLabel||'店舗'}`:''} / 決算期: ${c.fiscalYear||'-'}</div>
      ${c.brands&&c.brands.length?`<div class="chip-select" style="margin-bottom:24px;">${c.brands.map(b=>`<span class="chip selected">${b}</span>`).join('')}</div>`:''}
      <div class="kpi-grid">
        ${kpi('株価(仮置き)',nv(c.stockPrice,'円','loc'),'','c-navy')}
        ${kpi('時価総額',c.marketCap?Math.round(c.marketCap/10000).toLocaleString()+'百億円':'-','','c-navy')}
        ${kpi('売上高',nv(c.revenue,'','loc'),'百万円','c-gold')}
        ${kpi('営業利益率',nv(c.opMargin,'%'),'','c-green')}
        ${kpi('ROE',nv(c.roe,'%'),'','c-gold')}
        ${kpi('PER / PBR',nv(c.per,'','f1')+' / '+nv(c.pbr,'','f1'),'','c-navy')}
        ${kpi('配当利回り',nv(c.dividendYield,'%','f1'),'','c-green')}
        ${kpi(CFG.storesLabel||'店舗数',c.stores?c.stores.toLocaleString():'-','','c-navy')}
      </div>
      <div class="chart-row">
        ${SERIES.slice(0,2).map((s,i)=>`<div class="chart-panel"><div class="chart-panel-title">${s.label}</div><div class="chart-area short"><canvas id="dtS${i}"></canvas></div></div>`).join('')}
      </div>
      ${peers.length?`<div class="table-panel"><div class="table-header"><div class="table-header-title">同業比較 - ${SEGMENTS[c.segment]||''}</div></div><div class="table-scroll"><table>
        <thead><tr><th style="text-align:left">企業名</th><th>決算期</th><th>売上高</th><th>営業利益率</th><th>ROE</th><th>PER</th><th>PBR</th><th>配当利回り</th></tr></thead>
        <tbody><tr class="row-hl"><td><strong>${shortName(c.name)}</strong></td><td style="font-size:0.75rem">${c.fiscalYear||'-'}</td><td>${nv(c.revenue,'','loc')}</td><td>${nv(c.opMargin,'%')}</td><td>${nv(c.roe,'%')}</td><td>${nv(c.per,'','f1')}</td><td>${nv(c.pbr,'','f1')}</td><td>${nv(c.dividendYield,'%','f1')}</td></tr>
        ${peers.map(p=>`<tr><td>${shortName(p.name)}</td><td style="font-size:0.75rem">${p.fiscalYear||'-'}</td><td>${nv(p.revenue,'','loc')}</td><td>${nv(p.opMargin,'%')}</td><td>${nv(p.roe,'%')}</td><td>${nv(p.per,'','f1')}</td><td>${nv(p.pbr,'','f1')}</td><td>${nv(p.dividendYield,'%','f1')}</td></tr>`).join('')}
        </tbody></table></div></div>`:''}`;
    g('dtSel').addEventListener('change',e=>{selComp=companies.find(x=>x.code===e.target.value);rDetail();});
    dc(['dtS0','dtS1']);
    SERIES.slice(0,2).forEach((s,i)=>{
      const v=c[s.field];
      const cv=g('dtS'+i); if(!cv)return;
      if(Array.isArray(v)&&v.some(x=>x!=null)){
        const valid=v.filter(x=>x!=null);
        mc('dtS'+i,'bar',{labels:MONTHS,datasets:[{data:v,backgroundColor:v.map(x=>x==null?'rgba(200,200,200,0.3)':x>=CFG.heatmapPosThreshold?'rgba(45,122,79,0.5)':x<CFG.heatmapNegThreshold?'rgba(181,58,58,0.5)':'rgba(26,45,79,0.5)'),borderWidth:0}]},{scales:{y:{min:Math.min(...valid)-3}},plugins:{legend:{display:false}}});
      } else {
        cv.parentElement.innerHTML='<p style="text-align:center;color:#999;padding:2em">月次データ非開示 / 未入力 (ドラフト)</p>';
      }
    });
  }

  // ============ APPENDIX: MACRO ============
  function rMacro(){
    const el=g('sec-macro');
    const charts=CFG.macroCharts||[];
    const panels=charts.map((mch,i)=>`<div class="chart-panel"><div class="chart-panel-title">${mch.title}</div>${mch.sub?`<div class="chart-panel-sub">${mch.sub}</div>`:''}<div class="chart-area${charts.length<=2?' tall':''}"><canvas id="maC${i}"></canvas></div></div>`);
    const rows=[]; for(let i=0;i<panels.length;i+=2)rows.push(`<div class="chart-row${i+1>=panels.length?' single':''}">${panels.slice(i,i+2).join('')}</div>`);
    el.innerHTML=`
      ${secH('A','業界統計・マクロ付録','業界全体の市場規模・外部環境データ (仮置き値)')}
      ${DRAFT_NOTE}
      <div class="commentary">${CFG.macroCommentary}</div>
      ${rows.join('')}`;
    dc(charts.map((_,i)=>'maC'+i));
    charts.forEach((mch,i)=>{
      const datasets=mch.series.map((s,j)=>mch.type==='line'
        ?{label:s.label,data:s.data,borderColor:P[j%P.length],fill:false,tension:0.3,pointRadius:3,borderWidth:2}
        :{label:s.label,data:s.data,backgroundColor:P[j%P.length]+(mch.stacked?'cc':'99'),borderWidth:0});
      const opts={};
      if(mch.stacked)opts.scales={x:{stacked:true},y:{stacked:true}};
      if(mch.series.length===1)opts.plugins={legend:{display:false}};
      if(mch.yTitle)opts.scales=dm(opts.scales||{},{y:{title:{display:true,text:mch.yTitle}}});
      mc('maC'+i,mch.type||'line',{labels:mch.labels,datasets},opts);
    });
  }

  // ============ DOWNLOAD ============
  function requireAuth(){ if(!window.currentUser){ window.openModal('login'); return false; } return true; }
  window.downloadPDF=function(){ if(!requireAuth())return; window.print(); };
  window.downloadCSV=function(){
    if(!requireAuth())return;
    const BOM='\uFEFF';
    const headers=['コード','企業名','セグメント','決算期','株価(円)','時価総額(百万円)','売上高(百万円)','営業利益(百万円)','純利益(百万円)','営業利益率(%)','純利益率(%)','PER(倍)','PBR(倍)','ROE(%)','配当利回り(%)',CFG.storesLabel||'店舗数',
      ...SERIES.flatMap(s=>MONTHS.map(m=>`${s.label}_${m}`))];
    const rows=companies.map(c=>[
      c.code,c.name,SEGMENTS[c.segment]||c.segment,c.fiscalYear||'',
      c.stockPrice,c.marketCap,c.revenue,c.opProfit,c.netProfit,c.opMargin,c.netMargin,
      c.per,c.pbr,c.roe,c.dividendYield,c.stores,
      ...SERIES.flatMap(s=>{const v=c[s.field]||[];return MONTHS.map((_,i)=>v[i]!=null?v[i]:'');})
    ]);
    const esc=v=>{const s=v==null?'':String(v);return s.includes(',')||s.includes('"')||s.includes('\n')?'"'+s.replace(/"/g,'""')+'"':s;};
    const csv=BOM+headers.map(esc).join(',')+'\n'+rows.map(r=>r.map(esc).join(',')).join('\n');
    const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    a.download=CFG.sectorName+'セクター分析_ドラフト_'+new Date().toISOString().slice(0,10)+'.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // ---------- init ----------
  const cc=g('companyCount'); if(cc)cc.textContent=companies.length;
  render();
})();

/* ── Event Delegation: data-action ── */
document.addEventListener('click',function(e){
  var el=e.target.closest('[data-action]');
  if(!el)return;
  var fn=el.dataset.action;
  if(fn==='downloadPDF')downloadPDF();
  else if(fn==='downloadCSV')downloadCSV();
});
