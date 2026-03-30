/* ===== API設定 ===== */
const API_BASE = (function() {
  const host = window.location.hostname;
  if (host.includes('vercel.app')) return '';
  if (host.includes('aoyama-nogizaka.com') || host.includes('github.io')) {
    return 'https://aoyama-nogizakapublic.vercel.app';
  }
  if (host === 'localhost' || host === '127.0.0.1') return 'http://localhost:3456';
  return 'https://aoyama-nogizakapublic.vercel.app';
})();

/* ===== データ保持 ===== */
let _stockData = { price: null, sharesIssued: null, companyName: '', marketCapOku: null };
let _edinetData = {};

/* ===== STOCK DATA FETCH ===== */
async function fetchStockData() {
  const code = document.getElementById('stockCode').value.trim().toUpperCase();
  const info = document.getElementById('stockInfo');
  const btn = document.getElementById('btnFetch');

  if (!code || code.length !== 4 || !/^[0-9A-Za-z]{4}$/.test(code)) {
    info.textContent = '4桁の証券コードを入力してください（例: 7203）。';
    info.className = 'stock-info error';
    return;
  }

  btn.disabled = true;
  btn.textContent = '取得中...';
  info.textContent = 'kabutan + Yahoo!ファイナンスからデータ取得中...';
  info.className = 'stock-info';
  document.querySelectorAll('.auto-filled').forEach(el => el.classList.remove('auto-filled'));
  _edinetData = {};

  try {
    let stockResult = null;

    // Phase 1: Stock API (kabutan + Yahoo Finance)
    try {
      const resp = await fetch(API_BASE + '/api/stock/' + code, { signal: AbortSignal.timeout ? AbortSignal.timeout(15000) : undefined });
      if (resp.ok) {
        const json = await resp.json();
        if (json.success && json.data) {
          stockResult = json.data;
          _stockData.price = stockResult.price;
          _stockData.sharesIssued = stockResult.sharesIssued;
          _stockData.companyName = stockResult.companyName || '';
          _stockData.marketCapOku = stockResult.marketCapOku;
        }
      }
    } catch (e) {
      console.log('Stock API error:', e.message);
    }

    // Phase 2: EDINET (search + XBRL)
    try {
      info.textContent = 'EDINET書類検索中...';
      const sRes = await fetch(API_BASE + '/api/edinet/search/' + code, { signal: AbortSignal.timeout ? AbortSignal.timeout(15000) : undefined });
      if (!sRes.ok) throw new Error('EDINET検索エラー');
      const sData = await sRes.json();
      if (sData.success && sData.documents && sData.documents.length > 0) {
        info.textContent = 'XBRL財務データ解析中...';
        const xRes = await fetch(API_BASE + '/api/edinet/xbrl/' + sData.documents[0].docID, { signal: AbortSignal.timeout ? AbortSignal.timeout(30000) : undefined });
        if (!xRes.ok) throw new Error('XBRLデータ取得エラー');
        const xData = await xRes.json();
        if (xData.success && xData.data) {
          _edinetData = xData.data;
        }
      }
    } catch (e) {
      console.log('EDINET fetch error:', e.message);
    }

    // Show reference data
    showRefData(_edinetData);

    // Map and apply
    if (stockResult) {
      const mapped = mapApiData(stockResult, _edinetData);
      applyStockData(mapped);
      const edinetNote = Object.keys(_edinetData).length > 0 ? '' : '（EDINET未取得 — 一部手動入力が必要）';
      info.textContent = '\u2713 ' + (stockResult.companyName || code) + ' のデータを反映しました。' + edinetNote;
      info.className = 'stock-info success';
    } else {
      info.textContent = 'この銘柄のデータは自動取得できませんでした。手動で各項目を入力してください。';
      info.className = 'stock-info error';
    }
  } catch (e) {
    console.error('Fetch error:', e);
    info.textContent = 'データ取得中にエラーが発生しました。手動で入力してください。';
    info.className = 'stock-info error';
  } finally {
    btn.disabled = false;
    btn.textContent = 'データ取得';
  }
}

function mapApiData(d, e) {
  e = e || {};
  const mcapM = d.marketCapOku ? d.marketCapOku * 100 : null; // 百万円

  // NC計算
  const cash = (e.cashAndDeposits || 0) + (e.shortTermSecurities || 0);
  const debt = (e.shortTermBorrowings || 0) + (e.currentPortionLongTermDebt || 0) +
               (e.longTermBorrowings || 0) + (e.bondsPayable || 0) + (e.currentPortionBonds || 0);
  const netCash = cash - debt;
  const ncRatio = (e.cashAndDeposits != null && mcapM) ? netCash / mcapM * 100 : null;

  // 実質NC
  const secMV = e.policyHoldingsMarketValue || e.securitiesMarketValue || 0;
  const adjNcRatio = (e.cashAndDeposits != null && secMV > 0 && mcapM) ? (netCash + secMV) / mcapM * 100 : null;

  // EV/EBITDA
  let evEbitda = null;
  if (mcapM && e.operatingIncome != null && e.depreciationAndAmortization != null) {
    const ev = mcapM + debt - cash;
    const ebitda = e.operatingIncome + e.depreciationAndAmortization;
    if (ebitda > 0) evEbitda = ev / ebitda;
  }

  // 政策保有/純資産
  const netAssets = e.netAssets || e.shareholdersEquity;
  const crossRatio = (e.policyHoldingsMarketValue && netAssets) ? e.policyHoldingsMarketValue / netAssets * 100 : null;

  return {
    companyName: d.companyName,
    pbr: d.pbr,
    roe: d.roe,
    payout: d.payoutRatio,
    equity: d.equityRatio,
    marketcap: d.marketCapOku,
    cash: ncRatio != null ? parseFloat(ncRatio.toFixed(1)) : null,
    adjNcRatio: adjNcRatio != null ? parseFloat(adjNcRatio.toFixed(1)) : null,
    evEbitda: evEbitda != null ? parseFloat(evEbitda.toFixed(1)) : null,
    foreign: e.foreignOwnership || null,
    outside: e.outsideDirectorRatio || null,
    crosshold: crossRatio != null ? parseFloat(crossRatio.toFixed(1)) : null,
    treasury: e.treasurySharesCount || (e.treasuryShares ? Math.round(e.treasuryShares * 1000000) : null),
  };
}

function applyStockData(data) {
  const fields = [
    { id: 'risk_pbr', key: 'pbr' },
    { id: 'risk_roe', key: 'roe' },
    { id: 'risk_payout', key: 'payout' },
    { id: 'risk_cash', key: 'cash' },
    { id: 'risk_adj_nc', key: 'adjNcRatio' },
    { id: 'risk_ev_ebitda', key: 'evEbitda' },
    { id: 'risk_equity', key: 'equity' },
    { id: 'risk_foreign', key: 'foreign' },
    { id: 'risk_crosshold', key: 'crosshold' },
    { id: 'risk_outside', key: 'outside' },
    { id: 'risk_marketcap', key: 'marketcap' },
    { id: 'risk_treasury', key: 'treasury' },
  ];
  fields.forEach(f => {
    const el = document.getElementById(f.id);
    if (data[f.key] != null) {
      el.value = typeof data[f.key] === 'number' ? (Number.isInteger(data[f.key]) ? data[f.key] : data[f.key]) : data[f.key];
      el.classList.add('auto-filled');
    }
  });
  _stockData.companyName = data.companyName || '';
  if (_stockData.sharesIssued && _stockData.price) {
    document.getElementById('marketcapHint').textContent =
      '株価 ' + _stockData.price.toLocaleString() + '円 \u00d7 発行済 ' + _stockData.sharesIssued.toLocaleString() + '株（自己株式未控除）';
  }
}

/* ===== 自己株式控除で時価総額再計算 ===== */
function recalcMarketCap() {
  const treasury = parseFloat(document.getElementById('risk_treasury').value);
  if (!isNaN(treasury) && _stockData.price && _stockData.sharesIssued) {
    const effective = _stockData.sharesIssued - treasury;
    const newCap = Math.round(_stockData.price * effective / 100000000);
    document.getElementById('risk_marketcap').value = newCap;
    document.getElementById('risk_marketcap').classList.add('auto-filled');
    document.getElementById('marketcapHint').textContent =
      '株価 ' + _stockData.price.toLocaleString() + '円 \u00d7 (発行済 ' + _stockData.sharesIssued.toLocaleString() + ' - 自己株式 ' + treasury.toLocaleString() + ') = ' + newCap.toLocaleString() + '億円';
  }
}

/* ===== RISK CALCULATION ===== */
function calculateRisk() {
  const _p = id => { const v = parseFloat(document.getElementById(id).value); return isNaN(v) ? null : v; };
  const pbr = _p('risk_pbr');
  const roe = _p('risk_roe');
  const payout = _p('risk_payout');
  const cash = _p('risk_cash');
  const adjNc = _p('risk_adj_nc');
  const evEbitda = _p('risk_ev_ebitda');
  const equity = _p('risk_equity');
  const foreign = _p('risk_foreign');
  const crosshold = _p('risk_crosshold');
  const outside = _p('risk_outside');
  const marketcap = _p('risk_marketcap');
  const defense = document.getElementById('risk_defense').value;

  const factors = [];
  let total = 0, maxTotal = 0;

  if (pbr !== null) { let s; if(pbr<0.5)s=15;else if(pbr<0.7)s=12;else if(pbr<0.8)s=10;else if(pbr<1.0)s=8;else if(pbr<1.3)s=5;else s=2; factors.push({name:'PBR',score:s,max:15,value:pbr.toFixed(2),desc:pbr<1.0?'1倍割れ: アクティビストの主要ターゲット':'PBR1倍以上: 相対的に低リスク'}); total+=s; maxTotal+=15; }
  if (roe !== null) { let s; if(roe<3)s=10;else if(roe<5)s=8;else if(roe<8)s=6;else if(roe<10)s=3;else s=1; factors.push({name:'ROE',score:s,max:10,value:roe.toFixed(1)+'%',desc:roe<8?'資本効率に改善余地あり':'資本効率は概ね良好'}); total+=s; maxTotal+=10; }
  if (payout !== null) { let s; if(payout<20)s=8;else if(payout<30)s=6;else if(payout<40)s=4;else if(payout<50)s=2;else s=1; factors.push({name:'配当性向',score:s,max:8,value:payout.toFixed(1)+'%',desc:payout<30?'株主還元に課題':'株主還元は概ね適切'}); total+=s; maxTotal+=8; }
  if (cash !== null) { let s; if(cash>50)s=12;else if(cash>30)s=10;else if(cash>20)s=7;else if(cash>10)s=4;else s=2; factors.push({name:'NC/時価総額',score:s,max:12,value:cash.toFixed(1)+'%',desc:cash>30?'過剰現金: 自社株買い・特別配当の要求対象':'キャッシュ水準は適正範囲'}); total+=s; maxTotal+=12; }
  if (adjNc !== null) { let s; if(adjNc>80)s=10;else if(adjNc>50)s=8;else if(adjNc>30)s=6;else if(adjNc>15)s=3;else s=1; factors.push({name:'実質NC/時価総額',score:s,max:10,value:adjNc.toFixed(1)+'%',desc:adjNc>30?'有価証券含め過剰現金':'実質キャッシュ水準は適正'}); total+=s; maxTotal+=10; }
  if (evEbitda !== null) { let s; if(evEbitda<3)s=8;else if(evEbitda<5)s=7;else if(evEbitda<7)s=5;else if(evEbitda<10)s=3;else s=1; factors.push({name:'EV/EBITDA',score:s,max:8,value:evEbitda.toFixed(1)+'倍',desc:evEbitda<7?'割安水準: アクティビストの関心対象':'EV/EBITDAは適正〜割高'}); total+=s; maxTotal+=8; }
  if (equity !== null) { let s; if(equity>80)s=7;else if(equity>70)s=5;else if(equity>60)s=3;else if(equity>50)s=2;else s=1; factors.push({name:'自己資本比率',score:s,max:7,value:equity.toFixed(1)+'%',desc:equity>70?'資本過剰の可能性':'資本構成は適正'}); total+=s; maxTotal+=7; }
  if (foreign !== null) { let s; if(foreign>30)s=8;else if(foreign>20)s=6;else if(foreign>15)s=4;else if(foreign>10)s=3;else s=1; factors.push({name:'外国人持株比率',score:s,max:8,value:foreign.toFixed(1)+'%',desc:foreign>25?'海外アクティビストが参入しやすい環境':'海外投資家の比率は低め'}); total+=s; maxTotal+=8; }
  if (crosshold !== null) { let s; if(crosshold>20)s=8;else if(crosshold>10)s=6;else if(crosshold>5)s=3;else s=1; factors.push({name:'政策保有株式比率',score:s,max:8,value:crosshold.toFixed(1)+'%',desc:crosshold>10?'ガバナンス上の懸念材料':'政策保有は限定的'}); total+=s; maxTotal+=8; }
  if (outside !== null) { let s; if(outside<25)s=8;else if(outside<33)s=6;else if(outside<50)s=4;else s=1; factors.push({name:'社外取締役比率',score:s,max:8,value:outside.toFixed(1)+'%',desc:outside<33?'取締役会の独立性に課題':'取締役会の独立性は確保'}); total+=s; maxTotal+=8; }
  if (marketcap !== null) { let s; if(marketcap<100)s=3;else if(marketcap<500)s=8;else if(marketcap<1000)s=7;else if(marketcap<3000)s=5;else if(marketcap<5000)s=4;else s=2; factors.push({name:'時価総額',score:s,max:8,value:marketcap.toLocaleString()+'億円',desc:marketcap>=100&&marketcap<=2000?'アクティビストのターゲットゾーン':'ターゲットゾーン外'}); total+=s; maxTotal+=8; }
  if (defense) { let s; if(defense==='none')s=2;else if(defense==='expired')s=4;else s=1; const labels={none:'なし',active:'導入中',expired:'廃止済み'}; factors.push({name:'買収防衛策',score:s,max:4,value:labels[defense],desc:defense==='active'?'防衛策あり（ただし廃止圧力の可能性）':'防衛策なし/廃止: エントリーしやすい'}); total+=s; maxTotal+=4; }

  if (!factors.length) { alert('少なくとも1つの項目を入力してください。'); return; }

  const normalizedScore = Math.round((total / maxTotal) * 100);
  const resultDiv = document.getElementById('riskResult');
  resultDiv.classList.add('show');

  const scoreEl = document.getElementById('riskScoreValue');
  const levelEl = document.getElementById('riskLevel');
  const descEl = document.getElementById('riskLevelDesc');
  const circle = document.getElementById('riskScoreCircle');
  const nameEl = document.getElementById('resultCompanyName');

  scoreEl.textContent = normalizedScore;
  if (_stockData.companyName) {
    nameEl.textContent = _stockData.companyName + '\uff08' + (document.getElementById('stockCode').value || '') + '\uff09';
  } else {
    nameEl.textContent = '';
  }

  let level, levelClass, bgClass, desc;
  if (normalizedScore >= 65) {
    level = '高リスク'; levelClass = 'risk-high'; bgClass = 'risk-bg-high';
    desc = 'アクティビストの標的となる可能性が高い状態です。資本効率・株主還元・ガバナンスの早急な改善が推奨されます。';
  } else if (normalizedScore >= 40) {
    level = '中リスク'; levelClass = 'risk-medium'; bgClass = 'risk-bg-medium';
    desc = 'アクティビストが関心を持つ可能性があります。予防的な対策を検討することを推奨します。';
  } else {
    level = '低リスク'; levelClass = 'risk-low'; bgClass = 'risk-bg-low';
    desc = '現時点でアクティビストの標的となるリスクは比較的低いと考えられます。';
  }

  circle.className = 'risk-score-circle ' + bgClass;
  scoreEl.style.color = levelClass === 'risk-high' ? 'var(--red)' : levelClass === 'risk-medium' ? 'var(--orange)' : 'var(--green)';
  levelEl.textContent = level;
  levelEl.className = 'risk-level ' + levelClass;
  descEl.textContent = desc;

  const grid = document.getElementById('riskDetailGrid');
  grid.innerHTML = factors.map(f => {
    const pct = (f.score / f.max * 100).toFixed(0);
    const barColor = pct >= 70 ? 'var(--red)' : pct >= 40 ? 'var(--orange)' : 'var(--green)';
    return '<div class="risk-detail-item"><h4>' + f.name + '</h4><div><span class="score">' + f.score + '</span><span class="max-score"> / ' + f.max + '</span></div><div style="font-size:0.72rem;color:var(--text-mid);margin-top:4px;">入力値: ' + f.value + '</div><div class="bar"><div class="bar-fill" style="width:' + pct + '%;background:' + barColor + ';"></div></div><div style="font-size:0.68rem;color:var(--text-light);margin-top:6px;">' + f.desc + '</div></div>';
  }).join('');

  const recs = document.getElementById('riskRecommendations');
  const recommendations = [];
  factors.forEach(f => {
    const pct = f.score / f.max * 100;
    if (pct >= 60) {
      if (f.name === 'PBR') recommendations.push('PBR改善: 成長戦略の明確化、IR強化による市場評価の向上を検討');
      if (f.name === 'ROE') recommendations.push('ROE改善: 低収益事業の見直し、資本効率を意識した経営目標の設定');
      if (f.name === '配当性向') recommendations.push('株主還元: 配当性向の引上げ、自社株買いの実施検討');
      if (f.name === 'NC/時価総額') recommendations.push('現金活用: 成長投資の実行、株主還元の強化（特別配当・自社株買い）');
      if (f.name === '実質NC/時価総額') recommendations.push('有価証券の見直し: 政策保有株式の売却による資本効率改善');
      if (f.name === 'EV/EBITDA') recommendations.push('企業価値向上: 収益力強化または適正な株主還元で割安解消');
      if (f.name === '政策保有株式比率') recommendations.push('政策保有の縮減: CGコード対応として計画的な売却を推進');
      if (f.name === '社外取締役比率') recommendations.push('取締役会改革: 社外取締役比率の引上げ（過半数目標）');
      if (f.name === '買収防衛策') recommendations.push('対話体制の整備: アクティビスト接触時の対応マニュアル策定');
    }
  });
  if (normalizedScore >= 50) {
    recommendations.push('専門アドバイザーへの相談: 資本政策・IR戦略の包括的な見直し');
  }
  recs.innerHTML = recommendations.map(r => '<li>' + r + '</li>').join('');

  resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showRefData(e) {
  const sec = document.getElementById('refDataSection');
  const grid = document.getElementById('refDataGrid');
  if (!e || Object.keys(e).length === 0) { sec.style.display = 'none'; return; }
  const fmt = v => v != null ? v.toLocaleString() : '—';
  const items = [
    { label: '有価証券時価', val: e.securitiesMarketValue },
    { label: '有価証券簿価', val: e.securitiesBookValue },
    { label: '有価証券含み損益', val: (e.securitiesMarketValue != null && e.securitiesBookValue != null) ? e.securitiesMarketValue - e.securitiesBookValue : null },
    { label: '政策保有株式時価', val: e.policyHoldingsMarketValue },
    { label: '現金・預金', val: e.cashAndDeposits },
    { label: '投資不動産時価', val: e.investmentPropertyFairValue },
    { label: '投資不動産簿価', val: e.investmentPropertyBookValue },
    { label: '純資産', val: e.netAssets || e.shareholdersEquity },
    { label: '営業利益', val: e.operatingIncome },
    { label: '減価償却費', val: e.depreciationAndAmortization },
  ];
  const hasAny = items.some(i => i.val != null);
  if (!hasAny) { sec.style.display = 'none'; return; }
  grid.innerHTML = items.filter(i => i.val != null).map(i => {
    const color = (i.label.includes('含み損益') && i.val < 0) ? 'var(--red)' : 'var(--text-dark)';
    return '<div style="background:var(--off-white);border-radius:6px;padding:12px;"><div style="font-size:0.72rem;color:var(--text-light);margin-bottom:4px;">' + i.label + '</div><div style="font-size:0.95rem;font-weight:500;color:' + color + ';">' + fmt(i.val) + ' <span style="font-size:0.75rem;color:var(--text-light);">百万円</span></div></div>';
  }).join('');
  sec.style.display = '';
}

function clearRisk() {
  document.querySelectorAll('#riskForm input, #riskForm select').forEach(el => { el.value = ''; el.classList.remove('auto-filled'); });
  document.getElementById('riskResult').classList.remove('show');
  document.getElementById('stockCode').value = '';
  document.getElementById('stockInfo').textContent = '';
  document.getElementById('stockInfo').className = 'stock-info';
  document.getElementById('marketcapHint').textContent = '100〜5,000億円がターゲットゾーン（配点: 8点）';
  document.getElementById('treasuryHint').textContent = '入力すると時価総額を自己株式控除で再計算';
  _stockData = { price: null, sharesIssued: null, companyName: '', marketCapOku: null };
  _edinetData = {};
  document.getElementById('refDataSection').style.display = 'none';
}

// Enterキーで取得
document.getElementById('stockCode').addEventListener('keypress', function(e) {
  if (e.key === 'Enter') fetchStockData();
});

/* ── Event Delegation: data-action ── */
document.addEventListener('click', function(e) {
  var btn = e.target.closest('[data-action]');
  if (!btn) return;
  var action = btn.dataset.action;
  if (action === 'fetchStockData') fetchStockData();
  else if (action === 'calculateRisk') calculateRisk();
  else if (action === 'clearRisk') clearRisk();
});

/* ── Event Delegation: data-input ── */
document.addEventListener('input', function(e) {
  var el = e.target.closest('[data-input]');
  if (!el) return;
  if (el.dataset.input === 'recalcMarketCap') recalcMarketCap();
});
