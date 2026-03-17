/* Mobile Nav */
function toggleMenu(){document.getElementById("mobileMenu").classList.toggle("open")}

/* API BASE */
var API_BASE=(function(){var h=window.location.hostname;
if(h.includes("vercel.app"))return "";
if(h.includes("aoyama-nogizaka.com")||h.includes("github.io"))return "https://aoyama-nogizakapublic.vercel.app";
if(h==="localhost"||h==="127.0.0.1")return "";
return "https://aoyama-nogizakapublic.vercel.app"})();

/* Auth UI */
function updateAuthUI(user){
  var lo=document.getElementById("authLoggedOut"),li=document.getElementById("authLoggedIn"),cb=document.getElementById("btnCsvExport");
  if(user){lo.classList.add("hidden");li.classList.remove("hidden");document.getElementById("authUserName").textContent=user.displayName||user.email;if(cb)cb.disabled=false}
  else{lo.classList.remove("hidden");li.classList.add("hidden");if(cb)cb.disabled=true}
}
function showRegister(){document.getElementById("authLoginForm").classList.add("hidden");var rf=document.getElementById("authResetForm");if(rf)rf.classList.add("hidden");document.getElementById("authRegisterForm").classList.remove("hidden")}
function showLogin(){document.getElementById("authRegisterForm").classList.add("hidden");var rf=document.getElementById("authResetForm");if(rf)rf.classList.add("hidden");document.getElementById("authLoginForm").classList.remove("hidden")}
async function doLogin(){var e=document.getElementById("authEmail").value.trim(),p=document.getElementById("authPass").value,m=document.getElementById("authLoginMsg");if(!e||!p){m.textContent="メールとパスワードを入力してください";m.className="auth-msg error";return}try{await window.firebaseLogin(e,p);m.textContent="ログインしました";m.className="auth-msg success"}catch(err){m.textContent="ログイン失敗: "+err.message;m.className="auth-msg error"}}
async function doRegister(){var n=document.getElementById("regName").value.trim(),c=document.getElementById("regCompany").value.trim(),e=document.getElementById("regEmail").value.trim(),p=document.getElementById("regPass").value,aff=document.getElementById("regAffiliation")?document.getElementById("regAffiliation").value:"",affCode=document.getElementById("regAffiliationCode")?document.getElementById("regAffiliationCode").value:"",jt=document.getElementById("regJobTitle")?document.getElementById("regJobTitle").value:"",m=document.getElementById("authRegMsg");if(!n||!e||!p){m.textContent="必須項目を入力してください";m.className="auth-msg error";return}if(!aff){m.textContent="所属種別を選択してください";m.className="auth-msg error";return}if(p.length<6){m.textContent="パスワードは6文字以上";m.className="auth-msg error";return}try{await window.firebaseRegister(e,p,n,c,aff,affCode,jt);m.textContent="登録しました。確認メールをお送りしましたのでご確認ください。";m.className="auth-msg success"}catch(err){m.textContent="登録失敗: "+err.message;m.className="auth-msg error"}}
function doLogout(){window.firebaseLogout()}
/* Password Reset */
function showResetPassword(){document.getElementById("authLoginForm").classList.add("hidden");document.getElementById("authRegisterForm").classList.add("hidden");document.getElementById("authResetForm").classList.remove("hidden")}
async function doResetPassword(){var e=document.getElementById("resetEmail").value.trim(),m=document.getElementById("authResetMsg");if(!e){m.textContent="メールアドレスを入力してください";m.className="auth-msg error";return}try{await window.firebaseResetPassword(e);m.textContent="リセット用メールを送信しました。メールをご確認ください。";m.className="auth-msg success"}catch(err){m.textContent="エラー: "+err.message;m.className="auth-msg error"}}

/* Affiliation Code Toggle */
function toggleAffiliationCode(){var v=document.getElementById("regAffiliation").value;var el=document.getElementById("regAffiliationCode");if(el)el.style.display=v==="listed_company"?"":"none"}

/* My Page */
var AFF_LABELS={listed_company:"上場企業",institutional_investor:"機関投資家",individual_investor:"個人投資家",consulting:"コンサルティングファーム",legal:"法律事務所・弁護士",media:"メディア・報道機関",academic:"学術・研究機関",other:"その他"};
var JOB_LABELS={executive:"経営者・役員",department_head:"部長・本部長",manager:"課長・マネージャー",ir_officer:"IR担当",legal_compliance:"法務・コンプライアンス",finance:"経理・財務",analyst:"アナリスト・ファンドマネージャー",consultant:"コンサルタント・アドバイザー",individual:"個人",other:"その他"};

function openMyPage(){
  var user=window.currentUser;if(!user){alert("ログインしてください");return}
  var d=document.getElementById("myPageOverlay");
  if(!d){
    d=document.createElement("div");d.id="myPageOverlay";
    d.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:2000;display:flex;align-items:center;justify-content:center;";
    d.onclick=function(ev){if(ev.target===d)d.style.display="none"};
    d.innerHTML='<div style="background:#fff;border-radius:12px;padding:40px;max-width:480px;width:90%;position:relative;">'
      +'<button onclick="document.getElementById(\x27myPageOverlay\x27).style.display=\x27none\x27" style="position:absolute;top:12px;right:16px;background:none;border:none;font-size:20px;cursor:pointer;color:#888;">&times;</button>'
      +'<h2 style="font-family:\x27Noto Serif JP\x27,serif;font-weight:400;font-size:20px;color:#1a2d4f;margin-bottom:20px;text-align:center;">アカウント情報</h2>'
      +'<table style="width:100%;border-collapse:collapse;font-size:14px;" id="myPageTable"></table>'
      +'<div style="display:flex;gap:12px;margin-top:24px;justify-content:center;flex-wrap:wrap;">'
      +'<button onclick="handleMyPageReset()" style="background:#1a2d4f;color:#fff;border:none;padding:10px 24px;border-radius:4px;font-size:13px;cursor:pointer;">パスワードを変更</button>'
      +'<button id="btnResendVerify" onclick="handleResendVerify()" style="background:#9b8b6e;color:#fff;border:none;padding:10px 24px;border-radius:4px;font-size:13px;cursor:pointer;display:none;">認証メールを再送</button>'
      +'</div>'
      +'<div id="myPageMsg" style="display:none;text-align:center;padding:10px 14px;border-radius:4px;font-size:13px;margin-top:16px;"></div>'
      +'</div>';
    document.body.appendChild(d);
  }
  d.style.display="flex";
  var verified=user.emailVerified;
  document.getElementById("btnResendVerify").style.display=verified?"none":"";
  var rows=[
    {l:"お名前",v:user.displayName||"-"},
    {l:"メール",v:user.email||""},
    {l:"メール認証",v:verified?'<span style="color:#27ae60;">認証済み ✓</span>':'<span style="color:#c0392b;">未認証</span>'},
    {l:"所属種別",v:"-"},{l:"役職",v:"-"},{l:"会社名",v:"-"},{l:"登録日",v:"-"}
  ];
  var tbl=document.getElementById("myPageTable");
  tbl.innerHTML=rows.map(function(r){return'<tr style="border-bottom:1px solid #e0dcd5;"><td style="padding:12px 0;font-weight:500;color:#1a2d4f;width:100px;">'+r.l+'</td><td style="padding:12px 0;">'+r.v+'</td></tr>'}).join("");
  window.firebaseGetProfile(user.uid).then(function(p){
    if(p){var cells=tbl.querySelectorAll("td");
      cells[7].textContent=AFF_LABELS[p.affiliation]||p.affiliation||"-";
      cells[9].textContent=JOB_LABELS[p.jobTitle]||p.jobTitle||"-";
      cells[11].textContent=p.company||"-";
      cells[13].textContent=p.createdAt?new Date(p.createdAt.seconds*1000).toLocaleDateString("ja-JP"):"-";
    }
  }).catch(function(e){console.error(e)});
}
async function handleMyPageReset(){var m=document.getElementById("myPageMsg");try{await window.firebaseResetPassword(window.currentUser.email);m.textContent="パスワードリセット用メールを送信しました。";m.style.display="block";m.style.background="#d1fae5";m.style.color="#065f46"}catch(e){m.textContent="エラー: "+e.message;m.style.display="block";m.style.background="#fee2e2";m.style.color="#991b1b"}}
async function handleResendVerify(){var m=document.getElementById("myPageMsg");try{var mod=await import("https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js");await mod.sendEmailVerification(window.currentUser);m.textContent="認証メールを再送しました。";m.style.display="block";m.style.background="#d1fae5";m.style.color="#065f46"}catch(e){m.textContent=e.code==="auth/too-many-requests"?"送信回数が多すぎます。しばらくしてからお試しください。":"エラー: "+e.message;m.style.display="block";m.style.background="#fee2e2";m.style.color="#991b1b"}}


/* CSV Export (Member Only) */
function exportCSV(){
  if(!window.currentUser){alert("CSV出力は会員限定機能です。ログインしてください。");return}
  if(rankResults.length===0){alert("データがありません");return}
  var hd=["順位","コード","銘柄名","スコア","PBR","PER","ROE(%)","配当利回り(%)","配当性向(%)","自己資本比率(%)","時価総額(億円)","市場","セクター","土地簿価(百万円)","推定土地含み益(百万円)","有価証券含み益(百万円)","投資不動産含み益(百万円)","含み益合計(百万円)"];
  var sr=rankResults.slice().sort(function(a,b){return(b.score||0)-(a.score||0)});
  var rows=sr.map(function(d,i){var tg=(d.estimatedLandGain||0)+(d.securitiesGain||0)+(d.investmentPropertyGain||0);return[i+1,d.code,d.companyName||"",d.score||"",d.pbr!=null?d.pbr.toFixed(2):"",d.per!=null?d.per.toFixed(1):"",d.roe!=null?d.roe.toFixed(1):"",d.dividendYield!=null?d.dividendYield.toFixed(2):"",d.payoutRatio!=null?d.payoutRatio.toFixed(1):"",d.equityRatio!=null?d.equityRatio.toFixed(1):"",d.marketCapOku!=null?d.marketCapOku:"",d.market||"",d.sector||"",d.land!=null?d.land:"",d.estimatedLandGain!=null?d.estimatedLandGain:"",d.securitiesGain!=null?d.securitiesGain:"",d.investmentPropertyGain!=null?d.investmentPropertyGain:"",d.hasEdinetData?tg:""]});
  var bom="﻿";
  var csv=bom+[hd].concat(rows).map(function(r){return r.map(function(v){return String.fromCharCode(34)+String(v).replace(/"/g,String.fromCharCode(34,34))+String.fromCharCode(34)}).join(",")}).join("\n");
  var blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
  var url=URL.createObjectURL(blob);var a=document.createElement("a");a.href=url;
  a.download="activist-screening_"+new Date().toISOString().slice(0,10)+".csv";
  a.click();URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════════
// タブ切替
// ═══════════════════════════════════════════════════════════════
function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach((t, i) => {
    t.classList.toggle('active', (tabId === 'individual' && i === 0) || (tabId === 'ranking' && i === 1));
  });
  document.querySelectorAll('.tab-panel').forEach(tc => tc.classList.remove('active'));
  document.getElementById('tab-' + tabId).classList.add('active');
}

// ═══════════════════════════════════════════════════════════════
// 個別銘柄分析
// ═══════════════════════════════════════════════════════════════
let indData = {}; // 取得データ保持

async function fetchIndividual() {
  const code = document.getElementById('indCode').value.trim().toUpperCase();
  if (!/^[0-9A-Za-z]{4}$/.test(code)) { alert('4桁の証券コードを入力してください（例: 7203）'); return; }

  const btn = document.getElementById('btnIndFetch');
  const loading = document.getElementById('indLoading');
  const loadingText = document.getElementById('indLoadingText');
  btn.disabled = true;
  loading.classList.add('active');

  indData = { code };

  try {
    // Yahoo Finance取得
    loadingText.textContent = 'kabutan + Yahoo Finance からデータ取得中...';
    const yRes = await fetch(API_BASE+'/api/stock/' + code);
    if (!yRes.ok) {
      console.warn('Stock API returned status:', yRes.status);
      alert('データ取得に失敗しました（HTTPステータス: ' + yRes.status + '）');
    } else {
      const yText = await yRes.text();
      try {
        const y = JSON.parse(yText);
        if (y.success && y.data) {
          Object.assign(indData, y.data);
        } else {
          alert('データ取得: ' + (y.error || '取得失敗'));
        }
      } catch (parseErr) {
        console.warn('API response was not JSON:', yText.substring(0, 200));
        alert('APIからの応答を解析できませんでした。しばらく待ってから再度お試しください。');
      }
    }

    // EDINET取得（EDINETコード未入力なら証券コードで検索）
    const apiKey = document.getElementById('edinetApiKey').value.trim();
    const edinetCode = document.getElementById('indEdinet').value.trim();
    const searchCode = edinetCode || code; // EDINETコード優先、なければ証券コード
    if (searchCode) {
      loadingText.textContent = 'EDINET書類検索中...';
      const apiParam = apiKey ? '?apiKey=' + encodeURIComponent(apiKey) : '';
      const sRes = await fetch(API_BASE+'/api/edinet/search/' + searchCode + apiParam);
      const sData = await sRes.json();
      if (sData.success && sData.documents.length > 0) {
        // 証券コードで検索した場合、EDINETコードを自動セット
        if (!edinetCode && sData.documents[0].edinetCode) {
          document.getElementById('indEdinet').value = sData.documents[0].edinetCode;
        }
        loadingText.textContent = 'XBRL財務データ解析中...';
        const xRes = await fetch(API_BASE+'/api/edinet/xbrl/' + sData.documents[0].docID + apiParam);
        const xData = await xRes.json();
        if (xData.success && xData.data) {
          indData.edinet = xData.data;
        }
      }
    }

    // 計算 & 表示
    calculateAndDisplay();

  } catch (err) {
    alert('エラー: ' + err.message);
  } finally {
    btn.disabled = false;
    loading.classList.remove('active');
  }
}

function calculateAndDisplay() {
  const d = indData;
  const e = d.edinet || {};

  // 基本指標
  setMetric('m_price', d.price, '円', v => v.toLocaleString());
  setMetric('m_pbr', d.pbr, '倍', null, v => v < 1.0 ? 'negative' : '');
  setMetric('m_per', d.per, '倍');
  setMetric('m_roe', d.roe, '%', null, v => v < 8 ? 'warning' : 'positive');
  setMetric('m_eps', d.eps, '円');
  setMetric('m_bps', d.bps, '円', v => v.toLocaleString());
  setMetric('m_divyield', d.dividendYield, '%', null, v => v >= 3 ? 'positive' : '');
  setMetric('m_payout', d.payoutRatio, '%', null, v => v < 30 ? 'warning' : '');

  // 時価総額
  setMetric('m_mcap', d.marketCapOku, '億円', v => v.toLocaleString());
  const mcapCatEl = document.getElementById('m_mcap_cat');
  if (d.marketCapOku != null) {
    const cat = getMarketCapCategory(d.marketCapOku);
    mcapCatEl.innerHTML = `<span class="mcap-cat ${cat.cls}">${cat.label}</span>`;
  } else {
    mcapCatEl.textContent = '-';
  }

  // EDINET財務データ
  setMetric('m_cash', e.cashAndDeposits, '百万円', v => Math.round(v).toLocaleString());
  setMetric('m_netassets', e.netAssets || e.shareholdersEquity, '百万円', v => Math.round(v).toLocaleString());
  setMetric('m_foreign', e.foreignOwnership, '%');
  setMetric('m_outside_dir', e.outsideDirectorRatio, '%');
  setMetric('m_shares', d.sharesIssued, '株', v => v.toLocaleString());
  setMetric('m_treasury', e.treasuryShares, '株', v => v.toLocaleString());

  // 有利子負債計算
  const debt = (e.shortTermBorrowings || 0) + (e.currentPortionLongTermDebt || 0) +
               (e.longTermBorrowings || 0) + (e.bondsPayable || 0) + (e.currentPortionBonds || 0);
  if (debt > 0) {
    setMetric('m_debt', debt, '百万円', v => Math.round(v).toLocaleString());
  }

  // ネットキャッシュ
  const cash = (e.cashAndDeposits || 0) + (e.shortTermSecurities || 0);
  const netCash = cash - debt;
  if (e.cashAndDeposits != null) {
    setMetric('m_netcash', netCash, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
  }

  // NC/時価総額
  if (e.cashAndDeposits != null && d.marketCapOku) {
    const ncRatio = netCash / (d.marketCapOku * 100) * 100; // 百万円→億円変換
    indData._ncRatio = ncRatio;
    setMetric('m_nc_ratio', ncRatio, '%', v => v.toFixed(1), v => v > 30 ? 'negative' : v > 10 ? 'warning' : 'positive');
  }

  // 実質NC（NC＋有価証券時価）
  const secMV = e.policyHoldingsMarketValue || e.securitiesMarketValue || 0;
  if (e.cashAndDeposits != null && secMV > 0) {
    const adjNetCash = netCash + secMV;
    setMetric('m_adj_netcash', adjNetCash, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
    if (d.marketCapOku) {
      const adjNcRatio = adjNetCash / (d.marketCapOku * 100) * 100;
      indData._adjNcRatio = adjNcRatio;
      setMetric('m_adj_nc_ratio', adjNcRatio, '%', v => v.toFixed(1), v => v > 30 ? 'negative' : v > 10 ? 'warning' : 'positive');
    }
  }

  // 含み資産NC（実質NC＋投資不動産時価）
  const ipFV = e.investmentPropertyFairValue || 0;
  const adjNetCashVal = (e.cashAndDeposits != null ? netCash + secMV : null);
  if (adjNetCashVal != null && ipFV > 0) {
    const fullNetCash = adjNetCashVal + ipFV;
    setMetric('m_full_netcash', fullNetCash, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
    if (d.marketCapOku) {
      const fullNcRatio = fullNetCash / (d.marketCapOku * 100) * 100;
      indData._fullNcRatio = fullNcRatio;
      setMetric('m_full_nc_ratio', fullNcRatio, '%', v => v.toFixed(1), v => v > 30 ? 'negative' : v > 10 ? 'warning' : 'positive');
    }
  }

  // EV/EBITDA
  if (d.marketCapOku && e.operatingIncome != null && e.depreciationAndAmortization != null) {
    const ev = d.marketCapOku * 100 + debt - cash; // 百万円
    const ebitda = e.operatingIncome + e.depreciationAndAmortization;
    if (ebitda > 0) {
      const evEbitda = ev / ebitda;
      indData._evEbitda = evEbitda;
      setMetric('m_ev_ebitda', evEbitda, '倍', v => v.toFixed(1), v => v < 5 ? 'positive' : v < 10 ? 'warning' : 'negative');
    }
  }

  // 自己資本比率
  setMetric('m_equity_ratio', d.equityRatio, '%');

  // 含み益計算
  let secGain = null;
  if (e.securitiesBookValue != null && e.securitiesMarketValue != null) {
    secGain = e.securitiesMarketValue - e.securitiesBookValue;
    document.getElementById('manual_sec_gain').value = Math.round(secGain);
  }
  setMetric('m_sec_gain', secGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');

  // 政策保有株式セクション表示
  const phSection = document.getElementById('policyHoldingsSection');
  if (phSection && e.policyHoldingsMarketValue != null && e.policyHoldingsTop) {
    phSection.style.display = '';
    document.getElementById('ph_count').textContent = e.policyHoldingsCount;
    document.getElementById('ph_report_total').textContent = Math.round(e.policyHoldingsMarketValue).toLocaleString() + ' 百万円';
    document.getElementById('ph_current_total').textContent = '-';
    document.getElementById('ph_gain').textContent = '-';
    renderPolicyHoldingsTable(e.policyHoldingsTop, null);
  } else if (phSection) {
    phSection.style.display = 'none';
  }

  let propGain = null;
  if (e.investmentPropertyBookValue != null && e.investmentPropertyFairValue != null) {
    propGain = e.investmentPropertyFairValue - e.investmentPropertyBookValue;
    document.getElementById('manual_prop_gain').value = Math.round(propGain);
  }
  setMetric('m_prop_gain', propGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');

  // 土地含み益
  setMetric('m_land_bv', e.land, '百万円', v => Math.round(v).toLocaleString());
  let landGain = null;
  if (e.estimatedLandGain != null) {
    landGain = e.estimatedLandGain;
    document.getElementById('manual_land_gain').value = Math.round(landGain);
  }
  setMetric('m_land_gain', landGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
  const methodEl = document.getElementById('m_land_gain_method');
  if (e.landGainMethod) {
    methodEl.textContent = '推定方法: ' + e.landGainMethod;
    methodEl.style.display = 'block';
  } else {
    methodEl.style.display = 'none';
  }

  // 土地明細分析セクションを表示（EDINETコードまたは証券コードがある場合）
  const landSection = document.getElementById('landAnalysisSection');
  if (landSection) {
    const edinetCode = document.getElementById('indEdinet').value.trim();
    const stockCode = document.getElementById('indCode').value.trim();
    landSection.style.display = (edinetCode || stockCode) ? '' : 'none';
  }

  // 含み益合計（有価証券 + 投資不動産 + 土地）
  const totalGain = (secGain || 0) + (propGain || 0) + (landGain || 0);
  if (secGain != null || propGain != null || landGain != null) {
    setMetric('m_total_gain', totalGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
  }

  // 修正純資産 & 実質PBR
  const netAssetsVal = e.netAssets || e.shareholdersEquity;
  if (netAssetsVal && d.marketCapOku) {
    const taxRate = 0.3;
    const adjNav = netAssetsVal + totalGain * (1 - taxRate);
    indData._adjNav = adjNav;
    setMetric('m_adj_nav', adjNav, '百万円', v => Math.round(v).toLocaleString());

    const mcapMillion = d.marketCapOku * 100;
    const adjPbr = mcapMillion / adjNav;
    indData._adjPbr = adjPbr;
    setMetric('m_adj_pbr', adjPbr, '倍', v => v.toFixed(2), v => v < 1.0 ? 'negative' : '');
  }

  // 企業名・市場情報
  document.getElementById('indCompanyName').textContent = d.companyName || d.code;
  document.getElementById('indMarketInfo').textContent =
    [d.code, d.market, d.sector].filter(Boolean).join(' | ');

  // スコア計算
  calculateScore();

  // 結果表示
  document.getElementById('indResult').classList.remove('hidden');
  document.getElementById('indResult').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function recalcIndividual() {
  // 手動入力値で上書き
  const manSecGain = parseFloat(document.getElementById('manual_sec_gain').value);
  const manPropGain = parseFloat(document.getElementById('manual_prop_gain').value);
  const manLandGain = parseFloat(document.getElementById('manual_land_gain').value);
  const manCross = parseFloat(document.getElementById('manual_cross').value);

  if (!isNaN(manSecGain)) indData._manualSecGain = manSecGain;
  if (!isNaN(manPropGain)) indData._manualPropGain = manPropGain;
  if (!isNaN(manLandGain)) indData._manualLandGain = manLandGain;
  if (!isNaN(manCross)) indData._manualCross = manCross;

  // 再計算
  const e = indData.edinet || {};
  const secGain = indData._manualSecGain != null ? indData._manualSecGain :
    (e.securitiesMarketValue != null && e.securitiesBookValue != null ? e.securitiesMarketValue - e.securitiesBookValue : null);
  const propGain = indData._manualPropGain != null ? indData._manualPropGain :
    (e.investmentPropertyFairValue != null && e.investmentPropertyBookValue != null ? e.investmentPropertyFairValue - e.investmentPropertyBookValue : null);
  const landGain = indData._manualLandGain != null ? indData._manualLandGain :
    (e.estimatedLandGain != null ? e.estimatedLandGain : null);

  setMetric('m_sec_gain', secGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
  setMetric('m_prop_gain', propGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');
  setMetric('m_land_gain', landGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');

  const totalGain = (secGain || 0) + (propGain || 0) + (landGain || 0);
  setMetric('m_total_gain', totalGain, '百万円', v => Math.round(v).toLocaleString(), v => v > 0 ? 'positive' : 'negative');

  const netAssetsVal = e.netAssets || e.shareholdersEquity;
  if (netAssetsVal && indData.marketCapOku) {
    const adjNav = netAssetsVal + totalGain * 0.7;
    indData._adjNav = adjNav;
    setMetric('m_adj_nav', adjNav, '百万円', v => Math.round(v).toLocaleString());

    const adjPbr = (indData.marketCapOku * 100) / adjNav;
    indData._adjPbr = adjPbr;
    setMetric('m_adj_pbr', adjPbr, '倍', v => v.toFixed(2), v => v < 1.0 ? 'negative' : '');
  }

  calculateScore();
}

// ── スコアリングエンジン ──
const SCORE_CRITERIA = [
  { key: 'pbr', name: 'PBR', max: 15, unit: '倍',
    val: () => indData.pbr,
    fn: v => v < 0.5 ? 15 : v < 0.7 ? 12 : v < 0.8 ? 10 : v < 1.0 ? 8 : v < 1.3 ? 5 : 2 },
  { key: 'adjPbr', name: '実質PBR', max: 10, unit: '倍',
    val: () => indData._adjPbr,
    fn: v => v < 0.3 ? 10 : v < 0.5 ? 8 : v < 0.7 ? 6 : v < 1.0 ? 4 : 1 },
  { key: 'roe', name: 'ROE', max: 10, unit: '%',
    val: () => indData.roe,
    fn: v => v < 3 ? 10 : v < 5 ? 8 : v < 8 ? 5 : v < 10 ? 3 : 1 },
  { key: 'payout', name: '配当性向', max: 8, unit: '%',
    val: () => indData.payoutRatio,
    fn: v => v < 20 ? 8 : v < 30 ? 6 : v < 40 ? 4 : v < 50 ? 2 : 1 },
  { key: 'ncRatio', name: 'ネットキャッシュ/時価総額', max: 12, unit: '%',
    val: () => indData._ncRatio,
    fn: v => v > 50 ? 12 : v > 30 ? 10 : v > 20 ? 7 : v > 10 ? 4 : 1 },
  { key: 'adjNcRatio', name: '実質NC/時価総額', max: 10, unit: '%',
    val: () => indData._adjNcRatio,
    fn: v => v > 80 ? 10 : v > 50 ? 8 : v > 30 ? 6 : v > 15 ? 3 : 1 },
  { key: 'evEbitda', name: 'EV/EBITDA', max: 10, unit: '倍',
    val: () => indData._evEbitda,
    fn: v => v < 3 ? 10 : v < 5 ? 8 : v < 7 ? 6 : v < 10 ? 4 : v < 15 ? 2 : 1 },
  { key: 'equity', name: '自己資本比率', max: 8, unit: '%',
    val: () => indData.equityRatio,
    fn: v => v > 80 ? 8 : v > 70 ? 6 : v > 60 ? 4 : v > 50 ? 2 : 1 },
  { key: 'foreign', name: '外国人持株比率', max: 8, unit: '%',
    val: () => (indData.edinet || {}).foreignOwnership,
    fn: v => v > 30 ? 8 : v > 20 ? 6 : v > 15 ? 4 : v > 10 ? 3 : 1 },
  { key: 'cross', name: '政策保有株式/純資産', max: 8, unit: '%',
    val: () => indData._manualCross,
    fn: v => v > 20 ? 8 : v > 10 ? 6 : v > 5 ? 3 : 1 },
  { key: 'outside', name: '社外取締役比率', max: 8, unit: '%',
    val: () => (indData.edinet || {}).outsideDirectorRatio,
    fn: v => v < 25 ? 8 : v < 33 ? 5 : v < 50 ? 3 : 1 },
  { key: 'mcap', name: '時価総額', max: 8, unit: '億円',
    val: () => indData.marketCapOku,
    fn: v => v < 100 ? 3 : v < 500 ? 8 : v < 1000 ? 7 : v < 3000 ? 5 : v < 5000 ? 4 : 2 },
];

function calculateScore() {
  const tbody = document.getElementById('indDetailBody');
  tbody.innerHTML = '';
  let total = 0;

  for (const c of SCORE_CRITERIA) {
    const raw = c.val();
    let score = 0;
    let display = '-';
    if (raw != null && !isNaN(raw)) {
      score = c.fn(raw);
      display = (typeof raw === 'number' ? (Number.isInteger(raw) ? raw : raw.toFixed(2)) : raw) + ' ' + c.unit;
    }
    total += score;
    const pct = c.max > 0 ? (score / c.max * 100) : 0;
    const barCls = pct >= 70 ? 'bar-high' : pct >= 40 ? 'bar-mid' : 'bar-low';

    tbody.innerHTML += `<tr>
      <td>${c.name}</td>
      <td>${display}</td>
      <td style="text-align:right;font-weight:700;">${score}</td>
      <td style="text-align:right;color:var(--muted);">${c.max}</td>
      <td><div class="bar-container"><div class="bar-fill ${barCls}" style="width:${pct}%"></div></div></td>
    </tr>`;
  }

  // 合計表示
  const maxTotal = SCORE_CRITERIA.reduce((s, c) => s + c.max, 0);
  document.getElementById('indTotalScore').textContent = total;

  const badge = document.getElementById('indScoreBadge');
  if (total >= 65) {
    badge.textContent = '高リスク — アクティビスト標的可能性高';
    badge.className = 'score-badge badge-high';
  } else if (total >= 40) {
    badge.textContent = '中リスク — 注意が必要';
    badge.className = 'score-badge badge-mid';
  } else {
    badge.textContent = '低リスク';
    badge.className = 'score-badge badge-low';
  }
}

// ── ユーティリティ ──
function setMetric(id, value, unit, fmt, colorFn) {
  const el = document.getElementById(id);
  if (value != null && !isNaN(value)) {
    const formatted = fmt ? fmt(value) : (Number.isInteger(value) ? value : (typeof value === 'number' ? value.toFixed(2) : value));
    el.textContent = formatted + ' ' + unit;
    if (colorFn) {
      const cls = colorFn(value);
      if (cls) el.className = 'metric-value ' + cls;
      else el.className = 'metric-value';
    }
  } else {
    el.textContent = '-';
    el.className = 'metric-value';
  }
}

function getMarketCapCategory(mcap) {
  if (mcap < 50) return { label: 'ナノキャップ', cls: 'mcap-nano' };
  if (mcap < 300) return { label: 'マイクロキャップ', cls: 'mcap-micro' };
  if (mcap < 1000) return { label: 'スモールキャップ', cls: 'mcap-small' };
  if (mcap < 5000) return { label: 'ミッドキャップ', cls: 'mcap-mid' };
  return { label: 'ラージキャップ', cls: 'mcap-large' };
}

function resetIndividual() {
  indData = {};
  document.getElementById('indCode').value = '';
  document.getElementById('indEdinet').value = '';
  document.getElementById('indResult').classList.add('hidden');
  document.getElementById('manual_sec_gain').value = '';
  document.getElementById('manual_prop_gain').value = '';
  document.getElementById('manual_land_gain').value = '';
  document.getElementById('manual_cross').value = '';
  document.getElementById('m_land_gain_method').style.display = 'none';
  ['m_price','m_pbr','m_per','m_roe','m_eps','m_bps','m_divyield','m_payout',
   'm_mcap','m_mcap_cat','m_netassets','m_sec_gain','m_prop_gain','m_land_bv','m_land_gain','m_total_gain',
   'm_adj_nav','m_adj_pbr','m_cash','m_debt','m_netcash','m_nc_ratio','m_adj_netcash','m_adj_nc_ratio',
   'm_full_netcash','m_full_nc_ratio','m_ev_ebitda',
   'm_equity_ratio','m_foreign','m_outside_dir','m_shares','m_treasury']
    .forEach(id => { const el = document.getElementById(id); if (el) el.textContent = '-'; });
  // 土地明細セクションをリセット
  const landSection = document.getElementById('landAnalysisSection');
  if (landSection) landSection.style.display = 'none';
  const landResult = document.getElementById('landResult');
  if (landResult) landResult.classList.add('hidden');
}

// ═══════════════════════════════════════════════════════════════
// 土地明細分析
// ═══════════════════════════════════════════════════════════════
let landParcelsData = null;

async function fetchLandParcels() {
  const apiKey = document.getElementById('edinetApiKey').value.trim();
  const edinetCode = document.getElementById('indEdinet').value.trim();
  const stockCode = document.getElementById('indCode').value.trim();
  const searchCode = edinetCode || stockCode;
  if (!searchCode) { alert('EDINETコードまたは証券コードを入力してください'); return; }

  const btn = document.getElementById('btnLandAnalysis');
  const loading = document.getElementById('landLoading');
  const loadingText = document.getElementById('landLoadingText');
  btn.disabled = true;
  loading.classList.add('active');

  try {
    // EDINET書類検索でdocIDを取得（EDINETコードまたは証券コード）
    const apiParam = apiKey ? '?apiKey=' + encodeURIComponent(apiKey) : '';
    loadingText.textContent = 'EDINET書類検索中...';
    const sRes = await fetch(API_BASE + '/api/edinet/search/' + searchCode + apiParam);
    const sData = await sRes.json();
    if (!sData.success || !sData.documents || sData.documents.length === 0) {
      alert('有価証券報告書が見つかりません');
      return;
    }
    const docID = sData.documents[0].docID;

    // 土地明細分析API呼び出し
    loadingText.textContent = '土地明細を解析中（固定資産明細表 + 地価公示データ照合）...';
    const lRes = await fetch(API_BASE + '/api/edinet/land-parcels/' + docID + apiParam);
    const lData = await lRes.json();

    if (!lData.success) {
      alert('土地明細の解析に失敗: ' + (lData.error || ''));
      return;
    }

    landParcelsData = lData.data;
    displayLandParcels(lData.data);

  } catch (err) {
    alert('エラー: ' + err.message);
  } finally {
    btn.disabled = false;
    loading.classList.remove('active');
  }
}

function displayLandParcels(data) {
  const result = document.getElementById('landResult');
  result.classList.remove('hidden');

  // サマリー
  document.getElementById('land_parcel_count').textContent = data.parcelCount + '件';
  document.getElementById('land_total_bv').textContent =
    data.totalBookValue > 0 ? Math.round(data.totalBookValue).toLocaleString() + ' 百万円' : '-';
  document.getElementById('land_total_fv').textContent =
    data.totalEstimatedValue > 0 ? Math.round(data.totalEstimatedValue).toLocaleString() + ' 百万円' : '-';

  const gainEl = document.getElementById('land_total_gain');
  if (data.totalEstimatedGain != null) {
    gainEl.textContent = Math.round(data.totalEstimatedGain).toLocaleString() + ' 百万円';
    gainEl.style.color = data.totalEstimatedGain > 0 ? 'var(--green)' : 'var(--danger)';
  } else {
    gainEl.textContent = '-';
  }

  const methodEl = document.getElementById('land_gain_method_detail');
  if (data.gainMethod) {
    methodEl.textContent = '推定方法: ' + data.gainMethod;
  }

  // 明細テーブル
  const tbody = document.getElementById('landParcelsBody');
  tbody.innerHTML = '';

  if (data.parcels && data.parcels.length > 0) {
    data.parcels.forEach((p, i) => {
      const gain = (p.estimatedValue && p.bookValue) ? p.estimatedValue - p.bookValue : null;
      const gainColor = gain != null ? (gain > 0 ? 'color:var(--green);font-weight:600;' : 'color:var(--danger);') : '';
      tbody.innerHTML += `<tr>
        <td>${i + 1}</td>
        <td style="max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${p.name || '-'}</td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${p.address || ''}">${p.address || '-'}</td>
        <td style="text-align:right;">${p.area ? p.area.toLocaleString() : '-'}</td>
        <td style="text-align:right;">${p.bookValue ? Math.round(p.bookValue).toLocaleString() : '-'}</td>
        <td style="text-align:right;">${p.estimatedPricePerSqm ? Math.round(p.estimatedPricePerSqm).toLocaleString() : '-'}</td>
        <td style="text-align:right;">${p.estimatedValue ? Math.round(p.estimatedValue).toLocaleString() : '-'}</td>
        <td style="text-align:right;${gainColor}">${gain != null ? Math.round(gain).toLocaleString() : '-'}</td>
      </tr>`;
    });

  } else {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-light);">土地明細が見つかりませんでした</td></tr>';
  }
}

function applyLandGainEstimate() {
  if (!landParcelsData || landParcelsData.totalEstimatedGain == null) {
    alert('推定値がありません');
    return;
  }
  document.getElementById('manual_land_gain').value = Math.round(landParcelsData.totalEstimatedGain);
  recalcIndividual();
  alert('土地含み益推定値 ' + Math.round(landParcelsData.totalEstimatedGain).toLocaleString() + ' 百万円を反映しました');
}

// ═══════════════════════════════════════════════════════════════
// 政策保有株式 - 株価取得・含み益計算
// ═══════════════════════════════════════════════════════════════
var phPriceData = null; // 株価取得結果を保持

function renderPolicyHoldingsTable(holdings, prices) {
  var body = document.getElementById('policyHoldingsBody');
  body.innerHTML = '';
  var basis = document.getElementById('phPriceBasis') ? document.getElementById('phPriceBasis').value : 'lastClose';

  holdings.forEach(function(h, i) {
    var tr = document.createElement('tr');
    var priceInfo = prices ? prices.find(function(p) { return p.name === h.name; }) : null;
    var price = priceInfo ? priceInfo[basis] : null;
    var currentVal = (price && h.shares) ? h.shares * price / 1000000 : null;
    var gain = currentVal != null ? currentVal - h.marketValue : null;

    tr.innerHTML = '<td>' + (i+1) + '</td>'
      + '<td>' + h.name + '</td>'
      + '<td>' + (priceInfo && priceInfo.ticker ? priceInfo.ticker : '-') + '</td>'
      + '<td style="text-align:right;">' + (h.shares ? Math.round(h.shares).toLocaleString() : '-') + '</td>'
      + '<td style="text-align:right;">' + Math.round(h.marketValue).toLocaleString() + '</td>'
      + '<td style="text-align:right;">' + (price ? price.toLocaleString(undefined, {maximumFractionDigits:1}) : '-') + '</td>'
      + '<td style="text-align:right;">' + (currentVal != null ? Math.round(currentVal).toLocaleString() : '-') + '</td>'
      + '<td style="text-align:right;color:' + (gain > 0 ? 'var(--green)' : gain < 0 ? '#c0392b' : '') + ';">'
        + (gain != null ? (gain > 0 ? '+' : '') + Math.round(gain).toLocaleString() : '-') + '</td>';
    body.appendChild(tr);
  });
}

function updatePolicyHoldingsSummary(holdings, prices) {
  var basis = document.getElementById('phPriceBasis').value;
  var totalCurrent = 0;
  var totalReport = 0;
  var counted = 0;

  holdings.forEach(function(h) {
    totalReport += h.marketValue;
    var priceInfo = prices ? prices.find(function(p) { return p.name === h.name; }) : null;
    var price = priceInfo ? priceInfo[basis] : null;
    if (price && h.shares) {
      totalCurrent += h.shares * price / 1000000;
      counted++;
    }
  });

  document.getElementById('ph_current_total').textContent = counted > 0 ? Math.round(totalCurrent).toLocaleString() + ' 百万円' : '-';
  var gainEl = document.getElementById('ph_gain');
  if (counted > 0) {
    var gain = totalCurrent - totalReport;
    gainEl.textContent = (gain > 0 ? '+' : '') + Math.round(gain).toLocaleString() + ' 百万円';
    gainEl.style.color = gain > 0 ? 'var(--green)' : '#c0392b';
  } else {
    gainEl.textContent = '-';
    gainEl.style.color = '';
  }
}

async function fetchPolicyHoldingsPrices() {
  var e = indData.edinet || {};
  if (!e.policyHoldingsTop || e.policyHoldingsTop.length === 0) {
    alert('政策保有株式データがありません');
    return;
  }

  var loading = document.getElementById('phLoading');
  var btn = document.getElementById('btnFetchStockPrices');
  loading.style.display = '';
  btn.disabled = true;

  try {
    var resp = await fetch(API_BASE + '/api/stock-prices', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ holdings: e.policyHoldingsTop })
    });
    var data = await resp.json();
    if (!data.success) throw new Error(data.error || 'API error');

    phPriceData = data.results;
    renderPolicyHoldingsTable(e.policyHoldingsTop, phPriceData);
    updatePolicyHoldingsSummary(e.policyHoldingsTop, phPriceData);
  } catch (err) {
    alert('株価取得エラー: ' + err.message);
  } finally {
    loading.style.display = 'none';
    btn.disabled = false;
  }
}

function applyPolicyHoldingsGain() {
  var e = indData.edinet || {};
  if (!e.policyHoldingsTop) {
    alert('政策保有株式データがありません');
    return;
  }

  // 株価取得済みの場合は現在時価ベースで含み益を計算
  if (phPriceData) {
    var basis = document.getElementById('phPriceBasis').value;
    var totalCurrent = 0;
    var totalReport = 0;
    e.policyHoldingsTop.forEach(function(h) {
      totalReport += h.marketValue;
      var priceInfo = phPriceData.find(function(p) { return p.name === h.name; });
      var price = priceInfo ? priceInfo[basis] : null;
      if (price && h.shares) {
        totalCurrent += h.shares * price / 1000000;
      } else {
        totalCurrent += h.marketValue; // 株価不明分は報告書額を使用
      }
    });
    var gain = totalCurrent - totalReport;
    document.getElementById('manual_sec_gain').value = Math.round(gain);
    recalcIndividual();
    alert('推定含み益 ' + (gain > 0 ? '+' : '') + Math.round(gain).toLocaleString() + ' 百万円を反映しました（' + basis + '基準）');
  } else {
    // 株価未取得の場合は報告書の時価合計をそのまま反映
    document.getElementById('manual_sec_gain').value = Math.round(e.policyHoldingsMarketValue);
    recalcIndividual();
    alert('報告書時価合計 ' + Math.round(e.policyHoldingsMarketValue).toLocaleString() + ' 百万円を反映しました');
  }
}

// Google Maps連携は将来の有料化時に追加予定

// ═══════════════════════════════════════════════════════════════
// ランキング・スクリーニング
// ═══════════════════════════════════════════════════════════════
let rankResults = [];
let scanCancelled = false;

const PRESETS = {
  mega: ['8306', '8316', '8411', '8308', '8309', '8601', '8604', '8630', '8725', '8766'],
  lowpbr: ['5411', '5401', '7011', '7012', '6301', '6302', '5020', '5019', '5021', '8058',
           '8053', '3402', '3401', '4183', '4188', '5332', '5333', '3861', '3863', '4042'],
  auto: ['7203', '7267', '7201', '7270', '7269', '7211', '7202', '7272', '5108', '6902'],
  trading: ['8058', '8031', '8001', '8002', '8053', '8015', '8020'],
  it: ['9984', '9433', '9432', '9434', '4689', '4755', '3382', '6758', '6861', '6954'],
  conglomerate: ['6501', '6502', '6503', '7751', '6752', '6753', '6702', '6701', '8035', '7731'],
  // 追加プリセット
  realestate: ['8801', '8802', '8804', '8830', '3289', '3291', '3462', '8818', '8848', '3003'],
  food: ['2502', '2503', '2587', '2801', '2802', '2871', '2897', '2914', '2269', '2809'],
  pharma: ['4502', '4503', '4506', '4507', '4519', '4523', '4568', '4578', '4151', '4452'],
  chemical: ['4063', '4188', '4183', '4005', '4004', '4021', '4042', '4208', '4631', '4901'],
  steel: ['5401', '5411', '5406', '5423', '5444', '5471', '5480', '5481', '5486', '5801'],
  construction: ['1801', '1802', '1803', '1812', '1820', '1821', '1878', '1925', '1928', '1963'],
  retail: ['3382', '8267', '8252', '9983', '7532', '3086', '3099', '2651', '7453', '2670'],
  transport: ['9020', '9021', '9022', '9001', '9005', '9007', '9064', '9101', '9104', '9107'],
  utility: ['9501', '9502', '9503', '9504', '9505', '9506', '9507', '9508', '9509', '9531'],
  machinery: ['6301', '6302', '6305', '6361', '6367', '6471', '6473', '7004', '7011', '7012'],
  electronics: ['6501', '6502', '6503', '6701', '6702', '6752', '6758', '6861', '6954', '6971'],
  nikkei225: ['7203', '6758', '9984', '8306', '4502', '6861', '6954', '9433', '8058', '6501',
              '7267', '6367', '4063', '6902', '7741', '8316', '4503', '6762', '7751', '9432'],
};

function loadPreset(key) {
  document.getElementById('rankCodes').value = PRESETS[key].join(', ');
}

async function startRankingScan() {
  const text = document.getElementById('rankCodes').value.trim();
  if (!text) { alert('銘柄コードを入力してください'); return; }

  const codes = text.split(/[,\s\n]+/).map(c => c.trim().toUpperCase()).filter(c => /^[0-9A-Za-z]{4}$/.test(c));
  if (codes.length === 0) { alert('有効な4桁コードが見つかりません（例: 7203）'); return; }
  if (codes.length > 100) { alert('最大100銘柄まで'); return; }

  scanCancelled = false;
  rankResults = [];
  const btn = document.getElementById('btnRankScan');
  const loading = document.getElementById('rankLoading');
  const progress = document.getElementById('rankProgress');
  const loadingText = document.getElementById('rankLoadingText');

  btn.disabled = true;
  loading.classList.add('active');
  progress.classList.remove('hidden');
  document.getElementById('rankResult').classList.add('hidden');

  // バッチリクエスト
  loadingText.textContent = `${codes.length}銘柄をスキャン中...`;
  document.getElementById('rankProgressBar').style.width = '0%';
  document.getElementById('rankProgressText').textContent = `0 / ${codes.length}`;

  try {
    const res = await fetch(API_BASE+'/api/batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ codes })
    });
    const data = await res.json();

    if (data.success) {
      // スコア計算
      rankResults = data.data
        .filter(d => d.companyName && !d.error)
        .map(d => {
          const score = calcQuickScore(d);
          return { ...d, score };
        });

      document.getElementById('rankProgressBar').style.width = '100%';
      document.getElementById('rankProgressText').textContent = `${codes.length} / ${codes.length}`;

      // フィルタ & ソート & 表示
      filterAndDisplayRanking();
    }
  } catch (err) {
    alert('スキャンエラー: ' + err.message);
  } finally {
    btn.disabled = false;
    loading.classList.remove('active');
    setTimeout(() => progress.classList.add('hidden'), 1000);
  }
}

function cancelScan() {
  scanCancelled = true;
}

// 簡易スコア（Yahoo Financeデータのみ）
function calcQuickScore(d) {
  let score = 0;

  // PBR (max 20)
  if (d.pbr != null) {
    score += d.pbr < 0.5 ? 20 : d.pbr < 0.7 ? 16 : d.pbr < 0.8 ? 13 : d.pbr < 1.0 ? 10 : d.pbr < 1.3 ? 5 : 2;
  }

  // ROE (max 15)
  if (d.roe != null) {
    score += d.roe < 3 ? 15 : d.roe < 5 ? 12 : d.roe < 8 ? 8 : d.roe < 10 ? 4 : 1;
  }

  // 配当性向 (max 12)
  if (d.payoutRatio != null) {
    score += d.payoutRatio < 20 ? 12 : d.payoutRatio < 30 ? 9 : d.payoutRatio < 40 ? 6 : d.payoutRatio < 50 ? 3 : 1;
  }

  // 自己資本比率 (max 12)
  if (d.equityRatio != null) {
    score += d.equityRatio > 80 ? 12 : d.equityRatio > 70 ? 9 : d.equityRatio > 60 ? 6 : d.equityRatio > 50 ? 3 : 1;
  }

  // 配当利回り → 低すぎるとスコアUP (max 10)
  if (d.dividendYield != null) {
    score += d.dividendYield < 1.5 ? 10 : d.dividendYield < 2.5 ? 7 : d.dividendYield < 3.5 ? 4 : 2;
  }

  // 時価総額 (max 10) - sweet spot: 100-5000億
  if (d.marketCapOku != null) {
    score += d.marketCapOku < 100 ? 4 : d.marketCapOku < 500 ? 10 : d.marketCapOku < 1000 ? 9 :
             d.marketCapOku < 3000 ? 7 : d.marketCapOku < 5000 ? 5 : 3;
  }

  // PBR < 1.0のボーナス (max 10)
  if (d.pbr != null && d.pbr < 1.0) {
    // BPS対比の割安度
    const discount = (1 - d.pbr) * 10;
    score += Math.min(10, Math.round(discount));
  }

  return Math.min(100, score);
}

// ビュー切り替え
function switchRankView(view) {
  document.querySelectorAll('.rank-view-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.rank-view').forEach(v => v.classList.remove('active'));
  document.querySelector(`.rank-view-tab[onclick*="${view}"]`).classList.add('active');
  document.getElementById('rankView-' + view).classList.add('active');
}

function filterAndDisplayRanking() {
  const filter = document.getElementById('rankMcapFilter').value;
  const sortBy = document.getElementById('rankSortBy').value;

  let filtered = [...rankResults];

  // 時価総額フィルタ
  if (filter !== 'all') {
    filtered = filtered.filter(d => {
      if (d.marketCapOku == null) return false;
      switch (filter) {
        case 'nano': return d.marketCapOku < 50;
        case 'micro': return d.marketCapOku >= 50 && d.marketCapOku < 300;
        case 'small': return d.marketCapOku >= 300 && d.marketCapOku < 1000;
        case 'mid': return d.marketCapOku >= 1000 && d.marketCapOku < 5000;
        case 'large': return d.marketCapOku >= 5000;
        default: return true;
      }
    });
  }

  // ソート
  filtered.sort((a, b) => {
    switch (sortBy) {
      case 'score': return (b.score || 0) - (a.score || 0);
      case 'pbr': return (a.pbr || 999) - (b.pbr || 999);
      case 'roe': return (a.roe || 999) - (b.roe || 999);
      case 'divyield': return (b.dividendYield || 0) - (a.dividendYield || 0);
      case 'equity': return (b.equityRatio || 0) - (a.equityRatio || 0);
      case 'per': return (a.per || 999) - (b.per || 999);
      case 'payout': return (a.payoutRatio || 999) - (b.payoutRatio || 999);
      case 'mcap': return (b.marketCapOku || 0) - (a.marketCapOku || 0);
      default: return (b.score || 0) - (a.score || 0);
    }
  });

  // ── 概要ビュー ──
  const tbody = document.getElementById('rankTableBody');
  tbody.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const cat = d.marketCapOku != null ? getMarketCapCategory(d.marketCapOku) : null;
    const pbrCls = d.pbr != null && d.pbr < 1.0 ? 'style="color:var(--danger);font-weight:700;"' : '';
    const roeCls = d.roe != null && d.roe < 8 ? 'style="color:var(--orange);"' : '';
    tbody.innerHTML += `<tr>
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td class="name-cell" style="max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.7rem;color:var(--text-light);">${d.sector || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${pbrCls}>${d.pbr != null ? d.pbr.toFixed(2) : '-'}</td>
      <td ${roeCls}>${d.roe != null ? d.roe.toFixed(1) + '%' : '-'}</td>
      <td>${d.per != null ? d.per.toFixed(1) : '-'}</td>
      <td>${d.dividendYield != null ? d.dividendYield.toFixed(2) + '%' : '-'}</td>
      <td>${d.equityRatio != null ? d.equityRatio.toFixed(1) + '%' : '-'}</td>
      <td>${d.marketCapOku != null ? d.marketCapOku.toLocaleString() + '億' : '-'}</td>
      <td>${cat ? `<span class="mcap-cat ${cat.cls}">${cat.label}</span>` : '-'}</td>
    </tr>`;
  });

  // ── バリュエーションビュー ──
  const tbodyVal = document.getElementById('rankTableBody-valuation');
  tbodyVal.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const pbrCls = d.pbr != null && d.pbr < 1.0 ? 'style="color:var(--danger);font-weight:700;"' : '';
    const roeCls = d.roe != null && d.roe < 8 ? 'style="color:var(--orange);"' : '';
    tbodyVal.innerHTML += `<tr>
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${pbrCls}>${d.pbr != null ? d.pbr.toFixed(2) : '-'}</td>
      <td>${d.per != null ? d.per.toFixed(1) : '-'}</td>
      <td ${roeCls}>${d.roe != null ? d.roe.toFixed(1) + '%' : '-'}</td>
      <td>${d.eps != null ? d.eps.toFixed(1) : '-'}</td>
      <td>${d.bps != null ? d.bps.toLocaleString() : '-'}</td>
      <td>${d.marketCapOku != null ? d.marketCapOku.toLocaleString() + '億' : '-'}</td>
    </tr>`;
  });

  // ── 配当・還元ビュー ──
  const tbodyDiv = document.getElementById('rankTableBody-dividend');
  tbodyDiv.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const yieldCls = d.dividendYield != null && d.dividendYield >= 3 ? 'style="color:var(--green);font-weight:700;"' : '';
    const payoutCls = d.payoutRatio != null && d.payoutRatio < 30 ? 'style="color:var(--orange);"' : '';
    tbodyDiv.innerHTML += `<tr>
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${yieldCls}>${d.dividendYield != null ? d.dividendYield.toFixed(2) + '%' : '-'}</td>
      <td ${payoutCls}>${d.payoutRatio != null ? d.payoutRatio.toFixed(1) + '%' : '-'}</td>
      <td>${d.dps != null ? d.dps.toFixed(1) + '円' : '-'}</td>
      <td>${d.pbr != null ? d.pbr.toFixed(2) : '-'}</td>
      <td>${d.marketCapOku != null ? d.marketCapOku.toLocaleString() + '億' : '-'}</td>
    </tr>`;
  });

  // ── 財務健全性ビュー ──
  const tbodyFin = document.getElementById('rankTableBody-financial');
  tbodyFin.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const eqCls = d.equityRatio != null && d.equityRatio > 60 ? 'style="color:var(--green);font-weight:700;"' : '';
    const roeCls = d.roe != null && d.roe < 8 ? 'style="color:var(--orange);"' : '';
    const cat = d.marketCapOku != null ? getMarketCapCategory(d.marketCapOku) : null;
    tbodyFin.innerHTML += `<tr>
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${eqCls}>${d.equityRatio != null ? d.equityRatio.toFixed(1) + '%' : '-'}</td>
      <td ${roeCls}>${d.roe != null ? d.roe.toFixed(1) + '%' : '-'}</td>
      <td>${d.per != null ? d.per.toFixed(1) : '-'}</td>
      <td>${d.payoutRatio != null ? d.payoutRatio.toFixed(1) + '%' : '-'}</td>
      <td>${d.marketCapOku != null ? d.marketCapOku.toLocaleString() + '億' : '-'}</td>
      <td>${cat ? `<span class="mcap-cat ${cat.cls}">${cat.label}</span>` : '-'}</td>
    </tr>`;
  });

  // ── 含み益・資産ビュー ──
  const tbodyAssets = document.getElementById('rankTableBody-assets');
  tbodyAssets.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const pbrCls = d.pbr != null && d.pbr < 1.0 ? 'style="color:var(--danger);font-weight:700;"' : '';
    const totalGain = (d.estimatedLandGain || 0) + (d.securitiesGain || 0) + (d.investmentPropertyGain || 0);
    const hasData = d.hasEdinetData;
    const fmtM = v => v != null ? Math.round(v).toLocaleString() : '-';
    const gainColor = v => v != null && v > 0 ? 'style="color:var(--green);font-weight:600;"' : v != null && v < 0 ? 'style="color:var(--danger);"' : '';

    tbodyAssets.innerHTML += `<tr style="${!hasData ? 'opacity:0.5;' : ''}">
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${pbrCls}>${d.pbr != null ? d.pbr.toFixed(2) : '-'}</td>
      <td>${d.land != null ? fmtM(d.land) : '-'}</td>
      <td ${gainColor(d.estimatedLandGain)}>${d.estimatedLandGain != null ? fmtM(d.estimatedLandGain) : '-'}</td>
      <td ${gainColor(d.securitiesGain)}>${d.securitiesGain != null ? fmtM(d.securitiesGain) : '-'}</td>
      <td ${gainColor(d.investmentPropertyGain)}>${d.investmentPropertyGain != null ? fmtM(d.investmentPropertyGain) : '-'}</td>
      <td ${gainColor(totalGain > 0 ? totalGain : null)}>${hasData ? fmtM(totalGain) : '-'}</td>
      <td>${d.marketCapOku != null ? d.marketCapOku.toLocaleString() + '億' : '-'}</td>
    </tr>`;
  });

  document.getElementById('rankResultCount').textContent = filtered.length + '件';
  document.getElementById('rankResult').classList.remove('hidden');
}

// ソート列クリック
let currentSortCol = 'score';
let currentSortAsc = false;

function sortRanking(col) {
  if (currentSortCol === col) {
    currentSortAsc = !currentSortAsc;
  } else {
    currentSortCol = col;
    currentSortAsc = col === 'name' || col === 'code';
  }

  rankResults.sort((a, b) => {
    let va, vb;
    switch (col) {
      case 'rank': case 'score': va = a.score || 0; vb = b.score || 0; break;
      case 'code': va = a.code; vb = b.code; break;
      case 'name': va = a.companyName || ''; vb = b.companyName || ''; break;
      case 'pbr': va = a.pbr || 999; vb = b.pbr || 999; break;
      case 'roe': va = a.roe || 999; vb = b.roe || 999; break;
      case 'per': va = a.per || 999; vb = b.per || 999; break;
      case 'divyield': va = a.dividendYield || 0; vb = b.dividendYield || 0; break;
      case 'payout': va = a.payoutRatio || 999; vb = b.payoutRatio || 999; break;
      case 'equity': va = a.equityRatio || 0; vb = b.equityRatio || 0; break;
      case 'mcap': va = a.marketCapOku || 0; vb = b.marketCapOku || 0; break;
      case 'mcapCat': va = a.marketCapOku || 0; vb = b.marketCapOku || 0; break;
      case 'land': va = a.land || 0; vb = b.land || 0; break;
      case 'landGain': va = a.estimatedLandGain || 0; vb = b.estimatedLandGain || 0; break;
      case 'secGain': va = a.securitiesGain || 0; vb = b.securitiesGain || 0; break;
      case 'ipGain': va = a.investmentPropertyGain || 0; vb = b.investmentPropertyGain || 0; break;
      case 'totalGain': va = (a.estimatedLandGain||0)+(a.securitiesGain||0)+(a.investmentPropertyGain||0); vb = (b.estimatedLandGain||0)+(b.securitiesGain||0)+(b.investmentPropertyGain||0); break;
      case 'sector': va = a.sector || ''; vb = b.sector || ''; break;
      default: va = a.score || 0; vb = b.score || 0;
    }
    if (typeof va === 'string') return currentSortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    return currentSortAsc ? va - vb : vb - va;
  });

  filterAndDisplayRanking();
}

// フィルタ変更時の再表示
document.getElementById('rankMcapFilter').addEventListener('change', filterAndDisplayRanking);
document.getElementById('rankSortBy').addEventListener('change', filterAndDisplayRanking);

// ── localStorage: APIキー保存 ──
window.addEventListener('load', () => {
  const savedKey = localStorage.getItem('edinetApiKey');
  if (savedKey) document.getElementById('edinetApiKey').value = savedKey;
});
document.getElementById('edinetApiKey').addEventListener('change', (e) => {
  localStorage.setItem('edinetApiKey', e.target.value);
});
