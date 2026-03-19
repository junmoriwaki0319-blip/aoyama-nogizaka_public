/* Mobile Nav */
function toggleMenu(){document.getElementById("mobileMenu").classList.toggle("open")}

/* API BASE */
var API_BASE=(function(){var h=window.location.hostname;
if(h.includes("vercel.app"))return "";
if(h.includes("aoyama-nogizaka.com")||h.includes("github.io"))return "https://aoyama-nogizakapublic.vercel.app";
if(h==="localhost"||h==="127.0.0.1")return "";
return "https://aoyama-nogizakapublic.vercel.app"})();

/* Auth UI — 初期化時にFirebase authが既に解決済みなら即反映 */
if (window.currentUser !== undefined) { document.addEventListener('DOMContentLoaded', function(){ if(typeof updateAuthUI==='function') updateAuthUI(window.currentUser); }); }
function updateAuthUI(user){
  var lo=document.getElementById("authLoggedOut"),li=document.getElementById("authLoggedIn"),cb=document.getElementById("btnCsvExport");
  var gated=document.getElementById("screenerGated"),gatedR=document.getElementById("screenerGatedRank");
  var lwI=document.getElementById("loginWallIndividual"),lwR=document.getElementById("loginWallRanking");
  var navAuth=document.getElementById("navAuthBtns"),navUser=document.getElementById("navUserInfo"),navName=document.getElementById("navUserName");
  var mobAuth=document.getElementById("mobileAuthBtns"),mobUser=document.getElementById("mobileUserInfo"),mobName=document.getElementById("mobileUserName");
  if(user){
    var dn=(user.displayName||user.email)+' 様';
    lo.classList.add("hidden");li.classList.remove("hidden");document.getElementById("authUserName").textContent=user.displayName||user.email;if(cb)cb.disabled=false;
    if(gated)gated.classList.remove("blurred");if(gatedR)gatedR.classList.remove("blurred");
    if(lwI)lwI.style.display="none";if(lwR)lwR.style.display="none";
    if(navAuth)navAuth.style.display="none";if(navUser){navUser.style.display="flex";if(navName)navName.textContent=dn;}
    if(mobAuth)mobAuth.style.display="none";if(mobUser){mobUser.style.display="flex";if(mobName)mobName.textContent=dn;}
  }else{
    lo.classList.remove("hidden");li.classList.add("hidden");if(cb)cb.disabled=true;
    if(gated)gated.classList.add("blurred");if(gatedR)gatedR.classList.add("blurred");
    if(lwI)lwI.style.display="block";if(lwR)lwR.style.display="block";
    if(navAuth)navAuth.style.display="flex";if(navUser)navUser.style.display="none";
    if(mobAuth)mobAuth.style.display="flex";if(mobUser)mobUser.style.display="none";
  }
}
function showRegister(){document.getElementById("authLoginForm").classList.add("hidden");var rf=document.getElementById("authResetForm");if(rf)rf.classList.add("hidden");document.getElementById("authRegisterForm").classList.remove("hidden")}
function scrollToAuth(mode){if(mode==='register')showRegister();else showLogin();document.getElementById("authSection").scrollIntoView({behavior:"smooth",block:"center"})}
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
async function handleResendVerify(){var m=document.getElementById("myPageMsg");try{var mod=await import("https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js");await mod.sendEmailVerification(window.currentUser);m.innerHTML='認証メールを再送しました。<br><span style="font-size:11px;">届かない場合は迷惑メールフォルダをご確認いただくか、<a href="https://aoyama-nogizaka.com/#contact" style="color:#047857;text-decoration:underline;">お問い合わせフォーム</a>よりご連絡ください。</span>';m.style.display="block";m.style.background="#d1fae5";m.style.color="#065f46"}catch(e){m.textContent=e.code==="auth/too-many-requests"?"送信回数が多すぎます。しばらくしてからお試しください。":"エラー: "+e.message;m.style.display="block";m.style.background="#fee2e2";m.style.color="#991b1b"}}


/* CSV Export (Member Only) */
function exportCSV(){
  if(!window.currentUser){alert("CSV出力は会員限定機能です。ログインしてください。");return}
  if(rankResults.length===0){alert("データがありません");return}
  var hd=["順位","コード","銘柄名","スコア","PBR","PER","ROE(%)","配当利回り(%)","配当性向(%)","自己資本比率(%)","時価総額(億円)","市場","セクター","NC/時価総額(%)","実質NC/時価総額(%)","含み資産NC/時価総額(%)","EV/EBITDA","土地簿価(百万円)","推定土地含み益(百万円)","有価証券含み益(百万円)","投資不動産含み益(百万円)","含み益合計(百万円)"];
  var sr=rankResults.slice().sort(function(a,b){return(b.score||0)-(a.score||0)});
  var rows=sr.map(function(d,i){var tg=(d.estimatedLandGain||0)+(d.securitiesGain||0)+(d.investmentPropertyGain||0);return[i+1,d.code,d.companyName||"",d.score||"",d.pbr!=null?d.pbr.toFixed(2):"",d.per!=null?d.per.toFixed(1):"",d.roe!=null?d.roe.toFixed(1):"",d.dividendYield!=null?d.dividendYield.toFixed(2):"",d.payoutRatio!=null?d.payoutRatio.toFixed(1):"",d.equityRatio!=null?d.equityRatio.toFixed(1):"",d.marketCapOku!=null?d.marketCapOku:"",d.market||"",d.sector||"",d.ncRatio!=null?d.ncRatio.toFixed(1):"",d.adjNcRatio!=null?d.adjNcRatio.toFixed(1):"",d.fullNcRatio!=null?d.fullNcRatio.toFixed(1):"",d.evEbitda!=null?d.evEbitda.toFixed(1):"",d.land!=null?d.land:"",d.estimatedLandGain!=null?d.estimatedLandGain:"",d.securitiesGain!=null?d.securitiesGain:"",d.investmentPropertyGain!=null?d.investmentPropertyGain:"",d.hasEdinetData?tg:""]});
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
  if (!window.currentUser) { alert('この機能を利用するには無料会員登録が必要です。ページ下部からログインまたは新規登録してください。'); return; }
  const code = document.getElementById('indCode').value.trim().toUpperCase();
  if (!/^[0-9A-Za-z]{4}$/.test(code)) { alert('4桁の証券コードを入力してください（例: 7203）'); return; }

  const btn = document.getElementById('btnIndFetch');
  const loading = document.getElementById('indLoading');
  const loadingText = document.getElementById('indLoadingText');
  btn.disabled = true;
  loading.classList.add('active');

  // 前回の分析結果をクリア（土地明細・政策保有株式・株主構成の残留データを防止）
  resetIndividual();
  indData = { code };
  document.getElementById('indCode').value = code;
  document.getElementById('indEdinet').value = '';

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

    // 初回計算（結果パネルはまだ表示しない）
    calculateAndDisplay(true); // silent=true: 結果表示を抑制

    // === 全自動分析（土地+政策保有株式を並列実行）===
    loadingText.textContent = '詳細分析中（土地・政策保有株式・株主構成）...';
    await runAutoAnalysis();

    // 全分析完了 → 最終再計算 & 一括表示
    recalcIndividual();
    displayMajorShareholders();
    document.getElementById('indResult').classList.remove('hidden');
    document.getElementById('indResult').scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (err) {
    alert('エラー: ' + err.message);
  } finally {
    btn.disabled = false;
    loading.classList.remove('active');
  }
}

// 全自動分析チェーン（土地+政策保有株式を並列実行し、完了後に一括表示）
async function runAutoAnalysis() {
  const e = indData.edinet || {};
  const apiKey = document.getElementById('edinetApiKey').value.trim();
  const edinetCode = document.getElementById('indEdinet').value.trim();
  const stockCode = document.getElementById('indCode').value.trim();
  const searchCode = edinetCode || stockCode;
  const apiParam = apiKey ? '?apiKey=' + encodeURIComponent(apiKey) : '';

  const tasks = [];

  // セクションを明示的に非表示（resetIndividual後に残るケースの防止）
  var ls = document.getElementById('landAnalysisSection'); if (ls) ls.style.display = 'none';
  var lr = document.getElementById('landResult'); if (lr) lr.classList.add('hidden');
  var ps = document.getElementById('policyHoldingsSection'); if (ps) ps.style.display = 'none';

  // タスク1: 土地明細分析
  if (e.land && e.land > 0) {
    tasks.push((async () => {
      try {
        const sRes = await fetch(API_BASE + '/api/edinet/search/' + searchCode + apiParam);
        const sData = await sRes.json();
        if (sData.success && sData.documents && sData.documents.length > 0) {
          const docID = sData.documents[0].docID;
          const lRes = await fetch(API_BASE + '/api/edinet/land-parcels/' + docID + apiParam);
          const lData = await lRes.json();
          if (lData.success && lData.data) {
            landParcelsData = lData.data;
            document.getElementById('landAnalysisSection').style.display = '';
            // 拠点が検出された場合のみ明細を表示
            if (lData.data.parcelCount > 0) {
              displayLandParcels(lData.data);
              if (lData.data.totalEstimatedGain != null) {
                document.getElementById('manual_land_gain').value = Math.round(lData.data.totalEstimatedGain);
              }
            } else {
              // 不動産業等で主要な設備に土地明細がない場合
              var landResult = document.getElementById('landResult');
              if (landResult) {
                landResult.classList.remove('hidden');
                document.getElementById('land_parcel_count').textContent = '0件';
                var landBody = document.getElementById('landParcelsBody');
                if (landBody) landBody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-light);font-size:.72rem;">「主要な設備の状況」に土地明細がありません。不動産業の場合、保有物件は「賃貸等不動産」注記に開示されるため、土地明細分析の対象外となります。</td></tr>';
              }
            }
          }
        }
      } catch (err) { console.warn('土地自動分析エラー:', err); }
    })());
  }

  // タスク2: 政策保有株式の株価取得
  if (e.policyHoldingsTop && e.policyHoldingsTop.length > 0) {
    tasks.push((async () => {
      try {
        document.getElementById('policyHoldingsSection').style.display = '';
        const resp = await fetch(API_BASE + '/api/stock-prices', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ holdings: e.policyHoldingsTop })
        });
        const data = await resp.json();
        if (data.success && data.results) {
          phPriceData = data.results;
          renderPolicyHoldingsTable(e.policyHoldingsTop, phPriceData);
          updatePolicyHoldingsSummary(e.policyHoldingsTop, phPriceData);
          var basis = document.getElementById('phPriceBasis').value;
          var totalCurrent = 0, totalReport = 0;
          e.policyHoldingsTop.forEach(function(h) {
            totalReport += h.marketValue;
            var priceInfo = phPriceData.find(function(p) { return p.name === h.name; });
            var price = priceInfo ? priceInfo[basis] : null;
            if (price && h.shares) { totalCurrent += h.shares * price / 1000000; }
            else { totalCurrent += h.marketValue; }
          });
          document.getElementById('manual_sec_gain').value = Math.round(totalCurrent - totalReport);
        }
      } catch (err) { console.warn('政策保有株式自動分析エラー:', err); }
    })());
  }

  // 全タスク完了を待つ
  await Promise.all(tasks);
}

function calculateAndDisplay(silent) {
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
  if (mcapCatEl) {
    if (d.marketCapOku != null) {
      const cat = getMarketCapCategory(d.marketCapOku);
      mcapCatEl.innerHTML = `<span class="mcap-cat ${cat.cls}">${cat.label}</span>`;
    } else {
      mcapCatEl.textContent = '-';
    }
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
    var phc = document.getElementById('ph_count'); if (phc) phc.textContent = e.policyHoldingsCount;
    var phr = document.getElementById('ph_report_total'); if (phr) phr.textContent = Math.round(e.policyHoldingsMarketValue).toLocaleString() + ' 百万円';
    var phct = document.getElementById('ph_current_total'); if (phct) phct.textContent = '-';
    var phg = document.getElementById('ph_gain'); if (phg) phg.textContent = '-';
    renderPolicyHoldingsTable(e.policyHoldingsTop, null);
  } else if (phSection) {
    phSection.style.display = 'none';
  }

  // 賃貸等不動産（簿価・時価・含み益）
  setMetric('m_prop_bv', e.investmentPropertyBookValue, '百万円', v => Math.round(v).toLocaleString());
  setMetric('m_prop_fv', e.investmentPropertyFairValue, '百万円', v => Math.round(v).toLocaleString());
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
  var elCompany = document.getElementById('indCompanyName');
  var elMarket = document.getElementById('indMarketInfo');
  if (elCompany) elCompany.textContent = d.companyName || d.code;
  if (elMarket) elMarket.textContent = [d.code, d.market, d.sector].filter(Boolean).join(' | ');

  // スコア計算
  calculateScore();

  // 結果表示（silent=trueの場合は表示をスキップ、後で一括表示）
  if (!silent) {
    document.getElementById('indResult').classList.remove('hidden');
    document.getElementById('indResult').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
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
  if (!tbody) return;
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
  var scoreEl = document.getElementById('indTotalScore');
  if (scoreEl) scoreEl.textContent = total;

  const badge = document.getElementById('indScoreBadge');
  if (badge) {
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
}

// ── ユーティリティ ──
function setMetric(id, value, unit, fmt, colorFn) {
  const el = document.getElementById(id);
  if (!el) return;
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
   'm_mcap','m_mcap_cat','m_netassets','m_sec_gain','m_prop_bv','m_prop_fv','m_prop_gain','m_land_bv','m_land_gain','m_total_gain',
   'm_adj_nav','m_adj_pbr','m_cash','m_debt','m_netcash','m_nc_ratio','m_adj_netcash','m_adj_nc_ratio',
   'm_full_netcash','m_full_nc_ratio','m_ev_ebitda',
   'm_equity_ratio','m_foreign','m_outside_dir','m_shares','m_treasury']
    .forEach(id => { const el = document.getElementById(id); if (el) el.textContent = '-'; });
  // 土地明細セクションをリセット（内部データもクリア）
  const landSection = document.getElementById('landAnalysisSection');
  if (landSection) landSection.style.display = 'none';
  const landResult = document.getElementById('landResult');
  if (landResult) landResult.classList.add('hidden');
  landParcelsData = null;
  ['land_parcel_count','land_total_bv','land_total_fv','land_total_gain'].forEach(id => {
    var el = document.getElementById(id); if (el) el.textContent = '-';
  });
  var landMethod = document.getElementById('land_gain_method_detail');
  if (landMethod) landMethod.textContent = '';
  var landBody = document.getElementById('landParcelsBody');
  if (landBody) landBody.innerHTML = '';

  // 政策保有株式セクションをリセット（内部データもクリア）
  const phSection = document.getElementById('policyHoldingsSection');
  if (phSection) phSection.style.display = 'none';
  phPriceData = null;
  ['ph_count','ph_report_total','ph_current_total','ph_gain'].forEach(id => {
    var el = document.getElementById(id); if (el) el.textContent = '-';
  });
  var phBody = document.getElementById('policyHoldingsBody');
  if (phBody) phBody.innerHTML = '';

  // 株主構成セクションをリセット（内部データもクリア）
  const shSection = document.getElementById('shareholderSection');
  if (shSection) shSection.style.display = 'none';
  ['sh_count','sh_total_ratio','sh_float_ratio'].forEach(id => {
    var el = document.getElementById(id); if (el) el.textContent = '-';
  });
  var shChart = document.getElementById('shCategoryChart');
  if (shChart) shChart.innerHTML = '';
  var shBody = document.getElementById('shareholderBody');
  if (shBody) shBody.innerHTML = '';
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

// 用途推定キーワード（工場を先に判定 → 「本社工場」で工場が優先される）
var LAND_USE_RULES = [
  { key: 'factory',     label: '工場・倉庫', keywords: ['工場','製造所','製作所','倉庫','物流','配送','プラント','製鉄所','製油所','精製所'] },
  { key: 'commercial',  label: '商業施設',   keywords: ['店舗','ショッピング','商業','モール','販売','百貨店','SC'] },
  { key: 'residential', label: '住宅・寮',   keywords: ['社宅','寮','住宅','マンション','レジデンス'] },
  { key: 'idle',        label: '遊休地',     keywords: ['遊休','未利用','跡地'] },
  { key: 'office',      label: 'オフィス',   keywords: ['本社','本店','事務所','オフィス','事業所','支店','支社','営業所'] },
];

function guessLandUse(name, address, area) {
  // 面積10万㎡超は工場・大規模施設とみなす
  if (area && area > 100000) {
    return { key: 'factory', label: '工場・倉庫', keywords: [] };
  }
  var text = ((name || '') + ' ' + (address || '')).toLowerCase();
  for (var r of LAND_USE_RULES) {
    for (var kw of r.keywords) {
      if (text.includes(kw)) return r;
    }
  }
  return { key: 'other', label: 'その他', keywords: [] };
}

function getDefaultRate(useKey) {
  var el = document.getElementById('defRate_' + useKey);
  return el ? parseFloat(el.value) || 0.5 : 0.5;
}

function displayLandParcels(data) {
  const result = document.getElementById('landResult');
  if (!result) return;
  result.classList.remove('hidden');

  // サマリー
  document.getElementById('land_parcel_count').textContent = data.parcelCount + '件';
  document.getElementById('land_total_bv').textContent =
    data.totalBookValue > 0 ? Math.round(data.totalBookValue).toLocaleString() + ' 百万円' : '-';

  const methodEl = document.getElementById('land_gain_method_detail');
  if (data.gainMethod) {
    methodEl.textContent = '推定方法: ' + data.gainMethod;
  }

  // 各parcelに用途と調整係数を付与してからテーブル描画
  if (data.parcels && data.parcels.length > 0) {
    data.parcels.forEach(p => {
      // APIが用途情報を返していればそれを使い、なければフロントで推定
      if (p.useType) {
        p._useKey = p.useType;
        p._useLabel = p.useLabel || p.useType;
        p._basePricePerSqm = p.basePricePerSqm || null;
      } else {
        var use = guessLandUse(p.name, p.address, p.area);
        p._useKey = use.key;
        p._useLabel = use.label;
        // APIが一律0.5適用の旧版の場合は戻す
        if (p.estimatedPricePerSqm && !p.basePricePerSqm) {
          p._basePricePerSqm = Math.round(p.estimatedPricePerSqm / 0.5);
        }
      }
      // UIのデフォルト調整係数を適用
      p._rate = getDefaultRate(p._useKey);
    });
  }

  renderLandTable(data);
}

function renderLandTable(data) {
  const tbody = document.getElementById('landParcelsBody');
  tbody.innerHTML = '';

  var totalFV = 0, totalGain = 0;

  if (data.parcels && data.parcels.length > 0) {
    data.parcels.forEach((p, i) => {
      var rate = p._rate || 0.5;
      var adjPrice = p._basePricePerSqm ? Math.round(p._basePricePerSqm * rate) : (p.estimatedPricePerSqm || null);
      var adjValue = (adjPrice && p.area) ? Math.round(p.area * adjPrice / 1000000) : (p.estimatedValue || null);
      var gain = (adjValue != null && p.bookValue) ? adjValue - p.bookValue : null;
      var gainColor = gain != null ? (gain > 0 ? 'color:var(--green);font-weight:600;' : 'color:var(--danger);') : '';

      if (adjValue != null) totalFV += adjValue;
      if (gain != null) totalGain += gain;

      // 用途selectのオプション生成
      var useOptions = LAND_USE_RULES.map(r =>
        '<option value="' + r.key + '"' + (r.key === p._useKey ? ' selected' : '') + '>' + r.label + '</option>'
      ).join('') + '<option value="other"' + (p._useKey === 'other' ? ' selected' : '') + '>その他</option>';

      // 推定単価の内訳ツールチップ
      var priceTip = p._basePricePerSqm
        ? '基準単価 ' + Math.round(p._basePricePerSqm).toLocaleString() + '円/㎡ × ' + rate.toFixed(2) + '（' + (p._useLabel || 'その他') + '）'
        : '';

      tbody.innerHTML += '<tr>' +
        '<td>' + (i + 1) + '</td>' +
        '<td style="max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + (p.name || '') + '">' + (p.name || '-') + '</td>' +
        '<td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + (p.address || '') + '">' + (p.address || '-') + '</td>' +
        '<td><select style="font-size:.68rem;padding:1px 2px;border:1px solid var(--mid-gray);border-radius:3px;" title="事業所名から自動推定" onchange="updateParcelUse(' + i + ',this.value)">' + useOptions + '</select></td>' +
        '<td style="text-align:right;">' + (p.area ? p.area.toLocaleString() : '-') + '</td>' +
        '<td style="text-align:right;">' + (p.bookValue ? Math.round(p.bookValue).toLocaleString() : '-') + '</td>' +
        '<td style="text-align:center;" title="' + (p._useLabel || 'その他') + 'のデフォルト: ' + getDefaultRate(p._useKey).toFixed(2) + '"><input type="number" value="' + rate.toFixed(2) + '" step="0.05" min="0" max="3" style="width:48px;font-size:.68rem;text-align:center;padding:1px 2px;border:1px solid var(--mid-gray);border-radius:3px;" onchange="updateParcelRate(' + i + ',this.value)"></td>' +
        '<td style="text-align:right;cursor:help;" title="' + priceTip + '">' + (adjPrice ? Math.round(adjPrice).toLocaleString() : '-') + '</td>' +
        '<td style="text-align:right;">' + (adjValue != null ? Math.round(adjValue).toLocaleString() : '-') + '</td>' +
        '<td style="text-align:right;' + gainColor + '">' + (gain != null ? Math.round(gain).toLocaleString() : '-') + '</td>' +
        '</tr>';
    });
  } else {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-light);">土地明細が見つかりませんでした</td></tr>';
  }

  // サマリー更新
  document.getElementById('land_total_fv').textContent =
    totalFV > 0 ? Math.round(totalFV).toLocaleString() + ' 百万円' : '-';
  var gainEl = document.getElementById('land_total_gain');
  gainEl.textContent = Math.round(totalGain).toLocaleString() + ' 百万円';
  gainEl.style.color = totalGain > 0 ? 'var(--green)' : 'var(--danger)';

  // landParcelsDataの合計も更新
  if (landParcelsData) {
    landParcelsData.totalEstimatedValue = totalFV;
    landParcelsData.totalEstimatedGain = totalGain;
  }
}

function updateParcelUse(idx, useKey) {
  if (!landParcelsData || !landParcelsData.parcels[idx]) return;
  var p = landParcelsData.parcels[idx];
  var rule = LAND_USE_RULES.find(r => r.key === useKey) || { key: 'other', label: 'その他' };
  p._useKey = rule.key;
  p._useLabel = rule.label;
  p._rate = getDefaultRate(rule.key);
  renderLandTable(landParcelsData);
}

function updateParcelRate(idx, val) {
  if (!landParcelsData || !landParcelsData.parcels[idx]) return;
  landParcelsData.parcels[idx]._rate = parseFloat(val) || 0.5;
  renderLandTable(landParcelsData);
}

function recalcLandWithRates() {
  if (!landParcelsData || !landParcelsData.parcels) return;
  landParcelsData.parcels.forEach(p => {
    p._rate = getDefaultRate(p._useKey || 'other');
  });
  renderLandTable(landParcelsData);
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
  if (!body) return;
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

  var phCtEl = document.getElementById('ph_current_total');
  if (phCtEl) phCtEl.textContent = counted > 0 ? Math.round(totalCurrent).toLocaleString() + ' 百万円' : '-';
  var gainEl = document.getElementById('ph_gain');
  if (gainEl) {
    if (counted > 0) {
      var gain = totalCurrent - totalReport;
      gainEl.textContent = (gain > 0 ? '+' : '') + Math.round(gain).toLocaleString() + ' 百万円';
      gainEl.style.color = gain > 0 ? 'var(--green)' : '#c0392b';
    } else {
      gainEl.textContent = '-';
      gainEl.style.color = '';
    }
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
// 株主構成・大株主の状況
// ═══════════════════════════════════════════════════════════════

function displayMajorShareholders() {
  var e = indData.edinet || {};
  var section = document.getElementById('shareholderSection');
  if (!section || !e.majorShareholders || e.majorShareholders.length === 0) {
    if (section) section.style.display = 'none';
    return;
  }
  section.style.display = '';

  // サマリー
  var shc = document.getElementById('sh_count'); if (shc) shc.textContent = e.majorShareholdersCount + '名';
  var shr = document.getElementById('sh_total_ratio'); if (shr) shr.textContent = e.majorShareholdersTotalRatio.toFixed(1) + '%';

  // 浮動株比率概算（100% - 上位株主合計）
  var floatRatio = Math.max(0, 100 - e.majorShareholdersTotalRatio);
  var floatEl = document.getElementById('sh_float_ratio');
  if (floatEl) {
    floatEl.textContent = floatRatio.toFixed(1) + '%';
    floatEl.style.color = floatRatio > 60 ? 'var(--red)' : floatRatio > 40 ? 'var(--orange)' : 'var(--green)';
  }

  // カテゴリ別比率バー
  var cats = e.shareholderCategories || {};
  var catLabels = {
    trust: { label: '信託銀行（受託）', color: '#3498db' },
    foreign: { label: '外国法人等', color: '#e74c3c' },
    insurance: { label: '保険会社', color: '#2ecc71' },
    bank: { label: '銀行', color: '#9b59b6' },
    fund: { label: 'ファンド・投資会社', color: '#e67e22' },
    treasury: { label: '自己株式', color: '#95a5a6' },
    other: { label: 'その他（個人・事業法人等）', color: '#1abc9c' }
  };
  var chartHtml = '<div style="margin-bottom:8px;font-size:.72rem;font-weight:600;color:var(--navy);">株主カテゴリ別比率</div>';
  chartHtml += '<div style="display:flex;height:22px;border-radius:4px;overflow:hidden;margin-bottom:8px;">';
  for (var k in catLabels) {
    if (cats[k] && cats[k] > 0) {
      chartHtml += '<div style="width:' + cats[k] + '%;background:' + catLabels[k].color + ';min-width:2px;" title="' + catLabels[k].label + ': ' + cats[k].toFixed(1) + '%"></div>';
    }
  }
  chartHtml += '</div>';
  chartHtml += '<div style="display:flex;flex-wrap:wrap;gap:8px 16px;font-size:.68rem;">';
  for (var k in catLabels) {
    if (cats[k] && cats[k] > 0) {
      chartHtml += '<span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:' + catLabels[k].color + ';margin-right:4px;vertical-align:middle;"></span>' + catLabels[k].label + ' ' + cats[k].toFixed(1) + '%</span>';
    }
  }
  chartHtml += '</div>';
  document.getElementById('shCategoryChart').innerHTML = chartHtml;

  // 大株主一覧テーブル
  var tbody = document.getElementById('shareholderBody');
  tbody.innerHTML = '';
  var maxRatio = e.majorShareholders[0] ? e.majorShareholders[0].ratio : 1;
  e.majorShareholders.forEach(function(sh, i) {
    var barWidth = Math.max(2, (sh.ratio / maxRatio) * 100);
    var barColor = sh.ratio >= 10 ? 'var(--red)' : sh.ratio >= 5 ? 'var(--orange)' : 'var(--gold)';
    tbody.innerHTML += '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td style="font-family:var(--sans);max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + sh.name + '">' + sh.name + '</td>' +
      '<td style="text-align:right;">' + (sh.shares ? (sh.shares >= 10000 ? Math.round(sh.shares / 1000).toLocaleString() + '千株' : sh.shares.toLocaleString() + '株') : '-') + '</td>' +
      '<td style="text-align:right;font-weight:600;">' + sh.ratio.toFixed(2) + '%</td>' +
      '<td><div style="width:100%;height:8px;background:var(--light-gray);border-radius:4px;"><div style="width:' + barWidth + '%;height:100%;background:' + barColor + ';border-radius:4px;"></div></div></td>' +
      '</tr>';
  });
}

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
  if (!window.currentUser) { alert('この機能を利用するには無料会員登録が必要です。ページ下部からログインまたは新規登録してください。'); return; }
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
          // NC・EV指標を計算
          const debt = (d.shortTermBorrowings || 0) + (d.currentPortionLongTermDebt || 0) +
                       (d.longTermBorrowings || 0) + (d.bondsPayable || 0) + (d.currentPortionBonds || 0);
          const cash = (d.cashAndDeposits || 0) + (d.shortTermSecurities || 0);
          const netCash = cash - debt;
          const mcapM = d.marketCapOku ? d.marketCapOku * 100 : null; // 百万円
          let ncRatio = null, adjNcRatio = null, fullNcRatio = null, evEbitda = null;
          if (d.cashAndDeposits != null && mcapM) {
            ncRatio = netCash / mcapM * 100;
          }
          const secMV = d.securitiesMarketValue || 0;
          if (d.cashAndDeposits != null && secMV > 0 && mcapM) {
            adjNcRatio = (netCash + secMV) / mcapM * 100;
          }
          const ipFV = d.investmentPropertyFairValue || 0;
          if (d.cashAndDeposits != null && (secMV > 0 || ipFV > 0) && mcapM) {
            fullNcRatio = (netCash + secMV + ipFV) / mcapM * 100;
          }
          if (mcapM && d.operatingIncome != null && d.depreciationAndAmortization != null) {
            const ev = mcapM + debt - cash;
            const ebitda = d.operatingIncome + d.depreciationAndAmortization;
            if (ebitda > 0) evEbitda = ev / ebitda;
          }
          // risk-assessment と同じ配点で再計算
          const fullScore = calcQuickScore(d, ncRatio, adjNcRatio, evEbitda);
          return { ...d, score: fullScore, ncRatio, adjNcRatio, fullNcRatio, evEbitda };
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

// アクティビストスコア（risk-assessment.html と同一配点）
// 12項目 / 合計100点満点（正規化）
function calcQuickScore(d, ncRatio, adjNcRatio, evEbitda) {
  let total = 0, maxTotal = 0;

  // PBR (max 15)
  if (d.pbr != null) { total += d.pbr < 0.5 ? 15 : d.pbr < 0.7 ? 12 : d.pbr < 0.8 ? 10 : d.pbr < 1.0 ? 8 : d.pbr < 1.3 ? 5 : 2; maxTotal += 15; }
  // ROE (max 10)
  if (d.roe != null) { total += d.roe < 3 ? 10 : d.roe < 5 ? 8 : d.roe < 8 ? 6 : d.roe < 10 ? 3 : 1; maxTotal += 10; }
  // 配当性向 (max 8)
  if (d.payoutRatio != null) { total += d.payoutRatio < 20 ? 8 : d.payoutRatio < 30 ? 6 : d.payoutRatio < 40 ? 4 : d.payoutRatio < 50 ? 2 : 1; maxTotal += 8; }
  // NC/時価総額 (max 12)
  if (ncRatio != null) { total += ncRatio > 50 ? 12 : ncRatio > 30 ? 10 : ncRatio > 20 ? 7 : ncRatio > 10 ? 4 : 2; maxTotal += 12; }
  // 実質NC/時価総額 (max 10)
  if (adjNcRatio != null) { total += adjNcRatio > 80 ? 10 : adjNcRatio > 50 ? 8 : adjNcRatio > 30 ? 6 : adjNcRatio > 15 ? 3 : 1; maxTotal += 10; }
  // EV/EBITDA (max 8)
  if (evEbitda != null) { total += evEbitda < 3 ? 8 : evEbitda < 5 ? 7 : evEbitda < 7 ? 5 : evEbitda < 10 ? 3 : 1; maxTotal += 8; }
  // 自己資本比率 (max 7)
  if (d.equityRatio != null) { total += d.equityRatio > 80 ? 7 : d.equityRatio > 70 ? 5 : d.equityRatio > 60 ? 3 : d.equityRatio > 50 ? 2 : 1; maxTotal += 7; }
  // 時価総額 (max 8)
  if (d.marketCapOku != null) { total += d.marketCapOku < 100 ? 3 : d.marketCapOku < 500 ? 8 : d.marketCapOku < 1000 ? 7 : d.marketCapOku < 3000 ? 5 : d.marketCapOku < 5000 ? 4 : 2; maxTotal += 8; }

  // EDINET項目（バッチキャッシュにない場合はスキップ）
  // 外国人持株比率は batch.js 未対応のため除外
  // 政策保有・社外取締役・買収防衛策もバッチ未対応のため除外

  if (maxTotal === 0) return 0;
  return Math.round((total / maxTotal) * 100);
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

  // 時価総額区分別社数を集計・表示
  var mcapCounts = { nano: 0, micro: 0, small: 0, mid: 0, large: 0 };
  rankResults.forEach(function(d) {
    if (d.marketCapOku == null) return;
    if (d.marketCapOku < 50) mcapCounts.nano++;
    else if (d.marketCapOku < 300) mcapCounts.micro++;
    else if (d.marketCapOku < 1000) mcapCounts.small++;
    else if (d.marketCapOku < 5000) mcapCounts.mid++;
    else mcapCounts.large++;
  });
  var legend = document.getElementById('mcapCategoryLegend');
  if (legend && rankResults.length > 0) {
    legend.style.display = '';
    document.getElementById('mcapCountNano').textContent = mcapCounts.nano + '社';
    document.getElementById('mcapCountMicro').textContent = mcapCounts.micro + '社';
    document.getElementById('mcapCountSmall').textContent = mcapCounts.small + '社';
    document.getElementById('mcapCountMid').textContent = mcapCounts.mid + '社';
    document.getElementById('mcapCountLarge').textContent = mcapCounts.large + '社';
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
      case 'ncRatio': return (b.ncRatio || -999) - (a.ncRatio || -999);
      case 'adjNcRatio': return (b.adjNcRatio || -999) - (a.adjNcRatio || -999);
      case 'fullNcRatio': return (b.fullNcRatio || -999) - (a.fullNcRatio || -999);
      case 'evEbitda': return (a.evEbitda || 999) - (b.evEbitda || 999);
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

  // ── NC・EVビュー ──
  const tbodyNcev = document.getElementById('rankTableBody-ncev');
  tbodyNcev.innerHTML = '';
  filtered.forEach((d, i) => {
    const rank = i + 1;
    const rankCls = rank <= 3 ? `rank-${rank}` : 'rank-other';
    const hasData = d.hasEdinetData;
    const fmtPct = v => v != null ? v.toFixed(1) + '%' : '-';
    const ncCls = v => v != null && v > 30 ? 'style="color:var(--danger);font-weight:700;"' : v != null && v > 10 ? 'style="color:var(--orange);font-weight:600;"' : '';
    const evCls = d.evEbitda != null && d.evEbitda < 5 ? 'style="color:var(--green);font-weight:700;"' : d.evEbitda != null && d.evEbitda < 10 ? 'style="color:var(--orange);"' : '';
    tbodyNcev.innerHTML += `<tr style="${!hasData ? 'opacity:0.5;' : ''}">
      <td><span class="rank-badge ${rankCls}">${rank}</span></td>
      <td>${d.code}</td>
      <td class="name-cell" style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${d.companyName || '-'}</td>
      <td style="font-weight:700;color:${d.score >= 60 ? 'var(--danger)' : d.score >= 40 ? 'var(--orange)' : 'var(--green)'};">${d.score || '-'}</td>
      <td ${ncCls(d.ncRatio)}>${fmtPct(d.ncRatio)}</td>
      <td ${ncCls(d.adjNcRatio)}>${fmtPct(d.adjNcRatio)}</td>
      <td ${ncCls(d.fullNcRatio)}>${fmtPct(d.fullNcRatio)}</td>
      <td ${evCls}>${d.evEbitda != null ? d.evEbitda.toFixed(1) + '倍' : '-'}</td>
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
      case 'ncRatio': va = a.ncRatio ?? -999; vb = b.ncRatio ?? -999; break;
      case 'adjNcRatio': va = a.adjNcRatio ?? -999; vb = b.adjNcRatio ?? -999; break;
      case 'fullNcRatio': va = a.fullNcRatio ?? -999; vb = b.fullNcRatio ?? -999; break;
      case 'evEbitda': va = a.evEbitda || 999; vb = b.evEbitda || 999; break;
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
