// 認証モーダル制御
function openAuthModal(mode) {
  var overlay = document.getElementById('authModalOverlay');
  if (window.currentUser) {
    document.getElementById('authModalLoggedOut').classList.add('hidden');
    document.getElementById('authModalLoggedIn').classList.remove('hidden');
  } else {
    document.getElementById('authModalLoggedOut').classList.remove('hidden');
    document.getElementById('authModalLoggedIn').classList.add('hidden');
    if (mode === 'register') showModalRegister();
    else showModalLogin();
  }
  overlay.classList.add('show');
  document.body.style.overflow = 'hidden';
}
function closeAuthModal() {
  document.getElementById('authModalOverlay').classList.remove('show');
  document.body.style.overflow = '';
}
function showModalLogin() {
  document.getElementById('authModalLoginForm').classList.remove('hidden');
  document.getElementById('authModalRegisterForm').classList.add('hidden');
  document.getElementById('authModalResetForm').classList.add('hidden');
}
function showModalRegister() {
  document.getElementById('authModalLoginForm').classList.add('hidden');
  document.getElementById('authModalRegisterForm').classList.remove('hidden');
  document.getElementById('authModalResetForm').classList.add('hidden');
}
function showModalReset() {
  document.getElementById('authModalLoginForm').classList.add('hidden');
  document.getElementById('authModalRegisterForm').classList.add('hidden');
  document.getElementById('authModalResetForm').classList.remove('hidden');
}
async function doModalLogin() {
  var e = document.getElementById('authModalEmail').value.trim();
  var p = document.getElementById('authModalPass').value;
  var m = document.getElementById('authModalLoginMsg');
  if (!e || !p) { m.textContent = 'メールとパスワードを入力してください'; m.className = 'auth-msg error'; return; }
  try { await window.firebaseLogin(e, p); m.textContent = 'ログインしました'; m.className = 'auth-msg success'; setTimeout(closeAuthModal, 1000); }
  catch (err) {
    var messages = {
      'auth/invalid-credential': 'メールアドレスまたはパスワードが正しくありません。Googleで登録された方は「Googleでログイン」をご利用ください。',
      'auth/user-not-found': 'このメールアドレスは登録されていません。',
      'auth/wrong-password': 'パスワードが正しくありません。',
      'auth/too-many-requests': 'ログイン試行回数が多すぎます。しばらくしてからお試しください。',
      'auth/invalid-email': 'メールアドレスの形式が正しくありません。'
    };
    m.textContent = messages[err.code] || 'ログイン失敗: ' + (err.message || ''); m.className = 'auth-msg error';
  }
}
async function doModalRegister() {
  var n = document.getElementById('regModalName').value.trim();
  var c = document.getElementById('regModalCompany').value.trim();
  var e = document.getElementById('regModalEmail').value.trim();
  var p = document.getElementById('regModalPass').value;
  var aff = document.getElementById('regModalAffiliation').value;
  var jt = document.getElementById('regModalJobTitle').value;
  var m = document.getElementById('authModalRegMsg');
  if (!n || !e || !p) { m.textContent = '必須項目を入力してください'; m.className = 'auth-msg error'; return; }
  if (!aff) { m.textContent = '所属種別を選択してください'; m.className = 'auth-msg error'; return; }
  if (p.length < 8) { m.textContent = 'パスワードは8文字以上で設定してください。'; m.className = 'auth-msg error'; return; }
  try { await window.firebaseRegister(e, p, n, c, aff, '', jt); m.innerHTML = '登録しました。確認メールをお送りしました。<br><span style="font-size:.68rem;">メールが届かない場合は迷惑メールフォルダをご確認いただくか、<a href="https://aoyama-nogizaka.com/#contact" style="color:var(--green);text-decoration:underline;">お問い合わせフォーム</a>よりご連絡ください。</span>'; m.className = 'auth-msg success'; setTimeout(closeAuthModal, 3000); }
  catch (err) { const msgs = {'auth/email-already-in-use':'このメールアドレスは既に登録されています。ログインしてください。','auth/weak-password':'パスワードは8文字以上で設定してください。','auth/invalid-email':'メールアドレスの形式が正しくありません。'}; m.textContent = msgs[err.code] || '登録に失敗しました。'; m.className = 'auth-msg error'; }
}
async function doModalReset() {
  var e = document.getElementById('authModalResetEmail').value.trim();
  var m = document.getElementById('authModalResetMsg');
  if (!e) { m.textContent = 'メールアドレスを入力してください'; m.className = 'auth-msg error'; return; }
  try { await window.firebaseResetPassword(e); m.textContent = 'リセット用メールを送信しました。'; m.className = 'auth-msg success'; }
  catch (err) {
    var messages = {
      'auth/user-not-found': 'このメールアドレスは登録されていません。',
      'auth/invalid-email': 'メールアドレスの形式が正しくありません。',
      'auth/too-many-requests': '送信回数が多すぎます。しばらくしてからお試しください。'
    };
    m.textContent = messages[err.code || ''] || 'エラー: ' + err.message; m.className = 'auth-msg error';
  }
}

// Google ログイン
async function doModalGoogleLogin(context) {
  var m = document.getElementById(context === 'register' ? 'authModalRegMsg' : 'authModalLoginMsg');
  try {
    var result = await window.firebaseGoogleLogin();
    closeAuthModal();
    if (result && result.needsProfile) {
      openCompleteProfile();
    }
  } catch (err) {
    if (err.code === 'auth/popup-closed-by-user') return;
    var messages = {
      'auth/account-exists-with-different-credential': 'このメールアドレスは別の方法で登録されています。メールアドレスでログインしてください。',
      'auth/popup-blocked': 'ポップアップがブロックされました。ブラウザの設定を確認してください。'
    };
    m.textContent = messages[err.code] || 'Googleログインに失敗しました: ' + (err.message || '');
    m.className = 'auth-msg error';
  }
}

// プロフィール補完モーダル
function openCompleteProfile() {
  document.getElementById('completeProfileOverlay').classList.add('show');
  document.body.style.overflow = 'hidden';
}
function closeCompleteProfile() {
  document.getElementById('completeProfileOverlay').classList.remove('show');
  document.body.style.overflow = '';
}
async function doCompleteProfile() {
  var aff = document.getElementById('cpModalAffiliation').value;
  var jt = document.getElementById('cpModalJobTitle').value;
  var comp = document.getElementById('cpModalCompany').value.trim();
  var m = document.getElementById('completeProfileMsg');
  if (!aff || !jt || !comp) { m.textContent = '全ての項目を入力してください'; m.className = 'auth-msg error'; return; }
  try {
    var user = window.currentUser;
    if (!user) throw new Error('ログインされていません');
    await window.firebaseUpdateProfile(user.uid, { affiliation: aff, jobTitle: jt, company: comp });
    m.textContent = '登録しました'; m.className = 'auth-msg success';
    setTimeout(closeCompleteProfile, 1000);
  } catch (err) {
    m.textContent = 'プロフィールの保存に失敗しました: ' + (err.message || ''); m.className = 'auth-msg error';
  }
}

// scrollToAuthをモーダル版にオーバーライド
scrollToAuth = function(mode) { openAuthModal(mode); };

/* ── Event Delegation ── */
(function() {
  // Click delegation
  document.addEventListener('click', function(e) {
    var el = e.target.closest('[data-action]');
    if (!el) return;
    var parts = el.dataset.action.split(':');
    var fn = parts[0], arg = parts[1] || null;

    switch(fn) {
      case 'scrollToAuth': scrollToAuth(arg); break;
      case 'openMyPage': openMyPage(); break;
      case 'doLogout': doLogout(); break;
      case 'openAuthModal': openAuthModal(arg); break;
      case 'switchTab': switchTab(arg); break;
      case 'fetchIndividual': fetchIndividual(); break;
      case 'resetIndividual': resetIndividual(); break;
      case 'recalcIndividual': recalcIndividual(); break;
      case 'fetchLandParcels': fetchLandParcels(); break;
      case 'applyLandGainEstimate': applyLandGainEstimate(); break;
      case 'fetchPolicyHoldingsPrices': fetchPolicyHoldingsPrices(); break;
      case 'applyPolicyHoldingsGain': applyPolicyHoldingsGain(); break;
      case 'loadPreset': loadPreset(arg); break;
      case 'startRankingScan': startRankingScan(); break;
      case 'cancelScan': cancelScan(); break;
      case 'exportCSV': exportCSV(); break;
      case 'switchRankView': switchRankView(arg); break;
      case 'sortRanking': sortRanking(arg); break;
      case 'doLogin': doLogin(); break;
      case 'doRegister': doRegister(); break;
      case 'doModalLogin': doModalLogin(); break;
      case 'doModalRegister': doModalRegister(); break;
      case 'doModalReset': doModalReset(); break;
      case 'doModalGoogleLogin': doModalGoogleLogin(arg); break;
      case 'showLogin': showLogin(); break;
      case 'showRegister': showRegister(); break;
      case 'showResetPassword': showResetPassword(); break;
      case 'showModalLogin': showModalLogin(); break;
      case 'showModalRegister': showModalRegister(); break;
      case 'showModalReset': showModalReset(); break;
      case 'closeAuthModal': closeAuthModal(); break;
      case 'doCompleteProfile': doCompleteProfile(); break;
      case 'closeCompleteProfile': closeCompleteProfile(); break;
    }
    // Prevent default for anchor tags
    if (el.tagName === 'A') e.preventDefault();
  });

  // Change delegation
  document.addEventListener('change', function(e) {
    var el = e.target.closest('[data-change]');
    if (!el) return;
    var fn = el.dataset.change;
    if (fn === 'recalcLandWithRates') recalcLandWithRates();
    else if (fn === 'filterAndDisplayRanking') filterAndDisplayRanking();
    else if (fn === 'recalcPolicyHoldings') {
      if (typeof phPriceData !== 'undefined' && phPriceData) {
        var ed = (typeof indData !== 'undefined' && indData.edinet) ? indData.edinet : {};
        renderPolicyHoldingsTable(ed.policyHoldingsTop, phPriceData);
        updatePolicyHoldingsSummary(ed.policyHoldingsTop, phPriceData);
      }
    }
    else if (fn === 'toggleAffiliationCode') toggleAffiliationCode();
  });

  // Overlay click-to-close delegation
  document.addEventListener('click', function(e) {
    if (e.target.id === 'authModalOverlay') closeAuthModal();
    if (e.target.id === 'completeProfileOverlay') closeCompleteProfile();
  });
})();
