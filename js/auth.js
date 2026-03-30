/**
 * auth.js — 共通認証UIハンドラー
 *
 * 用途: ログイン/登録/リセット/マイページ等のモーダル操作・認証処理を一元管理
 * 依存: window.firebaseLogin, window.firebaseGoogleLogin, window.firebaseRegister,
 *       window.firebaseResetPassword, window.firebaseLogout, window.firebaseGetProfile,
 *       window.firebaseUpdateProfile, window.firebaseLinkPassword, window.firebaseDeleteAccount
 *       (すべて Firebase module スクリプトで設定済み)
 *
 * HTML側: onclick/onsubmit を排除し data 属性 + イベント委譲で駆動
 *   data-auth-action="openModal:login"      → クリック時に openModal('login')
 *   data-auth-submit="handleLogin"           → フォーム送信時に handleLogin(e)
 *   data-auth-change="toggleAffiliationCode" → change 時に toggleAffiliationCode()
 *
 * 配置先: js/auth.js
 * 読み込み: <script src="js/auth.js"></script> (Firebase module の前、</body>直前付近)
 */
(function () {
  'use strict';

  // ===== LABELS =====
  var AFFILIATION_LABELS = {
    listed_company: '上場企業',
    institutional_investor: '機関投資家',
    individual_investor: '個人投資家',
    consulting: 'コンサルティングファーム',
    legal: '法律事務所・弁護士',
    media: 'メディア・報道機関',
    academic: '学術・研究機関',
    other: 'その他'
  };

  var JOBTITLE_LABELS = {
    executive: '経営者・役員',
    department_head: '部長・本部長',
    manager: '課長・マネージャー',
    ir_officer: 'IR担当',
    legal_compliance: '法務・コンプライアンス',
    finance: '経理・財務',
    analyst: 'アナリスト・ファンドマネージャー',
    consultant: 'コンサルタント・アドバイザー',
    individual: '個人',
    other: 'その他'
  };

  // ===== MODAL =====
  function openModal(type) {
    var el = document.getElementById(type + 'Modal');
    if (el) { el.classList.add('show'); document.body.style.overflow = 'hidden'; }
  }

  function closeModal(type) {
    var el = document.getElementById(type + 'Modal');
    if (el) { el.classList.remove('show'); document.body.style.overflow = ''; }
  }

  function switchModal(type) {
    ['login', 'register', 'resetPassword'].forEach(function (t) {
      if (t !== type) closeModal(t);
    });
    openModal(type);
  }

  function switchToLogin() {
    closeModal('resetPassword');
    openModal('login');
  }

  // ===== AUTH HANDLERS =====

  function handleGoogleLogin(modalType) {
    var errEl = document.getElementById(modalType === 'login' ? 'loginError' : 'registerError');
    if (errEl) errEl.style.display = 'none';
    window.firebaseGoogleLogin().then(function (result) {
      closeModal(modalType);
      if (result && result.needsProfile) {
        openModal('completeProfile');
      }
    }).catch(function (err) {
      if (err.code === 'auth/popup-closed-by-user') return;
      var messages = {
        'auth/account-exists-with-different-credential': 'このメールアドレスは別の方法で登録されています。メールアドレスでログインしてください。',
        'auth/popup-blocked': 'ポップアップがブロックされました。ブラウザの設定を確認してください。'
      };
      if (errEl) {
        errEl.textContent = messages[err.code] || 'Googleログインに失敗しました: ' + (err.message || '');
        errEl.style.display = 'block';
      }
    });
  }

  function handleLogin(e) {
    e.preventDefault();
    var errEl = document.getElementById('loginError');
    var btn = document.getElementById('loginBtn');
    if (errEl) errEl.style.display = 'none';
    if (btn) { btn.textContent = 'ログイン中...'; btn.disabled = true; }

    var email = document.getElementById('loginEmail').value;
    var password = document.getElementById('loginPassword').value;

    window.firebaseLogin(email, password).then(function () {
      closeModal('login');
      document.getElementById('loginEmail').value = '';
      document.getElementById('loginPassword').value = '';
    }).catch(function (err) {
      var messages = {
        'auth/user-not-found': 'このメールアドレスは登録されていません。',
        'auth/wrong-password': 'パスワードが正しくありません。',
        'auth/invalid-credential': 'メールアドレスまたはパスワードが正しくありません。Googleで登録された方は「Googleでログイン」をご利用ください。',
        'auth/too-many-requests': 'ログイン試行回数が多すぎます。しばらくしてからお試しください。',
        'auth/invalid-email': 'メールアドレスの形式が正しくありません。'
      };
      var code = err.code || '';
      if (errEl) {
        errEl.textContent = messages[code] || 'ログインに失敗しました: ' + (err.message || '');
        errEl.style.display = 'block';
      }
    }).finally(function () {
      if (btn) { btn.textContent = 'ログイン'; btn.disabled = false; }
    });
  }

  function handleRegister(e) {
    e.preventDefault();
    var errEl = document.getElementById('registerError');
    var successEl = document.getElementById('registerSuccess');
    var btn = document.getElementById('registerBtn');
    if (errEl) errEl.style.display = 'none';
    if (successEl) successEl.style.display = 'none';
    if (btn) { btn.textContent = '登録中...'; btn.disabled = true; }

    var name = document.getElementById('regName').value;
    var company = document.getElementById('regCompany').value;
    var email = document.getElementById('regEmail').value;
    var password = document.getElementById('regPassword').value;
    var affiliation = document.getElementById('regAffiliation').value;
    var affiliationCode = document.getElementById('regAffiliationCode')
      ? document.getElementById('regAffiliationCode').value : '';
    var jobTitle = document.getElementById('regJobTitle').value;

    window.firebaseRegister(email, password, name, company, affiliation, affiliationCode, jobTitle)
      .then(function () {
        if (successEl) {
          successEl.innerHTML = '登録が完了しました。確認メールをお送りしましたのでご確認ください。<br><span style="font-size:12px;color:#065f46;">メールが届かない場合は迷惑メールフォルダをご確認いただくか、<a href="https://aoyama-nogizaka.com/#contact" style="color:#047857;text-decoration:underline;">お問い合わせフォーム</a>よりご連絡ください。</span>';
          successEl.style.display = 'block';
        }
        setTimeout(function () {
          closeModal('register');
          if (successEl) successEl.style.display = 'none';
        }, 2500);
      }).catch(function (err) {
        var messages = {
          'auth/email-already-in-use': 'このメールアドレスは既に登録されています。ログインしてください。',
          'auth/weak-password': 'パスワードは8文字以上で設定してください。',
          'auth/invalid-email': 'メールアドレスの形式が正しくありません。'
        };
        var code = err.code || '';
        if (errEl) {
          errEl.textContent = messages[code] || '登録に失敗しました: ' + (err.message || '');
          errEl.style.display = 'block';
        }
      }).finally(function () {
        if (btn) { btn.textContent = '無料で登録する'; btn.disabled = false; }
      });
  }

  function handleResetPassword(e) {
    e.preventDefault();
    var errEl = document.getElementById('resetError');
    var successEl = document.getElementById('resetSuccess');
    var btn = document.getElementById('resetBtn');
    if (errEl) errEl.style.display = 'none';
    if (successEl) successEl.style.display = 'none';
    if (btn) { btn.textContent = '送信中...'; btn.disabled = true; }

    var email = document.getElementById('resetEmail').value;
    window.firebaseResetPassword(email).then(function () {
      if (successEl) {
        successEl.textContent = 'リセット用メールを送信しました。メールをご確認ください。';
        successEl.style.display = 'block';
      }
    }).catch(function (err) {
      var messages = {
        'auth/user-not-found': 'このメールアドレスは登録されていません。',
        'auth/invalid-email': 'メールアドレスの形式が正しくありません。',
        'auth/too-many-requests': '送信回数が多すぎます。しばらくしてからお試しください。'
      };
      if (errEl) {
        errEl.textContent = messages[err.code || ''] || 'エラーが発生しました: ' + (err.message || '');
        errEl.style.display = 'block';
      }
    }).finally(function () {
      if (btn) { btn.textContent = 'リセットメールを送信'; btn.disabled = false; }
    });
  }

  function handleLogout() {
    window.firebaseLogout().catch(function (err) {
      console.error('Logout error:', err);
    });
  }

  function handleCompleteProfile(e) {
    e.preventDefault();
    var errEl = document.getElementById('completeProfileError');
    var btn = document.getElementById('completeProfileBtn');
    if (errEl) errEl.style.display = 'none';
    if (btn) { btn.textContent = '保存中...'; btn.disabled = true; }

    var user = window.currentUser;
    if (!user) {
      if (errEl) { errEl.textContent = 'ログインされていません'; errEl.style.display = 'block'; }
      if (btn) { btn.textContent = '登録して利用開始'; btn.disabled = false; }
      return;
    }

    var affiliation = document.getElementById('cpAffiliation').value;
    var affiliationCode = document.getElementById('cpAffiliationCode')
      ? document.getElementById('cpAffiliationCode').value : '';
    var jobTitle = document.getElementById('cpJobTitle').value;
    var company = document.getElementById('cpCompany').value;

    window.firebaseUpdateProfile(user.uid, {
      affiliation: affiliation,
      affiliationCode: affiliationCode,
      jobTitle: jobTitle,
      company: company
    }).then(function () {
      closeModal('completeProfile');
    }).catch(function (err) {
      if (errEl) {
        errEl.textContent = 'プロフィールの保存に失敗しました: ' + (err.message || '');
        errEl.style.display = 'block';
      }
    }).finally(function () {
      if (btn) { btn.textContent = '登録して利用開始'; btn.disabled = false; }
    });
  }

  // 所属種別で証券コード欄の表示切替 (登録フォーム)
  function toggleAffiliationCode() {
    var v = document.getElementById('regAffiliation').value;
    var g = document.getElementById('affiliationCodeGroup');
    if (g) g.style.display = v === 'listed_company' ? '' : 'none';
  }

  // 所属種別で証券コード欄の表示切替 (プロフィール完了フォーム)
  function toggleCpAffiliationCode() {
    var v = document.getElementById('cpAffiliation').value;
    var g = document.getElementById('cpAffiliationCodeGroup');
    if (g) g.style.display = v === 'listed_company' ? '' : 'none';
  }

  // ===== MY PAGE =====

  function openMyPage() {
    openModal('myPage');
    var user = window.currentUser;
    if (!user) return;
    document.getElementById('myName').textContent = user.displayName || '-';
    document.getElementById('myEmail').textContent = user.email || '-';
    var verified = user.emailVerified;
    var verifiedEl = document.getElementById('myEmailVerified');
    if (verifiedEl) {
      verifiedEl.innerHTML = verified
        ? '<span style="color:#27ae60;">認証済み ✓</span>'
        : '<span style="color:#c0392b;">未認証</span>';
    }
    var btnResend = document.getElementById('btnResendVerify');
    if (btnResend) btnResend.style.display = verified ? 'none' : '';
    var hasPassword = user.providerData && user.providerData.some(function (p) {
      return p.providerId === 'password';
    });
    var btnChange = document.getElementById('btnChangePassword');
    var btnSet = document.getElementById('btnSetPassword');
    if (btnChange) btnChange.style.display = hasPassword ? '' : 'none';
    if (btnSet) btnSet.style.display = hasPassword ? 'none' : '';
    var setPwForm = document.getElementById('setPasswordForm');
    if (setPwForm) setPwForm.style.display = 'none';
    var msg = document.getElementById('myPageMsg');
    if (msg) msg.style.display = 'none';

    window.firebaseGetProfile(user.uid).then(function (profile) {
      if (profile) {
        var el;
        el = document.getElementById('myAffiliation');
        if (el) el.textContent = AFFILIATION_LABELS[profile.affiliation] || profile.affiliation || '-';
        el = document.getElementById('myJobTitle');
        if (el) el.textContent = JOBTITLE_LABELS[profile.jobTitle] || profile.jobTitle || '-';
        el = document.getElementById('myCompany');
        if (el) el.textContent = profile.company || '-';
        el = document.getElementById('myCreatedAt');
        if (el) el.textContent = profile.createdAt
          ? new Date(profile.createdAt.seconds * 1000).toLocaleDateString('ja-JP') : '-';
      }
    }).catch(function (e) { console.error('Profile fetch error:', e); });
  }

  function handleMyPageResetPassword() {
    var msg = document.getElementById('myPageMsg');
    var user = window.currentUser;
    window.firebaseResetPassword(user.email).then(function () {
      if (msg) {
        msg.textContent = 'パスワードリセット用メールを送信しました。';
        msg.style.display = 'block';
        msg.style.background = '#d1fae5';
        msg.style.color = '#065f46';
      }
    }).catch(function (e) {
      if (msg) {
        msg.textContent = 'エラーが発生しました: ' + (e.message || '');
        msg.style.display = 'block';
        msg.style.background = '#fee2e2';
        msg.style.color = '#991b1b';
      }
    });
  }

  function showSetPasswordForm() {
    var form = document.getElementById('setPasswordForm');
    if (form) form.style.display = '';
    var msg = document.getElementById('myPageMsg');
    if (msg) msg.style.display = 'none';
    var input = document.getElementById('setPasswordInput');
    if (input) { input.value = ''; input.focus(); }
    var confirm = document.getElementById('setPasswordConfirm');
    if (confirm) confirm.value = '';
  }

  function handleSetPassword() {
    var msg = document.getElementById('myPageMsg');
    var pw = document.getElementById('setPasswordInput').value;
    var pw2 = document.getElementById('setPasswordConfirm').value;
    if (msg) msg.style.display = 'none';

    if (!pw || pw.length < 8) {
      if (msg) {
        msg.textContent = 'パスワードは8文字以上で設定してください。';
        msg.style.display = 'block'; msg.style.background = '#fee2e2'; msg.style.color = '#991b1b';
      }
      return;
    }
    if (pw !== pw2) {
      if (msg) {
        msg.textContent = 'パスワードが一致しません。';
        msg.style.display = 'block'; msg.style.background = '#fee2e2'; msg.style.color = '#991b1b';
      }
      return;
    }

    var user = window.currentUser;
    window.firebaseLinkPassword(user.email, pw).then(function () {
      if (msg) {
        msg.textContent = 'パスワードを設定しました。メールアドレスでもログインできます。';
        msg.style.display = 'block'; msg.style.background = '#d1fae5'; msg.style.color = '#065f46';
      }
      var form = document.getElementById('setPasswordForm');
      if (form) form.style.display = 'none';
      var btnSet = document.getElementById('btnSetPassword');
      if (btnSet) btnSet.style.display = 'none';
      var btnChange = document.getElementById('btnChangePassword');
      if (btnChange) btnChange.style.display = '';
    }).catch(function (e) {
      var messages = {
        'auth/provider-already-linked': 'すでにパスワードが設定されています。「パスワードを変更」からリセットしてください。',
        'auth/weak-password': 'パスワードが弱すぎます。8文字以上で設定してください。',
        'auth/requires-recent-login': 'セキュリティのため再ログインが必要です。一度ログアウトしてGoogleで再ログインしてください。'
      };
      if (msg) {
        msg.textContent = messages[e.code] || 'エラーが発生しました。しばらくしてから再度お試しください。';
        msg.style.display = 'block'; msg.style.background = '#fee2e2'; msg.style.color = '#991b1b';
      }
    });
  }

  function handleResendVerification() {
    var msg = document.getElementById('myPageMsg');
    import('https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js').then(function (mod) {
      return mod.sendEmailVerification(window.currentUser);
    }).then(function () {
      if (msg) {
        msg.innerHTML = '認証メールを再送しました。メールをご確認ください。<br><span style="font-size:11px;">届かない場合は迷惑メールフォルダをご確認いただくか、<a href="https://aoyama-nogizaka.com/#contact" style="color:#047857;text-decoration:underline;">お問い合わせフォーム</a>よりご連絡ください。</span>';
        msg.style.display = 'block'; msg.style.background = '#d1fae5'; msg.style.color = '#065f46';
      }
    }).catch(function (e) {
      if (msg) {
        msg.textContent = (e.code === 'auth/too-many-requests')
          ? '送信回数が多すぎます。しばらくしてからお試しください。'
          : 'エラー: ' + (e.message || '');
        msg.style.display = 'block'; msg.style.background = '#fee2e2'; msg.style.color = '#991b1b';
      }
    });
  }

  function handleDeleteAccount() {
    var confirm = document.getElementById('deleteConfirm');
    var section = document.getElementById('deleteAccountSection');
    if (confirm) confirm.style.display = '';
    if (section) section.style.display = 'none';
  }

  function cancelDeleteAccount() {
    var confirm = document.getElementById('deleteConfirm');
    var section = document.getElementById('deleteAccountSection');
    if (confirm) confirm.style.display = 'none';
    if (section) section.style.display = '';
  }

  function confirmDeleteAccount() {
    var msg = document.getElementById('myPageMsg');
    window.firebaseDeleteAccount().then(function () {
      closeModal('myPage');
      window.location.reload();
    }).catch(function (e) {
      if (msg) {
        if (e.code === 'auth/requires-recent-login') {
          msg.textContent = 'セキュリティのため再ログインが必要です。一度ログアウトして再度ログインしてからお試しください。';
        } else {
          msg.textContent = '退会処理に失敗しました: ' + (e.message || '');
        }
        msg.style.display = 'block'; msg.style.background = '#fee2e2'; msg.style.color = '#991b1b';
      }
      var confirmEl = document.getElementById('deleteConfirm');
      var section = document.getElementById('deleteAccountSection');
      if (confirmEl) confirmEl.style.display = 'none';
      if (section) section.style.display = '';
    });
  }

  // ===== HANDLER MAP =====
  var handlers = {
    openModal: function (arg) { openModal(arg); },
    closeModal: function (arg) { closeModal(arg); },
    switchModal: function (arg) { switchModal(arg); },
    switchToLogin: function () { switchToLogin(); },
    handleGoogleLogin: function (arg) { handleGoogleLogin(arg); },
    handleLogin: function (arg, e) { handleLogin(e); },
    handleRegister: function (arg, e) { handleRegister(e); },
    handleResetPassword: function (arg, e) { handleResetPassword(e); },
    handleCompleteProfile: function (arg, e) { handleCompleteProfile(e); },
    handleLogout: function () { handleLogout(); },
    openMyPage: function () { openMyPage(); },
    handleMyPageResetPassword: function () { handleMyPageResetPassword(); },
    showSetPasswordForm: function () { showSetPasswordForm(); },
    handleSetPassword: function () { handleSetPassword(); },
    handleResendVerification: function () { handleResendVerification(); },
    handleDeleteAccount: function () { handleDeleteAccount(); },
    cancelDeleteAccount: function () { cancelDeleteAccount(); },
    confirmDeleteAccount: function () { confirmDeleteAccount(); },
    toggleAffiliationCode: function () { toggleAffiliationCode(); },
    toggleCpAffiliationCode: function () { toggleCpAffiliationCode(); }
  };

  // ===== EVENT DELEGATION =====
  function init() {
    // Modal overlay click-to-close
    document.querySelectorAll('.modal-overlay').forEach(function (o) {
      o.addEventListener('click', function (e) {
        if (e.target === o) {
          o.classList.remove('show');
          document.body.style.overflow = '';
        }
      });
    });

    // Click delegation: data-auth-action="actionName" or "actionName:arg"
    document.addEventListener('click', function (e) {
      var el = e.target.closest('[data-auth-action]');
      if (!el) return;
      var attr = el.getAttribute('data-auth-action');
      var colonIdx = attr.indexOf(':');
      var action = colonIdx > -1 ? attr.substring(0, colonIdx) : attr;
      var arg = colonIdx > -1 ? attr.substring(colonIdx + 1) : '';
      if (handlers[action]) {
        e.preventDefault();
        handlers[action](arg, e);
      }
    }, false);

    // Form submit delegation: data-auth-submit="handlerName"
    document.querySelectorAll('[data-auth-submit]').forEach(function (form) {
      form.addEventListener('submit', function (e) {
        var action = form.getAttribute('data-auth-submit');
        if (handlers[action]) {
          handlers[action]('', e);
        }
      }, false);
    });

    // Change delegation: data-auth-change="handlerName"
    document.addEventListener('change', function (e) {
      var el = e.target.closest('[data-auth-change]');
      if (!el) return;
      var action = el.getAttribute('data-auth-change');
      if (handlers[action]) {
        handlers[action]('', e);
      }
    }, false);
  }

  // ===== EXPOSE ON WINDOW =====
  // Firebase module の onAuthStateChanged → updateAuthUI から呼ばれるため window に公開
  window.openModal = openModal;
  window.closeModal = closeModal;
  window.switchModal = switchModal;
  window.switchToLogin = switchToLogin;
  window.handleGoogleLogin = handleGoogleLogin;
  window.handleLogin = handleLogin;
  window.handleRegister = handleRegister;
  window.handleResetPassword = handleResetPassword;
  window.handleCompleteProfile = handleCompleteProfile;
  window.handleLogout = handleLogout;
  window.openMyPage = openMyPage;
  window.handleMyPageResetPassword = handleMyPageResetPassword;
  window.showSetPasswordForm = showSetPasswordForm;
  window.handleSetPassword = handleSetPassword;
  window.handleResendVerification = handleResendVerification;
  window.handleDeleteAccount = handleDeleteAccount;
  window.cancelDeleteAccount = cancelDeleteAccount;
  window.confirmDeleteAccount = confirmDeleteAccount;
  window.toggleAffiliationCode = toggleAffiliationCode;
  window.AFFILIATION_LABELS = AFFILIATION_LABELS;
  window.JOBTITLE_LABELS = JOBTITLE_LABELS;

  // ===== INIT =====
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, false);
  } else {
    init();
  }
})();
