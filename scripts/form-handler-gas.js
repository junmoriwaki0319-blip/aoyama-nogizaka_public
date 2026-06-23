/**
 * ====================================================================
 *  青山乃木坂パートナーズ — フォーム受信・ニュースレター配信 GAS
 * ====================================================================
 *
 *  ■ プロジェクト名: ANP_フォーム受信
 *  ■ デプロイ済みURL:
 *    https://script.google.com/macros/s/AKfycbwQby968Ts6rSFyzTZZFfblfFDYOogRPWYa7kYl5eHdXaemLIilu6FnOemvX5ygIBwt/exec
 *
 *  ■ 機能一覧:
 *    1. お問い合わせフォーム受信   (index.html #contact)
 *    2. ニュースレター登録         (news/index.html)
 *    3. 相談フォーム受信           (news/activist-report.html)
 *    4. ニュースレター一斉配信     (スプレッドシートメニューから実行)
 *    5. 配信解除                   (メール内URLクリックで自動処理)
 *    6. 月次サマリー自動生成       (スプレッドシートメニューから実行)
 *
 *  ■ データ保存先: Google Drive「ANP_フォーム受信データ」スプレッドシート
 *    - 「お問い合わせ」シート      … 問い合わせ内容を蓄積
 *    - 「ニュースレター」シート    … 現在の有効な登録者一覧
 *    - 「相談フォーム」シート      … 相談内容を蓄積
 *    - 「アクティビティログ」シート … 全イベント(登録/解除/配信/問合せ等)の時系列記録
 *    - 「配信履歴」シート          … 一斉配信の実行記録
 *    - 「月次サマリー」シート      … 月別の登録数・解除数・配信数の集計
 *
 *  ■ 通知先: jun.moriwaki@aoyama-nogizaka.com
 *
 *  ■ セキュリティ対策（2026追加）:
 *    - 全入力フィールドのサニタイズ（HTMLタグ無害化・制御文字除去・長さ制限）
 *    - XSS/SQLi等の攻撃ペイロード自動検知 → ブロック＆「セキュリティログ」記録
 *    - メールアドレス厳格検証（ヘッダーインジェクション対策）
 *    - ハニーポット（隠しフィールド company_url）によるボット遮断
 *    - 同一メールからの連続送信レート制限（10分で5件まで）
 *    - 配信解除ページの出力エスケープ（反射型XSS対策）
 *    ※ これらを有効化するには、下記手順でGASの「新しいバージョン」を再デプロイしてください。
 *
 *  ■ コード更新手順:
 *    1. このファイルの内容をGASエディタに貼り付け → 保存
 *    2.「デプロイ」→「デプロイを管理」→ 鉛筆アイコン
 *    3. バージョンを「新しいバージョン」→「デプロイ」
 *    ※ URLは変わりません。HTML側の修正は不要です。
 *
 * ====================================================================
 */

// === 設定 ===
var NOTIFY_EMAIL = 'jun.moriwaki@aoyama-nogizaka.com';
var SPREADSHEET_ID = '';

// === セキュリティ設定 ===
// 各フィールドの最大文字数（超過分は切り捨て）
var MAX_LEN = { name: 100, company: 150, department: 100, furigana: 100, email: 254, message: 5000, referral: 50 };
// ハニーポット（ボットが入力してしまう隠しフィールド名）。HTML側で CSS 非表示にしている。
var HONEYPOT_FIELD = 'company_url';
// 同一メールアドレスからの連続送信を制限（件数 / 秒）
var RATE_LIMIT_MAX = 5;
var RATE_LIMIT_WINDOW_SEC = 600; // 10分

// ====================================================================
//  共通ユーティリティ
// ====================================================================

function getOrCreateSheet(sheetName) {
  var ss;
  if (SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    var files = DriveApp.getFilesByName('ANP_フォーム受信データ');
    ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create('ANP_フォーム受信データ');
  }
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getUnsubscribeUrl(email) {
  return ScriptApp.getService().getUrl() + '?action=unsubscribe&email=' + encodeURIComponent(email);
}

function timestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function yearMonth() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
}

function initHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a2d4f').setFontColor('#fff').setFontWeight('bold');
  }
}

/** 全イベントを時系列で記録するアクティビティログ */
function logActivity(email, action, detail) {
  var sheet = getOrCreateSheet('アクティビティログ');
  initHeader(sheet, ['日時', '年月', 'メールアドレス', 'アクション', '詳細']);
  sheet.appendRow([timestamp(), yearMonth(), email, action, detail || '']);
}

// ====================================================================
//  セキュリティ・ユーティリティ
// ====================================================================

/** HTML特殊文字をエスケープ（HtmlService出力の反射型XSS対策） */
function escapeHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * テキスト入力の無害化。
 *  - 文字列化・前後空白除去・制御文字除去
 *  - HTMLタグ／山括弧を全角化して無害化（メール本文・シート保存用）
 *  - 最大文字数で切り捨て
 */
function sanitizeText(str, maxLen) {
  var s = String(str == null ? '' : str);
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, ''); // 制御文字除去（改行/タブは許可）
  s = s.replace(/</g, '＜').replace(/>/g, '＞'); // タグ構文を全角化して無害化
  s = s.trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

/** メールアドレスの厳格な検証（ヘッダーインジェクション対策含む） */
function isValidEmail(email) {
  var s = String(email == null ? '' : email).trim();
  if (s.length === 0 || s.length > MAX_LEN.email) return false;
  if (/[\r\n\t,;<>"'\\]/.test(s)) return false;       // 改行・区切り・引用符を拒否
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

/**
 * 攻撃／脆弱性診断のペイロードを検出。
 * 1つでも一致したら不審な送信として扱う。
 */
function looksMalicious(values) {
  var joined = values.join(' \n ').toLowerCase();
  var patterns = [
    /<\s*script/, /<\s*\/\s*script/, /<\s*img/, /<\s*svg/, /<\s*iframe/, /<\s*object/, /<\s*embed/,
    /javascript:/, /vbscript:/, /data:text\/html/,
    /on\w+\s*=/,                       // onerror= onload= onclick= 等
    /alert\s*\(/, /prompt\s*\(/, /confirm\s*\(/, /eval\s*\(/, /document\.(cookie|location|domain)/, /window\.location/,
    /<\s*[a-z][^>]*>/,                 // 任意のHTMLタグ
    /\{\{.*\}\}/, /\$\{.*\}/,          // テンプレートインジェクション
    /union\s+select/, /'\s*or\s*'1'\s*=\s*'1/, /;\s*drop\s+table/, /--\s*$/, // SQLi系
    /\.\.\/\.\.\//,                    // パストラバーサル
    /\$\(.*\)/                         // 簡易コマンド/jQuery風
  ];
  for (var i = 0; i < patterns.length; i++) {
    if (patterns[i].test(joined)) return true;
  }
  return false;
}

/** 不審な送信をセキュリティログに記録し、管理者へ通知（自動返信はしない） */
function logSecurityEvent(formType, email, rawData, reason) {
  var sheet = getOrCreateSheet('セキュリティログ');
  initHeader(sheet, ['日時', 'フォーム種別', '理由', 'メール(申告値)', '生データ(先頭500字)']);
  var raw = '';
  try { raw = JSON.stringify(rawData).slice(0, 500); } catch (e) { raw = '(解析不可)'; }
  sheet.appendRow([timestamp(), formType || '', reason || '', String(email || '').slice(0, 254), raw]);
  try {
    GmailApp.sendEmail(NOTIFY_EMAIL,
      '【警告】不審なフォーム送信をブロックしました',
      '自動的にブロックされた送信があります。対応は不要です（記録目的の通知）。\n\n' +
      '■ 種別: ' + (formType || '不明') + '\n' +
      '■ 理由: ' + (reason || '') + '\n' +
      '■ 申告メール: ' + String(email || '') + '\n' +
      '■ 受信日時: ' + timestamp() + '\n\n' +
      '詳細はスプレッドシート「セキュリティログ」をご確認ください。');
  } catch (e) { /* 通知失敗は無視 */ }
}

/** 同一メールからの連続送信をレート制限。true=制限超過 */
function isRateLimited(email) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'rl_' + Utilities.base64EncodeWebSafe(String(email || 'anon'));
    var count = parseInt(cache.get(key) || '0', 10) + 1;
    cache.put(key, String(count), RATE_LIMIT_WINDOW_SEC);
    return count > RATE_LIMIT_MAX;
  } catch (e) {
    return false; // キャッシュ障害時は通常処理を継続
  }
}

/** メール件名用に改行・制御文字を除去（ヘッダーインジェクション対策） */
function safeSubject(str, maxLen) {
  var s = String(str == null ? '' : str)
    .replace(/[\r\n]+/g, ' ')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
    .trim();
  return maxLen && s.length > maxLen ? s.slice(0, maxLen) : s;
}

// ====================================================================
//  POST受信: フォーム送信の振り分け
// ====================================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, message: '不正なリクエスト' });
    }
    var data = JSON.parse(e.postData.contents);
    var formType = data._formType;

    // --- 1. ハニーポット: ボットが隠しフィールドを埋めた場合は静かに破棄 ---
    if (data[HONEYPOT_FIELD]) {
      logSecurityEvent(formType, data.email, data, 'ハニーポット検知');
      return jsonResponse({ success: true }); // 攻撃者には成功を装う
    }

    // --- 2. 攻撃／脆弱性診断ペイロード検知: ブロックして記録（自動返信なし） ---
    var allValues = [];
    for (var k in data) { if (k.charAt(0) !== '_') allValues.push(String(data[k])); }
    if (looksMalicious(allValues)) {
      logSecurityEvent(formType, data.email, data, '不正ペイロード検知（XSS/インジェクション疑い）');
      return jsonResponse({ success: true }); // プローブにシグナルを与えない
    }

    // --- 3. メール検証（全フォーム共通の必須項目）---
    if (!isValidEmail(data.email)) {
      return jsonResponse({ success: false, message: 'メールアドレスの形式が正しくありません。' });
    }

    // --- 4. レート制限 ---
    if (isRateLimited(data.email)) {
      logSecurityEvent(formType, data.email, data, 'レート制限超過');
      return jsonResponse({ success: false, message: '送信回数が上限に達しました。しばらく時間をおいてお試しください。' });
    }

    // --- 5. 振り分け ---
    switch (formType) {
      case 'contact':      return handleContact(data);
      case 'newsletter':   return handleNewsletter(data);
      case 'consultation': return handleConsultation(data);
      default:             return jsonResponse({ success: false, message: '不明なフォーム種別' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: '送信処理に失敗しました。' });
  }
}

// ====================================================================
//  GET受信: 配信解除
// ====================================================================

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var email  = e && e.parameter && e.parameter.email;
  // 不正な形式のメールは配信解除処理に渡さない（反射型XSS・インジェクション対策）
  if (action === 'unsubscribe' && isValidEmail(email)) return handleUnsubscribe(email);
  return jsonResponse({ status: 'ok' });
}

function handleUnsubscribe(email) {
  var sheet = getOrCreateSheet('ニュースレター');
  if (sheet.getLastRow() <= 1) return unsubPage('登録が見つかりませんでした。', false);

  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  var removed = false;
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === email) { sheet.deleteRow(i + 2); removed = true; }
  }
  if (!removed) return unsubPage('登録が見つかりませんでした。', false);

  logActivity(email, '配信解除', 'メール内リンクから解除');
  GmailApp.sendEmail(NOTIFY_EMAIL, '【配信解除】' + email, 'メール: ' + email + '\n解除日時: ' + timestamp());
  return unsubPage(email + ' の配信を解除しました。', true);
}

function unsubPage(message, success) {
  var color = success ? '#2d7a4f' : '#b53a3a';
  var title = success ? '配信解除完了' : 'エラー';
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>配信解除 | 青山乃木坂パートナーズ</title>' +
    '<style>body{font-family:-apple-system,sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#f0eee9;}' +
    '.card{background:#fff;padding:48px 40px;max-width:440px;text-align:center;border-top:3px solid ' + color + ';}' +
    'h2{color:#1a2d4f;font-size:18px;margin-bottom:16px;}p{color:#444;font-size:14px;line-height:1.8;}a{color:#9b8b6e;}</style></head>' +
    '<body><div class="card"><h2>' + escapeHtml(title) + '</h2><p>' + escapeHtml(message) + '</p>' +
    '<p style="margin-top:24px;"><a href="https://aoyama-nogizaka.com">トップページへ戻る</a></p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('配信解除');
}

// ====================================================================
//  1. お問い合わせフォーム (index.html #contact)
// ====================================================================

function handleContact(d) {
  // 全フィールドを無害化（HTMLタグ全角化・制御文字除去・長さ制限）
  var company    = sanitizeText(d.company, MAX_LEN.company);
  var department = sanitizeText(d.department, MAX_LEN.department);
  var name       = sanitizeText(d.name, MAX_LEN.name);
  var furigana   = sanitizeText(d.furigana, MAX_LEN.furigana);
  var email      = String(d.email).trim();
  var message    = sanitizeText(d.message, MAX_LEN.message);
  var referral   = sanitizeText(d.referral, MAX_LEN.referral);

  if (!company || !department || !name || !message) {
    return jsonResponse({ success: false, message: '必須項目が未入力です。' });
  }

  var sheet = getOrCreateSheet('お問い合わせ');
  initHeader(sheet, ['受信日時', '貴社名', '部署・役職', 'お名前', 'フリガナ', 'メールアドレス', 'お問い合わせ内容', '流入経路']);
  var now = timestamp();
  sheet.appendRow([now, company, department, name, furigana, email, message, referral]);

  logActivity(email, 'お問い合わせ', company + ' ' + name);

  GmailApp.sendEmail(NOTIFY_EMAIL,
    safeSubject('【お問い合わせ】' + company + ' ' + name + ' 様', 200),
    '新しいお問い合わせがありました。\n\n' +
    '■ 貴社名: ' + company + '\n' +
    '■ 部署・役職: ' + department + '\n' +
    '■ お名前: ' + name + '\n' +
    '■ フリガナ: ' + (furigana || '未入力') + '\n' +
    '■ メール: ' + email + '\n' +
    '■ 流入経路: ' + (referral || '未選択') + '\n\n' +
    '■ お問い合わせ内容:\n' + message + '\n\n' +
    '─────────────────────\n受信日時: ' + now);

  GmailApp.sendEmail(email,
    '【青山乃木坂パートナーズ】お問い合わせを受け付けました',
    name + ' 様\n\n' +
    'この度はお問い合わせいただき、誠にありがとうございます。\n' +
    '以下の内容で受け付けいたしました。\n' +
    '担当者より2営業日以内にご連絡差し上げます。\n\n' +
    '─────────────────────\n' +
    '■ 貴社名: ' + company + '\n' +
    '■ 部署・役職: ' + department + '\n' +
    '■ お名前: ' + name + '\n' +
    '■ お問い合わせ内容:\n' + message + '\n' +
    '─────────────────────\n\n' +
    '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n',
    { name: '青山乃木坂パートナーズ' });

  return jsonResponse({ success: true });
}

// ====================================================================
//  2. ニュースレター登録 (news/index.html)
// ====================================================================

function handleNewsletter(d) {
  var email = String(d.email).trim(); // doPostでisValidEmail検証済み
  var sheet = getOrCreateSheet('ニュースレター');
  initHeader(sheet, ['登録日時', 'メールアドレス']);
  var now = timestamp();

  var existing = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  if (existing.includes(email)) return jsonResponse({ success: true, message: 'already_registered' });

  sheet.appendRow([now, email]);
  logActivity(email, 'ニュースレター登録', '');

  GmailApp.sendEmail(NOTIFY_EMAIL, safeSubject('【ニュースレター登録】' + email, 200),
    '新規ニュースレター登録:\n\nメール: ' + email + '\n登録日時: ' + now);

  GmailApp.sendEmail(email,
    '【青山乃木坂パートナーズ】ニュースレター登録完了',
    'ニュースレターへのご登録ありがとうございます。\n\n' +
    '今後、新しい論考・プレスリリース発行時にメールでお知らせいたします。\n\n' +
    '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n\n' +
    '※ 配信停止はこちら: ' + getUnsubscribeUrl(email) + '\n',
    { name: '青山乃木坂パートナーズ' });

  return jsonResponse({ success: true });
}

// ====================================================================
//  3. 相談フォーム (news/activist-report.html)
// ====================================================================

function handleConsultation(d) {
  var name       = sanitizeText(d.name, MAX_LEN.name);
  var company    = sanitizeText(d.company, MAX_LEN.company);
  var department = sanitizeText(d.department, MAX_LEN.department);
  var email      = String(d.email).trim();
  var message    = sanitizeText(d.message, MAX_LEN.message);
  var referral   = sanitizeText(d.referral, MAX_LEN.referral);

  if (!name || !company || !department) {
    return jsonResponse({ success: false, message: '必須項目が未入力です。' });
  }

  var sheet = getOrCreateSheet('相談フォーム');
  initHeader(sheet, ['受信日時', '氏名', '会社名', '部署・役職', 'メールアドレス', 'ご相談内容', '流入経路']);
  var now = timestamp();
  sheet.appendRow([now, name, company, department, email, message, referral]);

  logActivity(email, 'ご相談', company + ' ' + name);

  GmailApp.sendEmail(NOTIFY_EMAIL,
    safeSubject('【ご相談】' + company + ' ' + name + ' 様', 200),
    '新しいご相談がありました。\n\n' +
    '■ 氏名: ' + name + '\n' +
    '■ 会社名: ' + company + '\n' +
    '■ 部署・役職: ' + department + '\n' +
    '■ メール: ' + email + '\n' +
    '■ 流入経路: ' + (referral || '未選択') + '\n\n' +
    '■ ご相談内容:\n' + (message || '未入力') + '\n\n' +
    '受信日時: ' + now);

  GmailApp.sendEmail(email,
    '【青山乃木坂パートナーズ】ご相談を受け付けました',
    name + ' 様\n\n' +
    'この度はご相談いただき、誠にありがとうございます。\n' +
    '担当者より2営業日以内にご連絡差し上げます。\n\n' +
    '─────────────────────\n' +
    '■ 氏名: ' + name + '\n' +
    '■ 会社名: ' + company + '\n' +
    '■ ご相談内容:\n' + (message || '未入力') + '\n' +
    '─────────────────────\n\n' +
    '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n',
    { name: '青山乃木坂パートナーズ' });

  return jsonResponse({ success: true });
}

// ====================================================================
//  4. ニュースレター一斉配信 (スプレッドシートメニューから実行)
// ====================================================================

function sendNewsletter() {
  var ui = SpreadsheetApp.getUi();

  var subjectRes = ui.prompt('ニュースレター配信', '件名を入力してください\n（例: 【青山乃木坂パートナーズ】新しい論考を公開しました）', ui.ButtonSet.OK_CANCEL);
  if (subjectRes.getSelectedButton() !== ui.Button.OK) return;
  var subject = subjectRes.getResponseText().trim();
  if (!subject) { ui.alert('件名が空です。'); return; }

  var bodyRes = ui.prompt('ニュースレター配信', '本文を入力してください\n（URLなどを含めてください）', ui.ButtonSet.OK_CANCEL);
  if (bodyRes.getSelectedButton() !== ui.Button.OK) return;
  var bodyText = bodyRes.getResponseText().trim();
  if (!bodyText) { ui.alert('本文が空です。'); return; }

  var sheet = getOrCreateSheet('ニュースレター');
  if (sheet.getLastRow() <= 1) { ui.alert('登録者がいません。'); return; }
  var emails = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat().filter(function(e) { return e && e.includes('@'); });
  var uniqueEmails = [...new Set(emails)];

  var confirm = ui.alert('配信確認', uniqueEmails.length + '件のアドレスに送信します。\n\n件名: ' + subject + '\n\nよろしいですか？', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var successCount = 0, failCount = 0;
  var errors = [];

  uniqueEmails.forEach(function(email) {
    try {
      var footer = '\n\n─────────────────────\n' +
        '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n\n' +
        '※ このメールはニュースレター登録者にお送りしています。\n' +
        '配信停止: ' + getUnsubscribeUrl(email);
      GmailApp.sendEmail(email, subject, bodyText + footer, { name: '青山乃木坂パートナーズ', replyTo: NOTIFY_EMAIL });
      successCount++;
    } catch (e) {
      failCount++;
      errors.push(email + ': ' + e.toString());
    }
  });

  logActivity(NOTIFY_EMAIL, 'ニュースレター配信', '件名: ' + subject + ' / 成功: ' + successCount + ' / 失敗: ' + failCount);

  var hist = getOrCreateSheet('配信履歴');
  initHeader(hist, ['配信日時', '件名', '送信成功', '送信失敗', '登録者数']);
  var now = timestamp();
  hist.appendRow([now, subject, successCount, failCount, uniqueEmails.length]);

  var msg = '配信完了\n\n成功: ' + successCount + '件\n失敗: ' + failCount + '件';
  if (errors.length > 0) msg += '\n\nエラー:\n' + errors.join('\n');
  ui.alert(msg);

  GmailApp.sendEmail(NOTIFY_EMAIL, '【配信レポート】' + subject,
    'ニュースレター配信レポート\n\n' +
    '■ 件名: ' + subject + '\n' +
    '■ 配信日時: ' + now + '\n' +
    '■ 成功: ' + successCount + '件\n' +
    '■ 失敗: ' + failCount + '件\n' +
    '■ 登録者数: ' + uniqueEmails.length + '件\n' +
    (errors.length > 0 ? '\nエラー:\n' + errors.join('\n') : ''));
}

// ====================================================================
//  5. 月次サマリー生成 (スプレッドシートメニューから実行)
// ====================================================================

function generateMonthlySummary() {
  var logSheet = getOrCreateSheet('アクティビティログ');
  if (logSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('ログデータがありません。');
    return;
  }

  var logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
  var monthly = {};

  logs.forEach(function(row) {
    var ym = row[1];
    var action = row[3];
    if (!ym) return;
    if (!monthly[ym]) monthly[ym] = { 登録: 0, 解除: 0, 配信: 0, 問合せ: 0, 相談: 0 };
    if (action === 'ニュースレター登録') monthly[ym]['登録']++;
    else if (action === '配信解除') monthly[ym]['解除']++;
    else if (action === 'ニュースレター配信') monthly[ym]['配信']++;
    else if (action === 'お問い合わせ') monthly[ym]['問合せ']++;
    else if (action === 'ご相談') monthly[ym]['相談']++;
  });

  var summarySheet = getOrCreateSheet('月次サマリー');
  summarySheet.clearContents();
  var headers = ['年月', 'NL登録', 'NL解除', 'NL配信', 'お問い合わせ', 'ご相談'];
  summarySheet.appendRow(headers);
  summarySheet.getRange(1, 1, 1, headers.length).setBackground('#1a2d4f').setFontColor('#fff').setFontWeight('bold');

  var months = Object.keys(monthly).sort();
  months.forEach(function(ym) {
    var m = monthly[ym];
    summarySheet.appendRow([ym, m['登録'], m['解除'], m['配信'], m['問合せ'], m['相談']]);
  });

  // 現在の有効登録者数を末尾に追加
  var nlSheet = getOrCreateSheet('ニュースレター');
  var activeCount = nlSheet.getLastRow() <= 1 ? 0 : nlSheet.getLastRow() - 1;
  summarySheet.appendRow([]);
  summarySheet.appendRow(['現在の有効登録者数', activeCount]);
  summarySheet.getRange(summarySheet.getLastRow(), 1, 1, 2).setFontWeight('bold');

  summarySheet.autoResizeColumn(1);
  SpreadsheetApp.getUi().alert('月次サマリーを更新しました。\n\n現在の有効登録者数: ' + activeCount + '件');
}

// ====================================================================
//  6. スプレッドシートメニュー
// ====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ニュースレター')
    .addItem('配信する', 'sendNewsletter')
    .addItem('登録者数を確認', 'showSubscriberCount')
    .addSeparator()
    .addItem('月次サマリーを生成', 'generateMonthlySummary')
    .addToUi();
}

function showSubscriberCount() {
  var sheet = getOrCreateSheet('ニュースレター');
  var count = sheet.getLastRow() <= 1 ? 0 : sheet.getLastRow() - 1;
  SpreadsheetApp.getUi().alert('現在の登録者数: ' + count + '件');
}
