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
//  POST受信: フォーム送信の振り分け
// ====================================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data._formType) {
      case 'contact':      return handleContact(data);
      case 'newsletter':   return handleNewsletter(data);
      case 'consultation': return handleConsultation(data);
      default:             return jsonResponse({ success: false, message: '不明なフォーム種別' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

// ====================================================================
//  GET受信: 配信解除
// ====================================================================

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var email  = e && e.parameter && e.parameter.email;
  if (action === 'unsubscribe' && email) return handleUnsubscribe(email);
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
    '<body><div class="card"><h2>' + title + '</h2><p>' + message + '</p>' +
    '<p style="margin-top:24px;"><a href="https://aoyama-nogizaka.com">トップページへ戻る</a></p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('配信解除');
}

// ====================================================================
//  1. お問い合わせフォーム (index.html #contact)
// ====================================================================

function handleContact(d) {
  var sheet = getOrCreateSheet('お問い合わせ');
  initHeader(sheet, ['受信日時', '貴社名', '部署・役職', 'お名前', 'フリガナ', 'メールアドレス', 'お問い合わせ内容', '流入経路']);
  var now = timestamp();
  sheet.appendRow([now, d.company, d.department, d.name, d.furigana || '', d.email, d.message, d.referral || '']);

  logActivity(d.email, 'お問い合わせ', d.company + ' ' + d.name);

  GmailApp.sendEmail(NOTIFY_EMAIL,
    '【お問い合わせ】' + d.company + ' ' + d.name + ' 様',
    '新しいお問い合わせがありました。\n\n' +
    '■ 貴社名: ' + d.company + '\n' +
    '■ 部署・役職: ' + d.department + '\n' +
    '■ お名前: ' + d.name + '\n' +
    '■ フリガナ: ' + (d.furigana || '未入力') + '\n' +
    '■ メール: ' + d.email + '\n' +
    '■ 流入経路: ' + (d.referral || '未選択') + '\n\n' +
    '■ お問い合わせ内容:\n' + d.message + '\n\n' +
    '─────────────────────\n受信日時: ' + now);

  GmailApp.sendEmail(d.email,
    '【青山乃木坂パートナーズ】お問い合わせを受け付けました',
    d.name + ' 様\n\n' +
    'この度はお問い合わせいただき、誠にありがとうございます。\n' +
    '以下の内容で受け付けいたしました。\n' +
    '担当者より2営業日以内にご連絡差し上げます。\n\n' +
    '─────────────────────\n' +
    '■ 貴社名: ' + d.company + '\n' +
    '■ 部署・役職: ' + d.department + '\n' +
    '■ お名前: ' + d.name + '\n' +
    '■ お問い合わせ内容:\n' + d.message + '\n' +
    '─────────────────────\n\n' +
    '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n',
    { name: '青山乃木坂パートナーズ' });

  return jsonResponse({ success: true });
}

// ====================================================================
//  2. ニュースレター登録 (news/index.html)
// ====================================================================

function handleNewsletter(d) {
  var sheet = getOrCreateSheet('ニュースレター');
  initHeader(sheet, ['登録日時', 'メールアドレス']);
  var now = timestamp();

  var existing = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  if (existing.includes(d.email)) return jsonResponse({ success: true, message: 'already_registered' });

  sheet.appendRow([now, d.email]);
  logActivity(d.email, 'ニュースレター登録', '');

  GmailApp.sendEmail(NOTIFY_EMAIL, '【ニュースレター登録】' + d.email,
    '新規ニュースレター登録:\n\nメール: ' + d.email + '\n登録日時: ' + now);

  GmailApp.sendEmail(d.email,
    '【青山乃木坂パートナーズ】ニュースレター登録完了',
    'ニュースレターへのご登録ありがとうございます。\n\n' +
    '今後、新しい論考・プレスリリース発行時にメールでお知らせいたします。\n\n' +
    '青山乃木坂パートナーズ合同会社\nhttps://aoyama-nogizaka.com\n\n' +
    '※ 配信停止はこちら: ' + getUnsubscribeUrl(d.email) + '\n',
    { name: '青山乃木坂パートナーズ' });

  return jsonResponse({ success: true });
}

// ====================================================================
//  3. 相談フォーム (news/activist-report.html)
// ====================================================================

function handleConsultation(d) {
  var sheet = getOrCreateSheet('相談フォーム');
  initHeader(sheet, ['受信日時', '氏名', '会社名', '部署・役職', 'メールアドレス', 'ご相談内容', '流入経路']);
  var now = timestamp();
  sheet.appendRow([now, d.name, d.company, d.department, d.email, d.message || '', d.referral || '']);

  logActivity(d.email, 'ご相談', d.company + ' ' + d.name);

  GmailApp.sendEmail(NOTIFY_EMAIL,
    '【ご相談】' + d.company + ' ' + d.name + ' 様',
    '新しいご相談がありました。\n\n' +
    '■ 氏名: ' + d.name + '\n' +
    '■ 会社名: ' + d.company + '\n' +
    '■ 部署・役職: ' + d.department + '\n' +
    '■ メール: ' + d.email + '\n' +
    '■ 流入経路: ' + (d.referral || '未選択') + '\n\n' +
    '■ ご相談内容:\n' + (d.message || '未入力') + '\n\n' +
    '受信日時: ' + now);

  GmailApp.sendEmail(d.email,
    '【青山乃木坂パートナーズ】ご相談を受け付けました',
    d.name + ' 様\n\n' +
    'この度はご相談いただき、誠にありがとうございます。\n' +
    '担当者より2営業日以内にご連絡差し上げます。\n\n' +
    '─────────────────────\n' +
    '■ 氏名: ' + d.name + '\n' +
    '■ 会社名: ' + d.company + '\n' +
    '■ ご相談内容:\n' + (d.message || '未入力') + '\n' +
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
