/**
 * Google Apps Script: Firestore → Google Sheets 会員情報エクスポート
 *
 * === セットアップ手順 ===
 * 1. Google Sheets を新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付け
 * 4. FirestoreApp ライブラリを追加:
 *    スクリプトID: 1VUSl4b1r1eoNcRWotZM3e87ygkxvXltOgyDZhixqncz9lQ3MjfT1iKFw
 *    バージョン: 最新を選択
 * 5. FIREBASE_PROJECT_ID を確認（下記の定数）
 * 6. 初回実行時に Google アカウントの認証を許可
 * 7. トリガー設定（任意）: 時計アイコン → トリガーを追加 → exportUsersToSheet → 毎日/毎週
 */

// === 設定 ===
const FIREBASE_PROJECT_ID = 'aoyama-nogizaka-activist';
const SHEET_NAME = '会員一覧';

/**
 * メイン関数: Firestore の users コレクションをスプレッドシートに出力
 */
function exportUsersToSheet() {
  const firestore = FirestoreApp.getFirestore(FIREBASE_PROJECT_ID);
  const users = firestore.getDocuments('users');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ヘッダー行
  const headers = [
    'UID',
    '名前',
    'メールアドレス',
    '会社名・組織名',
    '所属カテゴリ',
    '証券コード',
    '役職',
    'プラン',
    '登録日',
    '最終更新日'
  ];

  // データ行を作成
  const rows = users.map(doc => {
    const d = doc.obj;
    const uid = doc.name ? doc.name.split('/').pop() : '';
    return [
      uid,
      d.name || '',
      d.email || '',
      d.company || '',
      formatAffiliation(d.affiliation || ''),
      d.affiliationCode || '',
      formatJobTitle(d.jobTitle || ''),
      d.plan || 'free',
      formatTimestamp(d.createdAt),
      formatTimestamp(d.updatedAt)
    ];
  });

  // シートをクリアして書き込み
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // ヘッダー行のスタイリング
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a2d4f');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の自動調整
  headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));

  // フィルター設定
  if (rows.length > 0) {
    const dataRange = sheet.getRange(1, 1, rows.length + 1, headers.length);
    if (sheet.getFilter()) sheet.getFilter().remove();
    dataRange.createFilter();
  }

  // サマリーシートの更新
  updateSummary(ss, rows);

  Logger.log(`${rows.length} 件の会員データをエクスポートしました`);
  return rows.length;
}

/**
 * サマリー情報を別シートに出力
 */
function updateSummary(ss, rows) {
  const summaryName = 'サマリー';
  let summary = ss.getSheetByName(summaryName);
  if (!summary) {
    summary = ss.insertSheet(summaryName);
  }
  summary.clearContents();

  const now = new Date();
  const total = rows.length;

  // 所属別カウント
  const affiliationCount = {};
  const jobTitleCount = {};
  rows.forEach(r => {
    const aff = r[4] || '未設定';
    const job = r[6] || '未設定';
    affiliationCount[aff] = (affiliationCount[aff] || 0) + 1;
    jobTitleCount[job] = (jobTitleCount[job] || 0) + 1;
  });

  const data = [
    ['最終更新', Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')],
    ['総会員数', total],
    [''],
    ['【所属別】', '人数'],
    ...Object.entries(affiliationCount).sort((a, b) => b[1] - a[1]),
    [''],
    ['【役職別】', '人数'],
    ...Object.entries(jobTitleCount).sort((a, b) => b[1] - a[1])
  ];

  summary.getRange(1, 1, data.length, 2).setValues(data.map(r => Array.isArray(r) && r.length === 2 ? r : [r[0] || '', r[1] || '']));

  // ヘッダースタイリング
  summary.getRange(1, 1, 2, 1).setFontWeight('bold');
  summary.getRange(4, 1, 1, 2).setBackground('#1a2d4f').setFontColor('#ffffff').setFontWeight('bold');
  summary.autoResizeColumn(1);
  summary.autoResizeColumn(2);
}

/**
 * 所属カテゴリのラベル変換
 */
function formatAffiliation(code) {
  const map = {
    'listed_company': '上場企業',
    'institutional_investor': '機関投資家',
    'individual_investor': '個人投資家',
    'consulting': 'コンサルティングファーム',
    'legal': '法律事務所・弁護士',
    'media': 'メディア・報道機関',
    'academic': '学術・研究機関',
    'other': 'その他'
  };
  return map[code] || code;
}

/**
 * 役職のラベル変換
 */
function formatJobTitle(code) {
  const map = {
    'executive': '経営者・役員',
    'department_head': '部長・本部長',
    'manager': '課長・マネージャー',
    'ir_officer': 'IR担当',
    'legal_compliance': '法務・コンプライアンス',
    'finance': '経理・財務',
    'analyst': 'アナリスト・ファンドマネージャー',
    'consultant': 'コンサルタント・アドバイザー',
    'individual': '個人',
    'other': 'その他'
  };
  return map[code] || code;
}

/**
 * Firestore Timestamp を日本時間文字列に変換
 */
function formatTimestamp(ts) {
  if (!ts) return '';
  try {
    const date = ts.toDate ? ts.toDate() : new Date(ts._seconds ? ts._seconds * 1000 : ts);
    return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  } catch (e) {
    return String(ts);
  }
}

/**
 * メニュー追加（スプレッドシートを開いた時）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('会員管理')
    .addItem('会員データを更新', 'exportUsersToSheet')
    .addToUi();
}
