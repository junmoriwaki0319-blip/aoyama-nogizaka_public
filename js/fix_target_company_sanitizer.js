/**
 * target_company サニタイザー パッチ
 *
 * 用途: activist-dashboard-page.js の loadData 関数内で、
 *       reports データ取得直後に呼び出す。
 *
 * 背景: XBRLパーサーがEDINET文書内のHTMLテーブルCSSを
 *        target_company として誤取得するバグ（28件確認）
 *
 * 適用方法:
 *   1. このファイルを /js/ に配置
 *   2. activist-dashboard.html の activist-dashboard-page.js の前に読み込む
 *      <script src="/js/fix_target_company_sanitizer.js"></script>
 *   3. loadData 関数内の reports = data.reports; の直後に
 *      reports = sanitizeReports(reports); を追加
 *
 * または、loadData を直接修正する場合:
 *   reports = data.reports; の行の直後に以下を挿入:
 *   reports = sanitizeReports(reports);
 */

// === 既知のバグレコード修正マップ（証券コード → 正しい企業名） ===
const TARGET_COMPANY_CORRECTIONS = {
  '3608': '株式会社TSIホールディングス',
  '6433': 'ヘファイスト株式会社',
  '7380': '株式会社十六フィナンシャルグループ',
  '8289': '株式会社Olympicグループ',
  '8233': '株式会社高島屋',
  '6086': 'シンメンテホールディングス株式会社',
  '3083': '株式会社シーズメン',
  '3224': '株式会社ゼネラル・オイスター',
  '9005': '東急株式会社'
};

// === CSS/HTML断片を検出する正規表現 ===
const CSS_FRAGMENT_PATTERN = /^[\d.]*(pt|px|em|rem|mm|%)|border|collapse|none\s*;|^\d+(\.\d+)?(pt|px)/;

/**
 * target_company フィールドがCSS/HTML断片で汚染されていないか検証する
 * @param {string} targetCompany - target_company の値
 * @returns {boolean} true = 汚染されている
 */
function isContaminated(targetCompany) {
  if (!targetCompany || typeof targetCompany !== 'string') return false;

  const tc = targetCompany.trim();

  // CSS断片パターンの検出
  if (CSS_FRAGMENT_PATTERN.test(tc)) return true;

  // HTMLエンティティの残骸
  if (tc.includes('&gt;') || tc.includes('&lt;') || tc.includes('&amp;')) return true;

  // HTMLタグの残骸
  if (tc.includes("'>") || tc.includes('">')) return true;

  // CSS プロパティ値のパターン
  if (/border-collapse|font-size|margin|padding/.test(tc)) return true;

  // 日本語・英字・数字のいずれも含まない異常値
  if (!/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\u3400-\u4DBFa-zA-Z]/.test(tc)) return true;

  return false;
}

/**
 * reports 配列の target_company を検証・修正する
 * @param {Array} reports - 報告データの配列
 * @returns {Array} サニタイズ済みの配列
 */
function sanitizeReports(reports) {
  if (!Array.isArray(reports)) return reports;

  let fixedCount = 0;

  for (let i = 0; i < reports.length; i++) {
    const report = reports[i];

    if (isContaminated(report.target_company)) {
      const correction = TARGET_COMPANY_CORRECTIONS[report.sec_code];

      if (correction) {
        // 既知の修正マップから企業名を取得
        report.target_company = correction;
        fixedCount++;
      } else {
        // 未知の証券コード → 「(証券コード: XXXX)」で代替表示
        report.target_company = '(データ修正中: ' + report.sec_code + ')';
        fixedCount++;
        console.warn(
          '[sanitizer] Unknown sec_code with contaminated target_company:',
          report.sec_code, '→', report.target_company,
          'doc_id:', report.doc_id
        );
      }
    }
  }

  if (fixedCount > 0) {
    console.info('[sanitizer] Fixed ' + fixedCount + ' contaminated target_company records');
  }

  return reports;
}

// グローバルに公開
window.sanitizeReports = sanitizeReports;
window.isContaminated = isContaminated;
