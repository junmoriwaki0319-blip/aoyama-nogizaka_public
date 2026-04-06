#!/usr/bin/env node
/**
 * yahoo-finance2 を使って広告セクター・エンタメセクターの企業データを取得
 * 出力: data/ad-agency-companies.json, data/entertainment-companies.json
 *
 * 使い方: node scripts/fetch-sector-financials.js
 */

const fs = require('fs');
const path = require('path');
const YahooFinance = require('yahoo-finance2').default;
const yahooFinance = new YahooFinance({ suppressNotices: ['yahooSurvey'] });

const BASE = path.resolve(__dirname, '..');
const DATA_DIR = path.join(BASE, 'data');

// 市場コード → Yahoo Finance サフィックス
const MARKET_SUFFIX = {
  '東P': '.T', '東S': '.T', '東G': '.T',
  '名N': '.N', '福Q': '.T', // 福証もYahooでは.T
};

const BATCH_SIZE = 5;       // 並列数
const BATCH_DELAY_MS = 300; // バッチ間のディレイ
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─── 広告セクター: 証券コード一覧（旧 ad-agency-page.js の companies 配列由来） ───
function extractAdAgencyCompanies() {
  return [
    { name: '電通グループ', ticker: '4324', market: '東P', tier: 'Tier1' },
    { name: '博報堂DYホールディングス', ticker: '2433', market: '東P', tier: 'Tier1' },
    { name: 'サイバーエージェント', ticker: '4751', market: '東P', tier: 'Tier1' },
    { name: 'GMOインターネットグループ', ticker: '9449', market: '東P', tier: 'Tier1' },
    { name: 'GMOアドパートナーズ', ticker: '4784', market: '東P', tier: 'Tier1' },
    { name: 'トランスコスモス', ticker: '9715', market: '東P', tier: 'Tier1' },
    { name: 'セプテーニHD', ticker: '4293', market: '東S', tier: 'Tier2' },
    { name: 'アドウェイズ', ticker: '2489', market: '東S', tier: 'Tier2' },
    { name: 'バリューコマース', ticker: '2491', market: '東P', tier: 'Tier2' },
    { name: 'ファンコミュニケーションズ', ticker: '2461', market: '東P', tier: 'Tier2' },
    { name: 'デジタルガレージ', ticker: '4819', market: '東P', tier: 'Tier2' },
    { name: 'ベクトル', ticker: '6058', market: '東P', tier: 'Tier2' },
    { name: 'アイモバイル', ticker: '6535', market: '東P', tier: 'Tier2' },
    { name: 'イーガーディアン', ticker: '6050', market: '東P', tier: 'Tier2' },
    { name: 'アイスタイル', ticker: '3660', market: '東P', tier: 'Tier2' },
    { name: 'Gunosy', ticker: '6047', market: '東P', tier: 'Tier2' },
    { name: 'GENOVA', ticker: '9341', market: '東P', tier: 'Tier2' },
    { name: 'シンクロ・フード', ticker: '3963', market: '東P', tier: 'Tier2' },
    { name: 'Appier Group', ticker: '4180', market: '東P', tier: 'Tier2' },
    { name: 'フリービット', ticker: '3843', market: '東P', tier: 'Tier2' },
    { name: 'アドバンスクリエイト', ticker: '8798', market: '東P', tier: 'Tier2' },
    { name: 'Macbee Planet', ticker: '7095', market: '東P', tier: 'Tier2' },
    { name: 'フリークアウトHD', ticker: '6094', market: '東G', tier: 'Tier3' },
    { name: 'ジーニー', ticker: '6562', market: '東G', tier: 'Tier3' },
    { name: 'メンバーズ', ticker: '2130', market: '東S', tier: 'Tier3' },
    { name: 'インタースペース', ticker: '2122', market: '東S', tier: 'Tier3' },
    { name: 'メディックス', ticker: '331A', market: '東S', tier: 'Tier3' },
    { name: 'セーラー広告', ticker: '2156', market: '東S', tier: 'Tier3' },
    { name: 'インサイト', ticker: '2172', market: '東S', tier: 'Tier3' },
    { name: 'ゲンダイエージェンシー', ticker: '2411', market: '東S', tier: 'Tier3' },
    { name: 'プラップジャパン', ticker: '2449', market: '東S', tier: 'Tier3' },
    { name: 'ジェイフロンティア', ticker: '2934', market: '東G', tier: 'Tier3' },
    { name: 'ヒット', ticker: '378A', market: '東G', tier: 'Tier3' },
    { name: 'Gモンスター', ticker: '157A', market: '東G', tier: 'Tier3' },
    { name: 'マスカットグループ', ticker: '195A', market: '東G', tier: 'Tier3' },
    { name: 'LIFULL', ticker: '2120', market: '東S', tier: 'Tier3' },
    { name: 'ITメディア', ticker: '2148', market: '東S', tier: 'Tier3' },
    { name: 'イオレ', ticker: '2334', market: '東G', tier: 'Tier3' },
    { name: 'サイネックス', ticker: '2376', market: '東S', tier: 'Tier3' },
    { name: 'オールアバウト', ticker: '2454', market: '東S', tier: 'Tier3' },
    { name: 'アウンコンサルティング', ticker: '2459', market: '東S', tier: 'Tier3' },
    { name: 'UNITED', ticker: '2497', market: '東G', tier: 'Tier3' },
    { name: 'ベクターHD', ticker: '2656', market: '東S', tier: 'Tier3' },
    { name: 'エフティグループ', ticker: '2763', market: '東S', tier: 'Tier3' },
    { name: 'クラシル(dely)', ticker: '299A', market: '東G', tier: 'Tier3' },
    { name: 'フォルシア', ticker: '304A', market: '東G', tier: 'Tier3' },
    { name: 'ネットイヤーグループ', ticker: '3622', market: '東G', tier: 'Tier3' },
    { name: 'アクセルマーク', ticker: '3624', market: '東G', tier: 'Tier3' },
    { name: '駅探', ticker: '3646', market: '東G', tier: 'Tier3' },
    { name: 'イルグルム', ticker: '3690', market: '東S', tier: 'Tier3' },
    { name: 'カヤック', ticker: '3904', market: '東G', tier: 'Tier3' },
    { name: 'ショーケース', ticker: '3909', market: '東S', tier: 'Tier3' },
    { name: 'はてな', ticker: '3930', market: '東G', tier: 'Tier3' },
    { name: 'カラダノート', ticker: '4014', market: '東G', tier: 'Tier3' },
    { name: 'クリーマ', ticker: '4017', market: '東G', tier: 'Tier3' },
    { name: 'ジオロケーションテクノロジー', ticker: '4018', market: '福Q', tier: 'Tier3' },
    { name: 'まぐまぐ', ticker: '4059', market: '東S', tier: 'Tier3' },
    { name: 'GMOコマース', ticker: '410A', market: '東G', tier: 'Tier3' },
    { name: 'GMOテック', ticker: '415A', market: '東G', tier: 'Tier3' },
    { name: 'ウリドキ', ticker: '418A', market: '名N', tier: 'Tier3' },
    { name: 'ネオマーケティング', ticker: '4196', market: '東S', tier: 'Tier3' },
    { name: 'THECOO', ticker: '4255', market: '東G', tier: 'Tier3' },
    { name: 'ニフティライフスタイル', ticker: '4262', market: '東G', tier: 'Tier3' },
    { name: 'Jストリーム', ticker: '4308', market: '東G', tier: 'Tier3' },
    { name: 'NEXYZ.', ticker: '4346', market: '東S', tier: 'Tier3' },
    { name: 'くふうカンパニー', ticker: '4376', market: '東G', tier: 'Tier3' },
    { name: 'CINC', ticker: '4378', market: '東G', tier: 'Tier3' },
    { name: 'ZUU', ticker: '4387', market: '東G', tier: 'Tier3' },
    { name: 'ミンカブ・ジ・インフォノイド', ticker: '4436', market: '東G', tier: 'Tier3' },
    { name: 'リビンテクノロジーズ', ticker: '4445', market: '東G', tier: 'Tier3' },
    { name: 'Speee', ticker: '4499', market: '東S', tier: 'Tier3' },
    { name: 'BRANU', ticker: '460A', market: '東G', tier: 'Tier3' },
    { name: 'ミラティブ', ticker: '472A', market: '東G', tier: 'Tier3' },
    { name: 'オリコン', ticker: '4800', market: '東S', tier: 'Tier3' },
    { name: 'マーキュリーリアルテック', ticker: '5025', market: '東G', tier: 'Tier3' },
    { name: 'AnyMind Group', ticker: '5027', market: '東G', tier: 'Tier3' },
    { name: 'ウネリー', ticker: '5034', market: '東G', tier: 'Tier3' },
    { name: 'ファインズ', ticker: '5125', market: '東G', tier: 'Tier3' },
    { name: 'アイズ', ticker: '5242', market: '東G', tier: 'Tier3' },
    { name: 'エキサイト', ticker: '5571', market: '東S', tier: 'Tier3' },
    { name: 'インバウンドプラットフォーム', ticker: '5587', market: '東G', tier: 'Tier3' },
    { name: 'ナイル', ticker: '5618', market: '東G', tier: 'Tier3' },
    { name: 'イード', ticker: '6038', market: '東G', tier: 'Tier3' },
    { name: 'レントラックス', ticker: '6045', market: '東G', tier: 'Tier3' },
    { name: 'イトクロ', ticker: '6049', market: '東G', tier: 'Tier3' },
    { name: 'トレンダーズ', ticker: '6069', market: '東G', tier: 'Tier3' },
    { name: 'アライドアーキテクツ', ticker: '6081', market: '東G', tier: 'Tier3' },
    { name: 'AppBank', ticker: '6177', market: '東G', tier: 'Tier3' },
    { name: 'GMOメディア', ticker: '6180', market: '東G', tier: 'Tier3' },
    { name: 'SMN', ticker: '6185', market: '東S', tier: 'Tier3' },
    { name: 'ホープ', ticker: '6195', market: '東G', tier: 'Tier3' },
    { name: 'DMソリューションズ', ticker: '6549', market: '東S', tier: 'Tier3' },
    { name: 'ログリー', ticker: '6579', market: '東G', tier: 'Tier3' },
    { name: 'EMネットジャパン', ticker: '7036', market: '東G', tier: 'Tier3' },
    { name: 'アクセスグループHD', ticker: '7042', market: '東S', tier: 'Tier3' },
    { name: 'ピアラ', ticker: '7044', market: '東S', tier: 'Tier3' },
    { name: 'ポート', ticker: '7047', market: '東G', tier: 'Tier3' },
    { name: 'バードマン', ticker: '7063', market: '東G', tier: 'Tier3' },
    { name: 'ブランディングテクノロジー', ticker: '7067', market: '東G', tier: 'Tier3' },
    { name: 'フィードフォースグループ', ticker: '7068', market: '東G', tier: 'Tier3' },
    { name: 'サイバー・バズ', ticker: '7069', market: '東G', tier: 'Tier3' },
    { name: 'インティメート・マージャー', ticker: '7072', market: '東G', tier: 'Tier3' },
    { name: 'INCLUSIVE HD', ticker: '7078', market: '東G', tier: 'Tier3' },
    { name: 'ジモティー', ticker: '7082', market: '東G', tier: 'Tier3' },
    { name: 'ハルメク', ticker: '7119', market: '東S', tier: 'Tier3' },
    { name: 'Retty', ticker: '7356', market: '東G', tier: 'Tier3' },
    { name: 'ジオコード', ticker: '7357', market: '東S', tier: 'Tier3' },
    { name: '東京通信グループ', ticker: '7359', market: '東G', tier: 'Tier3' },
    { name: 'ベビーカレンダー', ticker: '7363', market: '東G', tier: 'Tier3' },
    { name: '表示灯', ticker: '7368', market: '東S', tier: 'Tier3' },
    { name: '全研本社', ticker: '7371', market: '東G', tier: 'Tier3' },
    { name: 'アシロ', ticker: '7378', market: '東G', tier: 'Tier3' },
    { name: 'KYORITSU', ticker: '7795', market: '東S', tier: 'Tier3' },
    { name: 'セキ', ticker: '7857', market: '東S', tier: 'Tier3' },
    { name: '小林洋行', ticker: '8742', market: '東S', tier: 'Tier3' },
    { name: '売れるネット広告社', ticker: '9235', market: '東G', tier: 'Tier3' },
    { name: 'バリュークリエーション', ticker: '9238', market: '東G', tier: 'Tier3' },
    { name: 'FLネットワークス', ticker: '9241', market: '東G', tier: 'Tier3' },
    { name: 'デジタリフト', ticker: '9244', market: '東G', tier: 'Tier3' },
    { name: 'ラストワンマイル', ticker: '9252', market: '東G', tier: 'Tier3' },
    { name: 'CS-C', ticker: '9258', market: '東G', tier: 'Tier3' },
    { name: 'トリドリ', ticker: '9337', market: '東G', tier: 'Tier3' },
    { name: 'AViC', ticker: '9554', market: '東G', tier: 'Tier3' },
    { name: 'マイクロアド', ticker: '9553', market: '東G', tier: 'Tier3' },
    { name: 'グラッドキューブ', ticker: '9561', market: '東G', tier: 'Tier3' },
  ];
}

// ─── エンタメセクター: 証券コード一覧（旧 index.html のテーブル由来） ───
function extractEntertainmentCompanies() {
  return [
    // ゲーム開発・パブリッシャー
    { name: '任天堂', ticker: '7974', category: 'game-publisher' },
    { name: 'ソニーグループ', ticker: '6758', category: 'game-publisher' },
    { name: 'バンダイナムコHD', ticker: '7832', category: 'game-publisher' },
    { name: 'コナミグループ', ticker: '9766', category: 'game-publisher' },
    { name: 'カプコン', ticker: '9697', category: 'game-publisher' },
    { name: 'スクウェア・エニックスHD', ticker: '9684', category: 'game-publisher' },
    { name: 'セガサミーHD', ticker: '6460', category: 'game-publisher' },
    { name: 'コーエーテクモHD', ticker: '9658', category: 'game-publisher' },
    { name: 'ネクソン', ticker: '3659', category: 'game-publisher' },
    { name: '東宝', ticker: '9602', category: 'game-publisher' },
    { name: '松竹', ticker: '9601', category: 'game-publisher' },
    { name: '東映', ticker: '9605', category: 'game-publisher' },
    // モバイルゲーム
    { name: 'サイバーエージェント', ticker: '4751', category: 'mobile-game' },
    { name: 'MIXI', ticker: '2121', category: 'mobile-game' },
    { name: 'DeNA', ticker: '2432', category: 'mobile-game' },
    { name: 'ガンホー・オンライン', ticker: '3765', category: 'mobile-game' },
    { name: 'グリー', ticker: '3632', category: 'mobile-game' },
    { name: 'アカツキ', ticker: '3932', category: 'mobile-game' },
    { name: 'KLab', ticker: '3656', category: 'mobile-game' },
    { name: 'ドリコム', ticker: '3623', category: 'mobile-game' },
    { name: 'Aiming', ticker: '3911', category: 'mobile-game' },
    { name: 'マイネット', ticker: '3928', category: 'mobile-game' },
    { name: 'コロプラ', ticker: '3668', category: 'mobile-game' },
    { name: 'gumi', ticker: '3903', category: 'mobile-game' },
    { name: 'エイチーム', ticker: '3662', category: 'mobile-game' },
    { name: 'エクストリーム', ticker: '6033', category: 'mobile-game' },
    { name: 'ドリームインキュベータ', ticker: '3793', category: 'mobile-game' },
    { name: 'enish', ticker: '3667', category: 'mobile-game' },
    { name: 'はてな', ticker: '3930', category: 'mobile-game' },
    { name: 'モブキャストHD', ticker: '3664', category: 'mobile-game' },
    { name: 'ダブルスタンダード', ticker: '3925', category: 'mobile-game' },
    // アニメ・IP
    { name: '東映アニメーション', ticker: '4816', category: 'anime-ip' },
    { name: 'IGポート', ticker: '3791', category: 'anime-ip' },
    { name: 'エイベックス', ticker: '7860', category: 'anime-ip' },
    { name: 'テイクアンドギヴ・ニーズ', ticker: '4331', category: 'anime-ip' },
    { name: 'ディー・エル・イー', ticker: '3686', category: 'anime-ip' },
    { name: 'マーベラス', ticker: '7844', category: 'anime-ip' },
    { name: 'KADOKAWA', ticker: '9468', category: 'anime-ip' },
    { name: 'タカラトミー', ticker: '7867', category: 'anime-ip' },
    { name: 'オリエンタルランド', ticker: '4661', category: 'anime-ip' },
    { name: 'フィールズ', ticker: '2767', category: 'anime-ip' },
    { name: 'ハピネット', ticker: '7552', category: 'anime-ip' },
    { name: 'ブシロード', ticker: '7803', category: 'anime-ip' },
    { name: 'クリーク・アンド・リバー社', ticker: '4763', category: 'anime-ip' },
    // VTuber・メタバース
    { name: 'ANYCOLOR', ticker: '5765', category: 'vtuber-meta' },
    { name: 'カバー', ticker: '5253', category: 'vtuber-meta' },
    { name: 'UUUM', ticker: '3990', category: 'vtuber-meta' },
    // eスポーツ・周辺
    { name: 'GLOE', ticker: '9565', category: 'esports-peripheral' },
    { name: 'INCLUSIVE', ticker: '7078', category: 'esports-peripheral' },
    { name: 'IMAGICA GROUP', ticker: '6879', category: 'esports-peripheral' },
    { name: 'GMOペイメントゲートウェイ', ticker: '3769', category: 'esports-peripheral' },
    // 開発ツール・受託
    { name: 'シリコンスタジオ', ticker: '3907', category: 'dev-tools' },
    { name: 'アステリア', ticker: '3853', category: 'dev-tools' },
    { name: 'フォーサイド', ticker: '2330', category: 'dev-tools' },
    { name: 'イグニス', ticker: '3689', category: 'dev-tools' },
    { name: 'インターワークス', ticker: '6032', category: 'dev-tools' },
    // 追加企業群
    { name: 'オルトプラス', ticker: '3672', category: 'mobile-game' },
    { name: 'GameWith', ticker: '6552', category: 'mobile-game' },
    { name: 'SHIFT', ticker: '3697', category: 'mobile-game' },
    { name: 'トランスコスモス', ticker: '9715', category: 'game-publisher' },
    { name: 'ブロッコリー', ticker: '2706', category: 'anime-ip' },
    { name: 'メディアドゥ', ticker: '3678', category: 'anime-ip' },
    { name: 'ケイブ', ticker: '3760', category: 'anime-ip' },
    { name: '日本一ソフトウェア', ticker: '3851', category: 'game-publisher' },
    { name: 'インタートレード', ticker: '3747', category: 'game-publisher' },
    { name: 'カヤック', ticker: '3904', category: 'game-publisher' },
    { name: 'PKSHA Technology', ticker: '3993', category: 'mobile-game' },
    { name: 'トーセ', ticker: '4728', category: 'game-publisher' },
    { name: 'サン電子', ticker: '6736', category: 'game-publisher' },
    { name: '日本BS放送', ticker: '9414', category: 'anime-ip' },
    { name: 'ピクセルカンパニーズ', ticker: '2743', category: 'game-publisher' },
    { name: 'SEホールディングス・アンド・インキュベーションズ', ticker: '9478', category: 'esports-peripheral' },
    { name: 'インテア・ホールディングス', ticker: '3734', category: 'game-publisher' },
    { name: 'AppBank', ticker: '6177', category: 'mobile-game' },
    { name: 'アクセルマーク', ticker: '3624', category: 'mobile-game' },
    { name: 'イーブックイニシアティブジャパン', ticker: '3658', category: 'anime-ip' },
    { name: 'バンク・オブ・イノベーション', ticker: '4393', category: 'game-publisher' },
    { name: 'ワールドHD', ticker: '3612', category: 'game-publisher' },
    { name: 'ガーラ', ticker: '4777', category: 'mobile-game' },
    { name: 'ボルテージ', ticker: '3639', category: 'anime-ip' },
    { name: 'ピー・シー・エー', ticker: '9629', category: 'game-publisher' },
    { name: 'イー・ガーディアン', ticker: '6050', category: 'game-publisher' },
    { name: 'セプテーニHD', ticker: '4293', category: 'mobile-game' },
    { name: 'デジタルハーツHD', ticker: '3676', category: 'anime-ip' },
    { name: 'アイフィスジャパン', ticker: '7833', category: 'game-publisher' },
    { name: '弁護士ドットコム', ticker: '6027', category: 'game-publisher' },
    { name: '日本エンタープライズ', ticker: '4829', category: 'anime-ip' },
    { name: 'セレスポ', ticker: '9625', category: 'game-publisher' },
    { name: 'セレス', ticker: '3696', category: 'game-publisher' },
    { name: 'アイリッジ', ticker: '3917', category: 'mobile-game' },
    { name: 'Ubicom HD', ticker: '3937', category: 'game-publisher' },
    { name: 'イルグルム', ticker: '3690', category: 'anime-ip' },
    { name: 'サーバーワークス', ticker: '4434', category: 'game-publisher' },
    { name: 'リアルワールド', ticker: '3691', category: 'mobile-game' },
    { name: 'ディー・ディー・エス', ticker: '3782', category: 'game-publisher' },
    { name: 'sMedio', ticker: '3913', category: 'anime-ip' },
    { name: 'ポート', ticker: '7047', category: 'game-publisher' },
    { name: 'アイ・アールジャパンHD', ticker: '6035', category: 'game-publisher' },
  ];
}

// ─── Yahoo Finance からデータ取得（quote のみ + quoteSummary で営業利益率/ROE） ───
async function fetchCompanyData(ticker, market) {
  const suffix = MARKET_SUFFIX[market] || '.T';
  const symbol = ticker + suffix;
  try {
    // quote + quoteSummary を並列で同時発行
    const [quote, summary] = await Promise.all([
      yahooFinance.quote(symbol),
      yahooFinance.quoteSummary(symbol, {
        modules: ['financialData', 'defaultKeyStatistics'],
      }).catch(() => null),
    ]);

    let opMargin = null, roe = null, per = null, pbr = null;
    if (summary) {
      const fd = summary.financialData || {};
      const ks = summary.defaultKeyStatistics || {};
      opMargin = fd.operatingMargins != null ? +(fd.operatingMargins * 100).toFixed(1) : null;
      roe = fd.returnOnEquity != null ? +(fd.returnOnEquity * 100).toFixed(1) : null;
      per = ks.trailingPE ?? ks.forwardPE ?? null;
      pbr = ks.priceToBook ?? null;
      if (per != null) per = +per.toFixed(1);
      if (pbr != null) pbr = +pbr.toFixed(2);
    }
    // quote からもPER/PBRが取れる（summaryが失敗した場合のフォールバック）
    if (per == null && quote.trailingPE) per = +quote.trailingPE.toFixed(1);
    if (pbr == null && quote.priceToBook) pbr = +quote.priceToBook.toFixed(2);

    const price = quote.regularMarketPrice ?? null;
    const mcapRaw = quote.marketCap ?? 0;
    const marketCap = mcapRaw ? Math.round(mcapRaw / 1e8) : null;

    return { price, marketCap, operatingMargin: opMargin, roe, per, pbr };
  } catch (e) {
    return null; // エラー時はスキップ（ログはバッチ側で出す）
  }
}

// ─── バッチ処理: BATCH_SIZE社を並列取得 ───
async function fetchBatch(items, marketOverride) {
  const results = [];
  for (let i = 0; i < items.length; i += BATCH_SIZE) {
    const batch = items.slice(i, i + BATCH_SIZE);
    const batchResults = await Promise.all(
      batch.map(c => fetchCompanyData(c.ticker, marketOverride || c.market).then(data => ({ ...c, data })))
    );
    for (const r of batchResults) {
      if (r.data) {
        results.push(r);
        process.stdout.write('.');
      } else {
        process.stdout.write('x');
      }
    }
    if (i + BATCH_SIZE < items.length) await sleep(BATCH_DELAY_MS);
  }
  process.stdout.write('\n');
  return results;
}

// ─── メイン ───
async function main() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

  // === 広告セクター ===
  console.log('=== 広告セクター: 企業リスト抽出 ===');
  const adRaw = extractAdAgencyCompanies();
  console.log(`  ${adRaw.length}社を検出`);

  console.log(`=== 広告セクター: Yahoo Finance データ取得 (${BATCH_SIZE}社並列) ===`);
  const adResults = await fetchBatch(adRaw);
  const adCompanies = adResults.map(r => ({
    name: r.name, ticker: r.ticker, market: r.market, tier: r.tier,
    marketCap: r.data.marketCap, operatingMargin: r.data.operatingMargin,
    roe: r.data.roe, per: r.data.per, pbr: r.data.pbr, price: r.data.price,
  }));

  const adOut = path.join(DATA_DIR, 'ad-agency-companies.json');
  fs.writeFileSync(adOut, JSON.stringify(adCompanies, null, 2), 'utf8');
  console.log(`  → ${adCompanies.length}/${adRaw.length}社を出力\n`);

  // === エンタメセクター ===
  console.log('=== エンタメセクター: 企業リスト抽出 ===');
  const entRaw = extractEntertainmentCompanies();
  const seen = new Set();
  const entFiltered = entRaw.filter(c => {
    if (seen.has(c.ticker)) return false;
    seen.add(c.ticker);
    return true;
  });
  console.log(`  ${entFiltered.length}社（重複除外済）`);

  console.log(`=== エンタメセクター: Yahoo Finance データ取得 (${BATCH_SIZE}社並列) ===`);
  const entResults = await fetchBatch(entFiltered, '東P');
  const entCompanies = entResults.map(r => ({
    name: r.name, ticker: r.ticker, category: r.category,
    marketCap: r.data.marketCap, operatingMargin: r.data.operatingMargin,
    roe: r.data.roe, per: r.data.per, pbr: r.data.pbr, price: r.data.price,
  }));

  const entOut = path.join(DATA_DIR, 'entertainment-companies.json');
  fs.writeFileSync(entOut, JSON.stringify(entCompanies, null, 2), 'utf8');
  console.log(`  → ${entCompanies.length}/${entFiltered.length}社を出力\n`);

  console.log('=== 全セクター取得完了 ===');
  console.log(`  広告:    ${adCompanies.length}/${adRaw.length}社`);
  console.log(`  エンタメ: ${entCompanies.length}/${entFiltered.length}社`);
}

main().catch(e => {
  console.error('ERROR:', e);
  process.exit(1);
});
