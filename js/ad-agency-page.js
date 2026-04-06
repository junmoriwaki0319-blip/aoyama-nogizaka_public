(function() {
  'use strict';

  // Company Data - 120+ Japanese Advertising Companies
  // Note: ※推計値（Estimated Market Cap in 100M JPY units）
  const companies = [
    // Tier 1: 大手総合代理店
    { name: '電通グループ', ticker: '4324', market: '東P', marketCap: 12800, operatingMargin: 5.2, roe: -18.5, tier: 'Tier1' },
    { name: '博報堂DYホールディングス', ticker: '2433', market: '東P', marketCap: 7200, operatingMargin: 6.8, roe: 9.2, tier: 'Tier1' },
    { name: 'サイバーエージェント', ticker: '4751', market: '東P', marketCap: 5800, operatingMargin: 8.5, roe: 12.1, tier: 'Tier1' },
    { name: 'GMOインターネットグループ', ticker: '9449', market: '東P', marketCap: 3200, operatingMargin: 10.2, roe: 11.5, tier: 'Tier1' },
    { name: 'GMOアドパートナーズ', ticker: '4784', market: '東P', marketCap: 420, operatingMargin: 4.8, roe: 8.5, tier: 'Tier1' },
    { name: 'トランスコスモス', ticker: '9715', market: '東P', marketCap: 2100, operatingMargin: 4.5, roe: 7.8, tier: 'Tier1' },
    // Tier 2: デジタル広告・ネット広告中堅
    { name: 'セプテーニHD', ticker: '4293', market: '東S', marketCap: 850, operatingMargin: 6.2, roe: 10.5, tier: 'Tier2' },
    { name: 'アドウェイズ', ticker: '2489', market: '東S', marketCap: 280, operatingMargin: 3.5, roe: 6.8, tier: 'Tier2' },
    { name: 'バリューコマース', ticker: '2491', market: '東P', marketCap: 580, operatingMargin: 18.5, roe: 15.2, tier: 'Tier2' },
    { name: 'ファンコミュニケーションズ', ticker: '2461', market: '東P', marketCap: 280, operatingMargin: 12.5, roe: 9.8, tier: 'Tier2' },
    { name: 'デジタルガレージ', ticker: '4819', market: '東P', marketCap: 1200, operatingMargin: 8.2, roe: 6.5, tier: 'Tier2' },
    { name: 'ベクトル', ticker: '6058', market: '東P', marketCap: 680, operatingMargin: 7.5, roe: 11.2, tier: 'Tier2' },
    { name: 'アイモバイル', ticker: '6535', market: '東P', marketCap: 320, operatingMargin: 15.2, roe: 12.8, tier: 'Tier2' },
    { name: 'イーガーディアン', ticker: '6050', market: '東P', marketCap: 280, operatingMargin: 10.5, roe: 14.2, tier: 'Tier2' },
    { name: 'アイスタイル', ticker: '3660', market: '東P', marketCap: 520, operatingMargin: 5.8, roe: 8.5, tier: 'Tier2' },
    { name: 'Gunosy', ticker: '6047', market: '東P', marketCap: 120, operatingMargin: 8.2, roe: 6.5, tier: 'Tier2' },
    { name: 'GENOVA', ticker: '9341', market: '東P', marketCap: 180, operatingMargin: 12.5, roe: 15.8, tier: 'Tier2' },
    { name: 'シンクロ・フード', ticker: '3963', market: '東P', marketCap: 120, operatingMargin: 22.5, roe: 18.5, tier: 'Tier2' },
    { name: 'Appier Group', ticker: '4180', market: '東P', marketCap: 1800, operatingMargin: 5.2, roe: 8.5, tier: 'Tier2' },
    { name: 'フリービット', ticker: '3843', market: '東P', marketCap: 280, operatingMargin: 6.8, roe: 7.2, tier: 'Tier2' },
    { name: 'アドバンスクリエイト', ticker: '8798', market: '東P', marketCap: 220, operatingMargin: 8.5, roe: 10.2, tier: 'Tier2' },
    { name: 'Macbee Planet', ticker: '7095', market: '東P', marketCap: 580, operatingMargin: 6.8, roe: 15.5, tier: 'Tier2' },
    // Tier 3: グロース・スタンダード市場
    { name: 'フリークアウトHD', ticker: '6094', market: '東G', marketCap: 180, operatingMargin: 3.5, roe: 5.2, tier: 'Tier3' },
    { name: 'ジーニー', ticker: '6562', market: '東G', marketCap: 350, operatingMargin: 8.5, roe: 12.5, tier: 'Tier3' },
    { name: 'メンバーズ', ticker: '2130', market: '東S', marketCap: 180, operatingMargin: 5.5, roe: 8.2, tier: 'Tier3' },
    { name: 'インタースペース', ticker: '2122', market: '東S', marketCap: 80, operatingMargin: 4.2, roe: 7.5, tier: 'Tier3' },
    { name: 'メディックス', ticker: '331A', market: '東S', marketCap: 120, operatingMargin: 8.5, roe: 12.2, tier: 'Tier3' },
    { name: 'セーラー広告', ticker: '2156', market: '東S', marketCap: 30, operatingMargin: 2.8, roe: 4.5, tier: 'Tier3' },
    { name: 'インサイト', ticker: '2172', market: '東S', marketCap: 15, operatingMargin: 3.2, roe: 5.8, tier: 'Tier3' },
    { name: 'ゲンダイエージェンシー', ticker: '2411', market: '東S', marketCap: 25, operatingMargin: 4.5, roe: 6.2, tier: 'Tier3' },
    { name: 'プラップジャパン', ticker: '2449', market: '東S', marketCap: 45, operatingMargin: 6.8, roe: 8.5, tier: 'Tier3' },
    { name: 'ジェイフロンティア', ticker: '2934', market: '東G', marketCap: 35, operatingMargin: 2.5, roe: 3.8, tier: 'Tier3' },
    { name: 'ヒット', ticker: '378A', market: '東G', marketCap: 25, operatingMargin: 5.2, roe: 8.5, tier: 'Tier3' },
    { name: 'Gモンスター', ticker: '157A', market: '東G', marketCap: 15, operatingMargin: 3.8, roe: 5.2, tier: 'Tier3' },
    { name: 'マスカットグループ', ticker: '195A', market: '東G', marketCap: 20, operatingMargin: 4.5, roe: 6.8, tier: 'Tier3' },
    { name: 'LIFULL', ticker: '2120', market: '東S', marketCap: 180, operatingMargin: 3.2, roe: 2.8, tier: 'Tier3' },
    { name: 'ITメディア', ticker: '2148', market: '東S', marketCap: 250, operatingMargin: 12.5, roe: 10.8, tier: 'Tier3' },
    { name: 'イオレ', ticker: '2334', market: '東G', marketCap: 20, operatingMargin: 2.8, roe: 3.5, tier: 'Tier3' },
    { name: 'サイネックス', ticker: '2376', market: '東S', marketCap: 65, operatingMargin: 5.5, roe: 6.8, tier: 'Tier3' },
    { name: 'オールアバウト', ticker: '2454', market: '東S', marketCap: 45, operatingMargin: 2.2, roe: 3.5, tier: 'Tier3' },
    { name: 'アウンコンサルティング', ticker: '2459', market: '東S', marketCap: 15, operatingMargin: 3.5, roe: 5.2, tier: 'Tier3' },
    { name: 'UNITED', ticker: '2497', market: '東G', marketCap: 120, operatingMargin: 5.8, roe: 7.5, tier: 'Tier3' },
    { name: 'ベクターHD', ticker: '2656', market: '東S', marketCap: 30, operatingMargin: 2.5, roe: 3.2, tier: 'Tier3' },
    { name: 'エフティグループ', ticker: '2763', market: '東S', marketCap: 120, operatingMargin: 6.5, roe: 8.2, tier: 'Tier3' },
    { name: 'クラシル(dely)', ticker: '299A', market: '東G', marketCap: 250, operatingMargin: 5.2, roe: 8.5, tier: 'Tier3' },
    { name: 'フォルシア', ticker: '304A', market: '東G', marketCap: 35, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: 'ネットイヤーグループ', ticker: '3622', market: '東G', marketCap: 45, operatingMargin: 5.8, roe: 7.5, tier: 'Tier3' },
    { name: 'アクセルマーク', ticker: '3624', market: '東G', marketCap: 10, operatingMargin: -2.5, roe: -5.2, tier: 'Tier3' },
    { name: '駅探', ticker: '3646', market: '東G', marketCap: 25, operatingMargin: 8.2, roe: 6.5, tier: 'Tier3' },
    { name: 'イルグルム', ticker: '3690', market: '東S', marketCap: 35, operatingMargin: 5.5, roe: 6.8, tier: 'Tier3' },
    { name: 'カヤック', ticker: '3904', market: '東G', marketCap: 80, operatingMargin: 5.2, roe: 7.5, tier: 'Tier3' },
    { name: 'ショーケース', ticker: '3909', market: '東S', marketCap: 25, operatingMargin: 3.8, roe: 4.5, tier: 'Tier3' },
    { name: 'はてな', ticker: '3930', market: '東G', marketCap: 55, operatingMargin: 8.5, roe: 9.2, tier: 'Tier3' },
    { name: 'カラダノート', ticker: '4014', market: '東G', marketCap: 20, operatingMargin: 2.5, roe: 3.8, tier: 'Tier3' },
    { name: 'クリーマ', ticker: '4017', market: '東G', marketCap: 25, operatingMargin: 1.8, roe: 2.5, tier: 'Tier3' },
    { name: 'ジオロケーションテクノロジー', ticker: '4018', market: '福Q', marketCap: 15, operatingMargin: 5.5, roe: 7.2, tier: 'Tier3' },
    { name: 'まぐまぐ', ticker: '4059', market: '東S', marketCap: 20, operatingMargin: 6.8, roe: 5.5, tier: 'Tier3' },
    { name: 'GMOコマース', ticker: '410A', market: '東G', marketCap: 30, operatingMargin: 4.2, roe: 6.5, tier: 'Tier3' },
    { name: 'GMOテック', ticker: '415A', market: '東G', marketCap: 25, operatingMargin: 5.5, roe: 8.2, tier: 'Tier3' },
    { name: 'ウリドキ', ticker: '418A', market: '名N', marketCap: 10, operatingMargin: 2.8, roe: 3.5, tier: 'Tier3' },
    { name: 'ネオマーケティング', ticker: '4196', market: '東S', marketCap: 25, operatingMargin: 5.8, roe: 8.5, tier: 'Tier3' },
    { name: 'THECOO', ticker: '4255', market: '東G', marketCap: 25, operatingMargin: 2.5, roe: 3.8, tier: 'Tier3' },
    { name: 'ニフティライフスタイル', ticker: '4262', market: '東G', marketCap: 35, operatingMargin: 12.5, roe: 10.8, tier: 'Tier3' },
    { name: 'Jストリーム', ticker: '4308', market: '東G', marketCap: 55, operatingMargin: 8.5, roe: 7.2, tier: 'Tier3' },
    { name: 'NEXYZ.', ticker: '4346', market: '東S', marketCap: 80, operatingMargin: 5.2, roe: 6.5, tier: 'Tier3' },
    { name: 'くふうカンパニー', ticker: '4376', market: '東G', marketCap: 120, operatingMargin: 4.8, roe: 6.2, tier: 'Tier3' },
    { name: 'CINC', ticker: '4378', market: '東G', marketCap: 35, operatingMargin: 8.2, roe: 10.5, tier: 'Tier3' },
    { name: 'ZUU', ticker: '4387', market: '東G', marketCap: 35, operatingMargin: 5.5, roe: 7.8, tier: 'Tier3' },
    { name: 'ミンカブ・ジ・インフォノイド', ticker: '4436', market: '東G', marketCap: 120, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: 'リビンテクノロジーズ', ticker: '4445', market: '東G', marketCap: 25, operatingMargin: 12.5, roe: 15.2, tier: 'Tier3' },
    { name: 'Speee', ticker: '4499', market: '東S', marketCap: 180, operatingMargin: 10.5, roe: 12.8, tier: 'Tier3' },
    { name: 'BRANU', ticker: '460A', market: '東G', marketCap: 25, operatingMargin: 5.2, roe: 7.5, tier: 'Tier3' },
    { name: 'ミラティブ', ticker: '472A', market: '東G', marketCap: 80, operatingMargin: -5.2, roe: -8.5, tier: 'Tier3' },
    { name: 'オリコン', ticker: '4800', market: '東S', marketCap: 120, operatingMargin: 15.5, roe: 12.8, tier: 'Tier3' },
    { name: 'マーキュリーリアルテック', ticker: '5025', market: '東G', marketCap: 35, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: 'AnyMind Group', ticker: '5027', market: '東G', marketCap: 350, operatingMargin: 4.5, roe: 8.5, tier: 'Tier3' },
    { name: 'ウネリー', ticker: '5034', market: '東G', marketCap: 15, operatingMargin: -8.5, roe: -12.5, tier: 'Tier3' },
    { name: 'ファインズ', ticker: '5125', market: '東G', marketCap: 15, operatingMargin: 5.2, roe: 7.5, tier: 'Tier3' },
    { name: 'アイズ', ticker: '5242', market: '東G', marketCap: 25, operatingMargin: 8.5, roe: 12.2, tier: 'Tier3' },
    { name: 'エキサイト', ticker: '5571', market: '東S', marketCap: 35, operatingMargin: 4.5, roe: 5.8, tier: 'Tier3' },
    { name: 'インバウンドプラットフォーム', ticker: '5587', market: '東G', marketCap: 15, operatingMargin: 3.2, roe: 4.5, tier: 'Tier3' },
    { name: 'ナイル', ticker: '5618', market: '東G', marketCap: 25, operatingMargin: 2.8, roe: 3.5, tier: 'Tier3' },
    { name: 'イード', ticker: '6038', market: '東G', marketCap: 35, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: 'レントラックス', ticker: '6045', market: '東G', marketCap: 35, operatingMargin: 6.5, roe: 8.8, tier: 'Tier3' },
    { name: 'イトクロ', ticker: '6049', market: '東G', marketCap: 55, operatingMargin: 18.5, roe: 15.2, tier: 'Tier3' },
    { name: 'トレンダーズ', ticker: '6069', market: '東G', marketCap: 45, operatingMargin: 6.8, roe: 9.5, tier: 'Tier3' },
    { name: 'アライドアーキテクツ', ticker: '6081', market: '東G', marketCap: 35, operatingMargin: 5.2, roe: 7.5, tier: 'Tier3' },
    { name: 'AppBank', ticker: '6177', market: '東G', marketCap: 10, operatingMargin: 1.5, roe: 2.2, tier: 'Tier3' },
    { name: 'GMOメディア', ticker: '6180', market: '東G', marketCap: 80, operatingMargin: 10.5, roe: 15.2, tier: 'Tier3' },
    { name: 'SMN', ticker: '6185', market: '東S', marketCap: 55, operatingMargin: 5.5, roe: 6.8, tier: 'Tier3' },
    { name: 'ホープ', ticker: '6195', market: '東G', marketCap: 15, operatingMargin: -5.2, roe: -8.5, tier: 'Tier3' },
    { name: 'DMソリューションズ', ticker: '6549', market: '東S', marketCap: 45, operatingMargin: 4.5, roe: 5.8, tier: 'Tier3' },
    { name: 'ログリー', ticker: '6579', market: '東G', marketCap: 15, operatingMargin: -2.5, roe: -4.2, tier: 'Tier3' },
    { name: 'EMネットジャパン', ticker: '7036', market: '東G', marketCap: 15, operatingMargin: 4.5, roe: 6.8, tier: 'Tier3' },
    { name: 'アクセスグループHD', ticker: '7042', market: '東S', marketCap: 25, operatingMargin: 5.8, roe: 7.5, tier: 'Tier3' },
    { name: 'ピアラ', ticker: '7044', market: '東S', marketCap: 25, operatingMargin: 3.5, roe: 4.8, tier: 'Tier3' },
    { name: 'ポート', ticker: '7047', market: '東G', marketCap: 250, operatingMargin: 8.5, roe: 15.2, tier: 'Tier3' },
    { name: 'バードマン', ticker: '7063', market: '東G', marketCap: 15, operatingMargin: 2.5, roe: 3.5, tier: 'Tier3' },
    { name: 'ブランディングテクノロジー', ticker: '7067', market: '東G', marketCap: 15, operatingMargin: 3.8, roe: 5.2, tier: 'Tier3' },
    { name: 'フィードフォースグループ', ticker: '7068', market: '東G', marketCap: 45, operatingMargin: 5.5, roe: 7.8, tier: 'Tier3' },
    { name: 'サイバー・バズ', ticker: '7069', market: '東G', marketCap: 55, operatingMargin: 5.2, roe: 8.5, tier: 'Tier3' },
    { name: 'インティメート・マージャー', ticker: '7072', market: '東G', marketCap: 25, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: 'INCLUSIVE HD', ticker: '7078', market: '東G', marketCap: 15, operatingMargin: 2.8, roe: 3.5, tier: 'Tier3' },
    { name: 'ジモティー', ticker: '7082', market: '東G', marketCap: 55, operatingMargin: 15.5, roe: 12.8, tier: 'Tier3' },
    { name: 'ハルメク', ticker: '7119', market: '東S', marketCap: 180, operatingMargin: 8.5, roe: 12.5, tier: 'Tier3' },
    { name: 'Retty', ticker: '7356', market: '東G', marketCap: 25, operatingMargin: 2.5, roe: 3.8, tier: 'Tier3' },
    { name: 'ジオコード', ticker: '7357', market: '東S', marketCap: 25, operatingMargin: 8.5, roe: 10.5, tier: 'Tier3' },
    { name: '東京通信グループ', ticker: '7359', market: '東G', marketCap: 35, operatingMargin: 3.5, roe: 5.2, tier: 'Tier3' },
    { name: 'ベビーカレンダー', ticker: '7363', market: '東G', marketCap: 25, operatingMargin: 8.5, roe: 10.2, tier: 'Tier3' },
    { name: '表示灯', ticker: '7368', market: '東S', marketCap: 80, operatingMargin: 5.5, roe: 6.8, tier: 'Tier3' },
    { name: '全研本社', ticker: '7371', market: '東G', marketCap: 55, operatingMargin: 12.5, roe: 10.8, tier: 'Tier3' },
    { name: 'アシロ', ticker: '7378', market: '東G', marketCap: 120, operatingMargin: 15.2, roe: 18.5, tier: 'Tier3' },
    { name: 'KYORITSU', ticker: '7795', market: '東S', marketCap: 15, operatingMargin: 2.5, roe: 3.2, tier: 'Tier3' },
    { name: 'セキ', ticker: '7857', market: '東S', marketCap: 20, operatingMargin: 2.8, roe: 3.5, tier: 'Tier3' },
    { name: '小林洋行', ticker: '8742', market: '東S', marketCap: 30, operatingMargin: 3.2, roe: 4.5, tier: 'Tier3' },
    { name: '売れるネット広告社', ticker: '9235', market: '東G', marketCap: 25, operatingMargin: 15.5, roe: 12.8, tier: 'Tier3' },
    { name: 'バリュークリエーション', ticker: '9238', market: '東G', marketCap: 15, operatingMargin: 5.5, roe: 8.2, tier: 'Tier3' },
    { name: 'FLネットワークス', ticker: '9241', market: '東G', marketCap: 15, operatingMargin: 3.5, roe: 5.2, tier: 'Tier3' },
    { name: 'デジタリフト', ticker: '9244', market: '東G', marketCap: 25, operatingMargin: 5.8, roe: 8.5, tier: 'Tier3' },
    { name: 'ラストワンマイル', ticker: '9252', market: '東G', marketCap: 20, operatingMargin: 4.5, roe: 6.2, tier: 'Tier3' },
    { name: 'CS-C', ticker: '9258', market: '東G', marketCap: 15, operatingMargin: 3.5, roe: 5.8, tier: 'Tier3' },
    { name: 'トリドリ', ticker: '9337', market: '東G', marketCap: 35, operatingMargin: 5.5, roe: 8.2, tier: 'Tier3' },
    { name: 'AViC', ticker: '9554', market: '東G', marketCap: 120, operatingMargin: 8.5, roe: 15.5, tier: 'Tier3' },
    { name: 'マイクロアド', ticker: '9553', market: '東G', marketCap: 80, operatingMargin: 6.5, roe: 8.8, tier: 'Tier3' },
    { name: 'グラッドキューブ', ticker: '9561', market: '東G', marketCap: 25, operatingMargin: 8.5, roe: 12.2, tier: 'Tier3' }
  ];

  function populateTable() {
    const tbody = document.getElementById('companiesTableBody');
    tbody.innerHTML = '';
    companies.forEach((company, index) => {
      const row = document.createElement('tr');
      row.innerHTML = `
                    <td>${index + 1}</td>
                    <td><strong>${company.name}</strong></td>
                    <td>${company.ticker}</td>
                    <td>${company.market}</td>
                    <td>${company.marketCap.toLocaleString()}</td>
                    <td>${company.operatingMargin.toFixed(1)}%</td>
                    <td>${company.roe.toFixed(1)}%</td>
                    <td><span style="padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: 600; background-color: ${company.tier === 'Tier1' ? 'var(--navy)' : company.tier === 'Tier2' ? 'var(--info)' : 'var(--text-tertiary)'}; color: white;">${company.tier}</span></td>
                `;
      tbody.appendChild(row);
    });
  }

  function createMarketCapChart() {
    const ctx = document.getElementById('marketCapChart');
    if (!ctx) return;

    const topCompanies = companies
      .filter(c => c.marketCap > 1500)
      .sort((a, b) => b.marketCap - a.marketCap)
      .slice(0, 30);

    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: topCompanies.map(c => c.name),
        datasets: [{
          label: 'Market Cap (Billions JPY)',
          data: topCompanies.map(c => c.marketCap),
          backgroundColor: 'var(--navy)',
          borderColor: 'var(--gold)',
          borderWidth: 2,
          borderRadius: 4,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: {
              color: 'var(--text-primary)',
              font: { size: 12 }
            }
          },
          tooltip: {
            backgroundColor: 'var(--bg-white)',
            titleColor: 'var(--navy)',
            bodyColor: 'var(--text-primary)',
            borderColor: 'var(--navy)',
            borderWidth: 1,
            padding: 12,
            displayColors: false,
            callbacks: {
              label: function(ctx) {
                return '時価総額: ¥' + ctx.parsed.x.toLocaleString() + '億';
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { color: 'var(--text-tertiary)' },
            grid: { color: 'var(--border-light)' }
          },
          y: {
            ticks: { color: 'var(--text-primary)', font: { size: 11 } },
            grid: { display: false }
          }
        }
      }
    });
  }

  function setupNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
      item.addEventListener('click', function() {
        navItems.forEach(i => i.classList.remove('active'));
        this.classList.add('active');
      });
    });
  }

  // Initialize - called by auth module after login, or immediately for table structure
  window.loadPremiumData = function() {
    populateTable();
    createMarketCapChart();
    setupNavigation();
  };

  // Init nav on load (for non-authenticated state too)
  document.addEventListener('DOMContentLoaded', function() {
    setupNavigation();
  });
})();
