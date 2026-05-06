# 商社セクター・ダッシュボード — 公開前チェックリスト

最終更新: 2026-05-07

`/sogo-shosha.html` を一般公開する前にクリアすべき項目を網羅。
本ファイルは公開後に削除可。

---

## ファクトチェック必須 (DRAFT 状態)

[data/sogo-shosha/refs/company-narratives.json](refs/company-narratives.json) は `_fact_check_status: DRAFT` 状態。
公開前に編集サイドで以下のレビューが必要:

### 1. 自動取得 KPI (Yahoo Finance) — 検証不要
✅ DOE / 配当性向 / ROE / ROIC / PBR / PER / EV·EBITDA / ネット負債/EBITDA / 現預金/時価総額
出所: `yahoo-finance2` npm 経由。基準期は `lastFiscalYearEnd: 2026-03-31` (FY25)。

### 2. 手動キュレーション (要確認)

#### 各社ナラティブの固有名詞
- 三菱商事: 資源権益 (Anglo American 系銅鉱山持分の表記が「南米」で十分か、Quellaveco / Los Pelambres 等の固有名詞を入れるか)
- 三井物産: IHH Healthcare 出資比率 (本文 "約 33%" 表記、IR 開示の最新値で確認)
- 伊藤忠: ファミマ店舗数 (本文 "約 1.6 万店" は概数。直近は約 16,300 店)
- 丸紅: 海外電力 IPP の表現 (本文「世界最大級」は IR の主張範囲に収めた一般化表記)
- 双日: ネット負債/EBITDA 10.38 倍の M&A 起因仮説は「推測」と明示済。**実態を IR で確認のうえ本文を確定**
- 豊田通商: CFAO のアフリカ展開国数 (本文「50 ヶ国超」は IR 公表値で確認)

#### バフェット保有比率
[refs/buffett-holdings.json](refs/buffett-holdings.json) の `holdings_pct.<ticker>.trend` は **概算値**。
直近の正確な値は EDINET 大量保有報告書 (5 社 × 直近 14 条 / 変更報告) で確認のこと。
- 5 期分の時系列ロジックは保ったまま、最新 1〜2 期のみ正確値に差し替えるのが現実的

#### 戦略・提携施策 (refs/strategic-initiatives.json)
- 三菱商事: 「ENEOS との Stronghold 北米 LNG」 → **要確認 (固有プロジェクト名の信頼性低)**
- 三井物産: ExxonMobil ブルーアンモニア事業の 2027 年稼働目標 → **要確認**
- 双日: 「東芝 / Hyundai 系 EV 電池サプライチェーン」 → **要確認 (具体的な提携先の特定難)**
- 全社: 各施策の発表日 / 出資比率 / 投資金額は IR ニュースリリースと照合

---

## システム・運用面

### 環境・デプロイ
- [x] vercel.json に `/sogo-shosha → /sogo-shosha.html` の rewrite 追加済
- [x] sitemap.xml に `/sogo-shosha` エントリ追加済
- [ ] **本番デプロイ前**: 他セクターページ (food-service, saas, ad-agency 等) の cross-nav に「商社」リンクを追加 (現状は sogo-shosha.html 内のみ active)
- [ ] **本番デプロイ前**: index.html (TOP) に商社ダッシュボードへの導線追加

### 認証・公開モード
- 現在は **公開モード** (ログイン不要、`premium-blur` 不使用)
- 他のセクターダッシュボード (entertainment, ad-agency 等) は Firebase auth + Firestore で会員限定
- [ ] 公開モードのまま出すのか、会員限定にするのかを決定 → 決定後に設計実装
  - 会員限定にする場合: `firebase-auth-sogo-shosha.js` の作成と Firestore Document `premiumContent/sogo-shosha-companies` への移行が必要

### コンテンツ・トーン
- [x] 絵文字・`<ul>`・`<code>` は本文中から削除済 (entertainment / ad-agency と一致)
- [x] バッジ色は既存パレット (navy/gold/green/red) のみ使用
- [x] FY24 / FY25 表記の不一致を解消 (全箇所 FY25 = 2026/3 期に統一済)
- [x] 「プレビュー版」バナーは preview-banner として表示中 → **公開時に削除**

### データ更新タイミング
- 株価系 (PER/PBR/時価総額): デプロイ前に `node scripts/sogo-shosha/fetch_yahoo.js && python scripts/sogo-shosha/build_matrix.py` を再実行
- 通期決算系: 2026/3 期 (FY25) 決算発表完了後の Yahoo 更新待ち (タイムラグ数日〜2週間)
- バフェット保有: 四半期ごと (13F 提出時 / 次回 2026年5月15日)

---

## 段階公開オプション

### Option A: 完全公開 (現状そのまま)
- メリット: SEO 露出最大化、リード獲得への寄与大
- リスク: ファクトチェック前のドラフトが拡散する可能性

### Option B: ベータ公開 (URL 直アクセスのみ・ナビ非掲載)
- メリット: 一部ユーザーのフィードバックを得つつ、ファクトチェック完了まで時間を稼げる
- 実装: cross-nav と sitemap からのみ削除し、URL は生かす

### Option C: 会員限定
- メリット: ファクトチェックのレビュー期間中もリード獲得 (登録) を促進できる
- 実装: Firebase auth 移行 (1〜2 日)

森脇さんの判断項目: 「公開タイミング・ローンチ告知」(README §判断ログ で Pending) と併せて確定する。

---

## 既知の制約 (フェーズ2 で補強予定)

詳細: [known-issues.md](known-issues.md)

- 19 KPI のうち 10 項目が未取得 (政策保有株式・大量保有報告件数・自己株買い・自己資本比率・インタレスト・カバレッジ等)
- 補強には EDINET API キーと有報 XBRL の追加パースが必要
- 双日・豊田通商の「資源/非資源」分解は IR で開示されない (構造的制約)
