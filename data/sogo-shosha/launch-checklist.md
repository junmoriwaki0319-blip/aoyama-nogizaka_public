# 商社セクター・ダッシュボード — ローンチチェックリスト

**最終更新**: 2026-05-07 (本番ローンチ完了後の状態反映)
**ステータス**: ✅ **本番公開完了** (2026-05-07)

---

## 本番ローンチ完了サマリ (2026-05-07)

| 項目 | 状態 | 詳細 |
|------|------|------|
| `/sogo-shosha.html` 本番反映 | ✅ 完了 | main commit `a4dfab8`, Vercel deployment ID `DzkjSxpAF1AW2Q9e6UEUjdEha61K`, 本番 URL `https://aoyama-nogizaka.com/sogo-shosha` で HTTP 200 (50,650b / 0.46s) |
| 他 5 セクターの cross-nav に商社追加 | ✅ 完了 | food-service / saas / ad-agency / digital-media / entertainment-sector-dashboard の 5 ページに同時反映 (commit `e058de7`)。本番 5/5 反映確認済 |
| sitemap.xml 反映 | ✅ 完了 | PR #10 (commit `e54ea87`)。`/sogo-shosha` 含む 15 URL、全 lastmod 2026-05-07。本番 sitemap Last-Modified: Thu, 07 May 2026 03:20:12 GMT |
| top index.html nav 動線追加 | ✅ 完了 | PR #11 (commit `cb22f4c`)。nav-dropdown / footer-nav / mobile-menu の 3 箇所に `/sogo-shosha.html` 追加。Vercel 自動デプロイ Status Ready / Duration 11s |

### §9 4判断項目の事後確定値

| 判断 | 確定値 | 備考 |
|------|--------|------|
| 判断1: URL 命名 | **B** (sogo-shosha.html) | 既存セクターページのファイル名規約に揃える |
| 判断2: Tier 2 範囲 | **B** (含めない) | 5 大商社 + 双日 + 豊田通商の 7 社で完結。専門商社は別レポートで対応 |
| 判断3: 公開タイミング | **B** (7 社まとめて公開) | データ充足率 47.4% で公開、フェーズ2 で残り KPI 補強 |
| 判断4: ローンチ告知 | **A** (静かに追加) | nav に静かに追加、news/ Special Report 連動なし |

---

## ファクトチェック (REVIEWED 2026-05-07)

[refs/company-narratives.json](refs/company-narratives.json) は `_fact_check_status: REVIEWED`。
2026-05-07 一括修正適用済 (詳細: [fact-check-brief.md](fact-check-brief.md))。

主な修正:
- 三井物産 narrative: 「鉄鉱石 (Vale など)」→「豪州・ブラジル等」、「サハリン 2 撤退後」→「ロシア政府の承継事業に切り替え保有継続」
- 三井物産 strategic: 「Twelve Benefit」→ 一般化「e-fuel / SAF 領域の米スタートアップ」
- 伊藤忠: 「CITIC Pacific」→「CITIC GROUP (中信集団)」 (出資先の正式名)
- 豊田通商: 「CFAO アフリカ 50 ヶ国超」→ 一般化「アフリカ広域の流通網」
- 三菱商事: 三菱食品 TOB「6 月完了予定」→「2025 年 5 月-6 月に実施」
- バフェット保有: 2026-Q1 列追加 (三菱 10.8 / 三井 10.4 / 伊藤忠 10.1 / 丸紅 9.8 / 住友 9.7)

### 自動取得 KPI (Yahoo Finance) — 検証不要
DOE / 配当性向 / ROE / ROIC / PBR / PER / EV·EBITDA / ネット負債/EBITDA / 現預金/時価総額
出所: `yahoo-finance2` npm 経由。基準期は `lastFiscalYearEnd: 2026-03-31` (FY25)。

---

## システム・運用面

### 環境・デプロイ
- [x] vercel.json に `/sogo-shosha → /sogo-shosha.html` の rewrite 追加済 (commit `7e9dc08`)
- [x] sitemap.xml 反映済 (PR #10, commit `e54ea87`)
- [x] 他セクターページ (food-service / saas / ad-agency / digital-media / entertainment) の cross-nav に「商社」追加済 (commit `e058de7`、本番 5/5 反映確認済)
- [x] index.html (TOP) に商社ダッシュボードへの導線追加済 (PR #11, commit `cb22f4c`、nav-dropdown / mobile-menu / footer-nav の 3 箇所)

### 認証・公開モード
- [x] **会員限定モード** で実装済 (Firebase auth + login wall + premium-blur、他セクターと同等)
  - `js/firebase-auth-sogo-shosha.js` 新規作成、`sogo-shosha.html` に 5 つの auth モーダル + login-wall を追加
  - 注: 実質的なコンテンツ・ゲーティング (Firestore 移行) はフェーズ2 で別途対応。現状は UI ゲーティングのみ (静的 JSON は直接 fetch で見える状態)

### コンテンツ・トーン
- [x] 絵文字・`<ul>`・`<code>` は本文中から削除済 (entertainment / ad-agency と一致)
- [x] バッジ色は既存パレット (navy/gold/green/red) のみ使用
- [x] FY24 / FY25 表記の不一致を解消 (全箇所 FY25 = 2026/3 期に統一済)
- [x] 「プレビュー版」バナー削除済 (auth 実装時に login-wall に役割を移譲)

### データ更新タイミング
- 株価系 (PER/PBR/時価総額): デプロイ前に `node scripts/sogo-shosha/fetch_yahoo.js && python scripts/sogo-shosha/build_matrix.py` を再実行
- 通期決算系: 2026/3 期 (FY25) 決算発表完了後の Yahoo 更新待ち (タイムラグ数日〜2週間)
- バフェット保有: 四半期ごと (13F 提出時 / 次回 2026年5月15日)

---

## Open (本番反映後の残タスク)

- [ ] **本番計測**: Playwright で 44×44 検証 (`audits/measure-tap-targets.js` 流用)。AV 対処レーンと並行回避すること
- [ ] **smoke test 32 項目**: `npx playwright test tests/smoke.spec.js` を本番 URL に対して実行
- [ ] **launch announcement**: 必要に応じて news/ Special Report 連動公開 (現状 §9 判断4=A なので保留、後日判断)
- [ ] **research/sogo-shosha-7-20260507 ブランチの remote 削除**: rebase で origin と乖離、本番反映済みで不要
- [ ] **(フェーズ2) Firestore migration による実質的コンテンツ・ゲーティング**: 現状は UI ゲーティングのみ
- [ ] **(フェーズ2) EDINET API キー取得 → 不足 KPI 補強**: 政策保有 / 大量保有件数 / 自己株買い / 自己資本比率 / インタレスト・カバレッジ / 政策保有縮減率 等の 10 項目

---

## 段階公開オプション (採用結果)

採用: **Option A** (完全公開) + 会員限定 UI ゲーティングのハイブリッド

- nav / sitemap / 5 セクター cross-nav にすべて追加 → SEO 露出最大化
- 会員限定 UI で blur + login wall を表示 → 登録動線確保
- ファクトチェックは事前 REVIEWED 済み

(Option B: ベータ公開 / Option C: 完全会員限定 は不採用)

---

## 既知の制約 (フェーズ2 で補強予定)

詳細: [known-issues.md](known-issues.md)

- 19 KPI のうち 10 項目が未取得 (政策保有株式・大量保有報告件数・自己株買い・自己資本比率・インタレスト・カバレッジ等)
- 補強には EDINET API キーと有報 XBRL の追加パースが必要
- 双日・豊田通商の「資源/非資源」分解は IR で開示されない (構造的制約)
- 静的 JSON 配信のため URL 直接アクセスでデータが見える状態 (UI ゲーティングのみ、フェーズ2 で Firestore 移行予定)
