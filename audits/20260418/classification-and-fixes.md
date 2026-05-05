# WCAG 2.5.5 Enhanced (44×44px) — 違反分類と Must-fix 対応結果

- 計測日: 2026-04-18 / 修正適用: 2026-04-19
- 基準: WCAG 2.5.5 Level AAA Enhanced (44×44 CSS px)
- 実測: Playwright + Chromium headless, iPhone SE viewport (375×667)
- 生レポート: [tap-targets-enhanced-44px.md](./tap-targets-enhanced-44px.md) / [tap-targets-enhanced-44px.json](./tap-targets-enhanced-44px.json)
- Lighthouse (2.5.8 AA 基準) ベースライン: [../lighthouse-mobile-20260418.json](../lighthouse-mobile-20260418.json) — score 1.0 (0 violations)

## 経緯

前日セッションの「index のタッチターゲット 27件」という申し送りは、元となる監査ログ/JSONが保存されておらず検証不能だった。本日 Playwright によるライブサイト全12ページ計測を正として実行した結果、**実測値は全ページ合計 21 件**。数字の乖離は hp-backlog.md の P1 項目更新時に明記する。

## 分類サマリ

| 層 | 件数 | 本PR対応 |
|---|---:|---|
| 🔴 Must-fix | **1** | ✅ 修正済み |
| 🟡 Should-fix | 7 | ⏸ hp-backlog.md へ繰越 |
| 🟢 Acceptable-as-is | 13 | 記録のみ |
| **合計** | **21** | |

---

## 🔴 Must-fix (1件) — 本PRで対応

### `/risk-assessment` `.cta-primary` "無料相談を申し込む"

- セレクタ: `body > div.container:nth-of-type(4) > div.cta-section:nth-of-type(5) > div.cta-card > div.cta-links > a.cta-primary:nth-of-type(1)`
- 役割: リード獲得の主要CTA（`cta-card` 内の primary action）
- Before: **181.17 × 43.69** (height 0.31px 不足)
- After: **183.17 × 47** (width +2, height +3.31 / 基準クリア)

**根本原因**: `.cta-links a` 共通ルールは `padding: 12px 28px` で `border` 指定なし。`.cta-secondary` のみ `border: 1px solid rgba(255,255,255,.3)` を持つため +2px で 45.69px に達していたが、`.cta-primary` は border 無しで 43.69px 止まりだった。

**修正内容** (`risk-assessment.html` 228行目):
```css
/* before */
.cta-primary { background: var(--gold); color: var(--white); }
/* after */
.cta-primary { background: var(--gold); color: var(--white); border: 1px solid var(--gold); }
```

**視覚デザインへの影響**: `border-color` = `background-color` (`var(--gold)` = `#9b8b6e`) のため、定義上ピクセル差は発生しない。box-sizing が `content-box` のため外寸が +2px（cta-secondary の外寸と一致）し、タッチ領域だけが拡張される。

**検証スクショ**:
- Before (production): [screenshots/cta-card-before.png](./screenshots/cta-card-before.png) / [screenshots/cta-primary-before.png](./screenshots/cta-primary-before.png)
- After (file://): [screenshots/cta-card-after.png](./screenshots/cta-card-after.png) / [screenshots/cta-primary-after.png](./screenshots/cta-primary-after.png)
- 計測: [screenshots/measurements-before.json](./screenshots/measurements-before.json) / [screenshots/measurements-after.json](./screenshots/measurements-after.json)

**注意事項**:
- 本番配信経路: `vercel.json` rewrites で `/risk-assessment` → `/risk-assessment.html` なので編集対象は `risk-assessment.html` が正。
- 重複ファイル `risk-assessment/index.html` は同じバグを持つが配信経路外。**別タスクで同期するかの判断が必要** → hp-backlog.md へ記載。

---

## 🟡 Should-fix (7件) — hp-backlog.md へ繰越

| # | ページ | 要素 | 現在 | 不足 | 推奨修正 |
|---:|---|---|---:|---:|---|
| 1 | /news/ | `.filter-btn` "すべて" | 80.33×40 | H 4px | `min-height:44px` を `.filter-btn` に追加 |
| 2 | /news/ | `.filter-btn` "論考・分析" | 108.16×40 | H 4px | 同上（共通ルールで一括） |
| 3 | /news/ | `.filter-btn` "プレスリリース" | 133.98×40 | H 4px | 同上 |
| 4 | /news/ | `.filter-btn` "お知らせ" | 94.67×40 | H 4px | 同上 |
| 5 | /news/ | `.filter-btn` "データ・ダッシュボード" | 187.53×40 | H 4px | 同上 |
| 6 | /activist-dashboard.html | 外部リンク "アクティビストの定義" | 283.95×42.39 | H 1.61px | `<details>`内インラインリンク、`line-height` 調整で対応可 |
| 7 | /food-service.html | クロスセル "SaaSレポートを見る →" | 197.44×39.19 | H 4.81px | CTAスタイル適用（`.cta-primary` 等に揃える） |

**まとめ**: news フィルタUIは共通クラスなので1ルール修正で5件解決（実質2タスク）。食品→SaaSクロスセルはデザイン統一の機会。

---

## 🟢 Acceptable-as-is (13件) — 記録のみ

### skip-link × 12件（全ページ共通、false positive）

全ページ共通の「メインコンテンツへスキップ」リンク。

```css
.skip-link { position: absolute; top: -100px; left: 0; ... transition: top 0.2s; }
.skip-link:focus { top: 0; }
```

- `top: -100px` でオフスクリーン配置、`:focus` 時のみ `top: 0` で可視化される標準パターン
- モバイル/タッチデバイスでは focus を取る手段がなく、実質タッチターゲットではない
- Playwright の可視性判定（display/visibility/opacity チェック）では除外できなかったが、**実機モバイルでは永久に不可視 = タップ不能**のため WCAG 2.5.5 の対象外
- 計測スクリプトの改良点として `getBoundingClientRect().bottom < 0 || .right < 0` チェック追加を今後検討

### EDINET 引用リンク × 1件 — /activist-screener.html

- `<details> > p > a` inside `.s-card`: "EDINET" 48.73×20
- 本文中のインライン書誌参照。単独要素化すると段落の意味単位が崩れる
- 周辺のアフォーダンス（`<details>` のサマリ部分）が十分大きく、実操作上の不便はない
- WCAG 2.5.5 例外条項: "Inline: The target is in a sentence or its size is otherwise constrained by the line-height of non-target text" に該当

---

## hp-backlog.md 更新の保留

指定されたパス `/sessions/youthful-confident-mayer/mnt/.claude/hp-backlog.md` は当環境から到達不能（Linux サンドボックスパス、Windows からマウントされていない）。`/c/Users/jun-m/` 配下にも `hp-backlog.md` は存在せず。

**Should-fix 7件 / Acceptable 13件 / Must-fix 1件済 / 本日の audit ファイルパス** の4点を backlog の「P1: モバイル タッチターゲット違反の実測」セクションに反映する必要があるが、正しいパスの指示を待って実施する。

本ファイル（`classification-and-fixes.md`）をそのまま backlog に貼付/ link してもよい構成にしてある。

---

## デプロイ

本PR はローカル変更のみ。Vercel への反映は未実施（`npx vercel --prod`）。デプロイ後、本番URL で `node audits/measure-tap-targets.js` を再実行すれば `/risk-assessment` の Must-fix が消えているはず（Should-fix/Acceptable は残存）。
