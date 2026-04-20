---
name: stash assessment 2026-04-20
description: 2026-04-20 housekeeping 向け stash@{0}/stash@{1} 中身確認と現状反映状況の診断
type: project
---

# Stash 中身確認 (2026-04-20)

前提: 現在ブランチは `chore/tap-targets-audit`（直近 commit `7d78d99 docs: 2026-04-19 daily summary`）。
stash は 2 件。いずれも 2026-04-19 以前に作成された WIP。

本ドキュメントは stash の中身の **診断と提案** のみを記録する。drop/apply/keep の実行は
森脇さんの判断を待ってから行う。

---

## stash@{0}

- **作成時タイトル**: `WIP: playwright.config.js on chore/tap-targets-audit`
- **変更ファイル数**: 1
- **増減**: `playwright.config.js` +2 / -1
- **patch 保存先**: `./audits/20260420/stash-0.patch`

### 差分の要約

`playwright.config.js` の `baseURL` を `process.env.AUDIT_BASE || 'https://aoyama-nogizaka.com'`
に変更し、Preview URL で tap-targets 等を計測できるようにするもの。

### 現ブランチへの反映状況

**既に取り込み済み**（コメント文言まで完全一致）。現 HEAD の `playwright.config.js` の
該当行は `baseURL: process.env.AUDIT_BASE || 'https://aoyama-nogizaka.com'` になっており、
stash 内容と完全一致する。恐らく `21142ee docs: tap-targets audit methodology + A+B result snapshots + AUDIT_BASE env`
で commit 済み。

### 提案

**drop 推奨**。既に HEAD に反映済みで、stash を適用しても no-op または空コミットになる。
適用漏れで失うものは何もない。

---

## stash@{1}

- **作成時タイトル**: `WIP on chore/tap-targets-audit before prod-redeploy switchover (2026-04-19)`
- **変更ファイル数**: 4
- **増減**: 合計 +2,034 / -1,227 行
  - `css/mobile-touch-font-fix.css` +6 / -6
  - `css/mobile-touch-font-fix.min.css` +1 / -1
  - `playwright.config.js` +2 / -1
  - `tests/tap-targets-report.json` +2,025 / -1,219
- **patch 保存先**: `./audits/20260420/stash-1.patch`

### 差分の要約

tap-targets Lighthouse 監査で `44px` では不足と判定されたアイコンボタン群を
`min-width: 44px → 48px` に昇格する修正。対象は `.btn-cta-register` / `.btn-gold`
/ `.watchlist-add-btn` / `.nav-hamburger` / `.watchlist-remove-btn` / `.watchlist-toggle-btn`
/ `td a` / `.edinet-link` の計 6 箇所（.css 側）+ min.css 反映 + playwright.config
（stash@{0} と同一変更）+ tap-targets-report.json 再生成結果。

### 現ブランチへの反映状況

**ほぼ取り込み済み（ただし残件あり）**。
- 現 HEAD の `css/mobile-touch-font-fix.css` は `min-width: 48px` が 6 箇所にヒット、
  `min-width: 44px` が **1 箇所残存** (`css/mobile-touch-font-fix.css:619` の
  `table a, .activist-table a` セレクタ内)。stash@{1} ではこの 1 箇所は変更対象外
  だったため、stash を apply しても解消しない。
- `playwright.config.js` は前述の通り反映済み。
- `tests/tap-targets-report.json` は現ファイル 2,022 行、stash の新版は約 2,025 行相当
  でほぼ同規模。直近 commit `21142ee` 内で再生成版が commit 済みと推定されるが、
  内容が完全一致かは未 diff。stash を apply すると JSON の一部が巻き戻る恐れあり。

### 提案

**drop 推奨**。
- 44→48 の本体 CSS 変更は既に commit `b2823b4 / acf9aaa / 90812ec` で反映済み。
- 残る `table a, .activist-table a` の 44px は stash の対象外なので、別 issue として
  今日の T2 違反セレクタ抽出のアウトプットで扱えばよい。
- `tests/tap-targets-report.json` を apply すると、直近 commit 済みの最新レポートが
  上書きされるリスクがある。

もし保険として残すなら keep でも害はないが、残し続けると毎朝のチェックで混乱する。

---

## 総合サマリ（森脇さんへの確認）

| stash | 要点 | 推奨 |
|---|---|---|
| stash@{0} | playwright.config AUDIT_BASE 対応。現 HEAD と完全一致。 | **drop** |
| stash@{1} | 44→48 CSS 修正 + playwright.config + tap-targets-report.json。本体 CSS は commit 済み、JSON 上書きリスクあり。 | **drop** |

残存する `css/mobile-touch-font-fix.css:619` の 44px は本日 T2 の違反抽出で拾う想定
（stash を apply しても解消しないため、stash 処理とは切り離す）。

判断指示 (drop / apply / keep) を頂ければ、フェーズ2 housekeeping に進みます。

---

## 実施結果

- 2026-04-20 両 stash drop 完了（patch 保全済み）
  - `stash@{1}` → dropped (0f477f5). 保全 patch: `./audits/20260420/stash-1.patch`
  - `stash@{0}` → dropped (37175dd). 保全 patch: `./audits/20260420/stash-0.patch`
- drop 後 `git stash list` 空確認済み
