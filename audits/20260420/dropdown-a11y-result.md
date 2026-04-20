# データ ドロップダウン a11y テスト結果 — 2026-04-20

Target URL: https://aoyama-nogizaka.com/ (main nav `.nav-dropdown`)
Test file: [tests/a11y/data-dropdown.spec.ts](../../tests/a11y/data-dropdown.spec.ts)
Command: `npx playwright test tests/a11y/data-dropdown.spec.ts --project=chromium`

## サマリ

| 項目 | 件数 |
|---|---|
| 合計テスト | 5 |
| 成功 (passed) | 1 |
| 失敗 (failed) | **4** |

## ケース別

| # | ケース | 結果 | FIXME |
|---|---|---|---|
| 1 | `aria-expanded` が click で true/false に切り替わる | ❌ FAIL | toggle に `aria-expanded` 属性が未付与 |
| 2 | Tab でメニュー項目を順に巡回できる | ❌ FAIL | Enter では開かない (CSS :hover のみ)。`aria-controls` も未定義 |
| 3 | Shift+Tab でトグルへ戻る | ❌ FAIL | 2 と同じ根本原因 |
| 4 | Esc でメニューが閉じてトグルへフォーカス復帰 | ❌ FAIL | keyboard 開閉ロジック未実装 / Esc ハンドラ無し |
| 5 | 外部クリックでメニューが閉じる | ✅ PASS | そもそも click では開かない (hover のみ) ため "閉じている" が通過する偽陽性 |

## 根本原因（現実装の不足）

- `a.nav-dropdown-toggle` は `href="#"` の `<a>` タグで、ボタンではない
- JS によるクリック/キーボードハンドラが **一切実装されていない**
- `aria-expanded` / `aria-controls` / `aria-haspopup` が付与されていない
- `<ul>` 側に `role="menu"` or `aria-labelledby` がない
- 展開ロジックは `.nav-dropdown:hover .nav-dropdown-menu { display: block; }` の **CSS 擬似クラスのみ**

→ モバイル / キーボードユーザーには完全に使えない状態。SR も展開可能要素として認識しない。

## 推奨修正（別 PR で対応予定）

1. `<a href="#">` を `<button type="button">` に置換し、
   `aria-haspopup="menu"` / `aria-expanded="false"` / `aria-controls="navDataMenu"` を付与
2. `<ul>` に `id="navDataMenu"` / `role="menu"` を付与、各 `<li><a>` を `role="menuitem"` に
3. 小さな JS ハンドラ:
   - click / Enter / Space で aria-expanded トグル & .open クラス付与
   - 矢印キー / Tab でのロービング
   - Esc で閉じて toggle にフォーカス復帰
   - `document` 外部クリック (mousedown) で閉じる
4. `.nav-dropdown.open .nav-dropdown-menu { display: block; }` を追加し、`:hover` とも両立

## Artifacts

- スクリーンショット: `test-results/a11y-data-dropdown-*/test-failed-*.png`
- Playwright config 更新: [playwright.config.js](../../playwright.config.js) に `chromium` project 追加（`testMatch: 'a11y/**/*.spec.ts'`）
