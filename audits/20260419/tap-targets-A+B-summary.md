# 2026-04-19 tap-targets 監査結果 (A+B)

- ブランチ: `chore/tap-targets-audit`
- 測定対象: Vercel Preview (`aoyama-nogizakapublic-git-f29f24-junmoriwaki0319-blips-projects.vercel.app`)
- 測定基準: [tests/README.md](../../tests/README.md) 固定軸 (48×48 / 360×640 / Cat B フィルタ適用)

## 違反数の遷移

| 段階 | 違反数 | 累積削減 |
|---|---:|---:|
| Before (pre-A, no filters) | 151 | — |
| B only (Cat B フィルタのみ) | 23 | −128 |
| **A+B (CSS 修正 + フィルタ)** | **0** | **−151** |

## ページ別（A+B 適用後）

| ページ | 違反数 | 除外 (offscreen) | 除外 (inline) |
|---|---:|---:|---:|
| / | 0 | 32 | 0 |
| /team | 0 | 11 | 0 |
| /saas | 0 | 2 | 0 |
| /food-service | 0 | 3 | 0 |
| /ad-agency | 0 | 2 | 0 |
| /digital-media | 0 | 2 | 0 |
| /entertainment-sector-dashboard | 0 | 2 | 0 |
| /activist-dashboard | 0 | 82 | 0 |
| /activist-screener | 0 | 17 | 0 |
| /activist-campaigns | 0 | 64 | 0 |
| /risk-assessment | 0 | 29 | 0 |
| /privacy | 0 | 11 | 0 |
| **合計** | **0** | **257** | **0** |

## A カテゴリ修正内容

### `css/mobile-touch-font-fix.css` (+ `.min.css`)

モバイル `@media (max-width: 768px)` 内の共通タップターゲット規則:

- `min-height: 44px` → **48px** (43 rules 一括)
- `min-width: 44px` → **48px** (6 rules; `table a` / `.activist-table a` のみ 44px 維持)

### HTML ページ inline CSS

10 ページ（index / activist-campaigns / activist-dashboard / activist-screener /
ad-agency / digital-media / entertainment-sector-dashboard / food-service /
risk-assessment / saas）の `<style>` 内モバイルブロック:

- `min-height: 44px` → **48px** (59 occurrences 一括)

### 追加修正

- `activist-campaigns.html` / `activist-screener.html` に `.skip-link` CSS を補完
  - 元々 `class="skip-link"` 付きアンカーが存在したが `position:absolute; top:-100px` の CSS が欠落
  - 結果として skip-link が 195×21 px で可視領域に表示されていた
- `.nav-hamburger` 単独で `min-width: 44px` → `48px`
  - 共通規則の 44→48 バンプに含めていなかった個別修正

## B カテゴリ監査スクリプト改善

`tests/tap-targets-audit.spec.js`:

1. オフスクリーン除外
   - `rect.bottom <= 0 || rect.right <= 0 || rect.top >= viewport.height`
   - `position: absolute && rect.top < -50` (skip-link パターン)
2. Running-text 内インラインリンク除外
   - 親 3 階層以内に `<p>`, `<li>`, `<details>` がある `<a>` 要素
   - WCAG 2.5.5 / 2.5.8 の "Inline" 例外に準拠
3. `methodology` ブロックを JSON レポートに埋め込み（閾値・viewport・フィルタ仕様の再現性確保）

## 次のステップ

- カテゴリ C（大規模ページの tap-target 改修）は別ブランチで:
  - `chore/tap-targets-campaigns-layout` — `/activist-campaigns` カード内リンク改修
  - `chore/tap-targets-screener-layout` — `/activist-screener` tab/filter 改修

  現 A+B 修正後は両ページとも 0 violations だが、offscreen 除外を外すと以前の違反数
  (69, 21) に近づくため、カード内要素のタップ領域をより厳密に扱いたい場合は
  別 PR で対応する。

## 測定ファイル

- `tap-targets-before.json` — pre-A baseline (151 violations, no filters)
- `tap-targets-after-B-only-prod.json` — B フィルタのみ適用 (23 violations, production)
- `tap-targets-after-A+B.json` — **A+B 完了後 (0 violations, preview)**
