# 44×44 タップターゲット violation 分類 (2026-05-07)

## 計測ベース

- 計測元: PR #6 ブランチ `chore/audit-20260505` の `audits/20260505/tap-targets-44x44-*.json` (Playwright, viewport 375x667)
- 計測ログ: `audits/20260505/_playwright-44x44.log`
- 合計 violation: **75件 / 13ページ**

### ページ別件数（再現確認済）

| ページ | violation |
|--------|----------|
| home (`/`) | 1 |
| team | 1 |
| privacy | 1 |
| news (`/news/`) | 6 |
| activist-dashboard | 2 |
| risk-assessment | 1 |
| activist-screener | 2 |
| food-service | 2 |
| saas | 1 |
| ad-agency | 1 |
| digital-media | 1 |
| entertainment-sector-dashboard | 1 |
| **news/activist-shareholder-proposals-japan** | **55** |

合計 75件。

## 型別分類

合計 6 つの修正型 (A〜F) と 1 つの known-issue 型 (G)。

### Type A — `.skip-link` (13件 / 全 13 ページ)

- 該当 violation: 13件 (各ページ1件、全ページに横断)
- 計測サイズ: 214×37 (ほとんどのページ) / 195×21 (team)。幅は 44px 以上、**高さのみ不足**（deficit h: 7〜23）
- 定義箇所: 21 個の HTML ファイルで `<style>` ブロックにインライン定義（`activist-dashboard.html` と `activist-dashboard/index.html` の両方等、ミラー含む）
- 既存ルール: `position: absolute; top: -100px; left: 0; ... padding: 8px 16px; font-size: 14px;`
- 修正方針: 既存セレクタに `min-height: 44px; display: inline-flex; align-items: center;` を追加
- 影響: `position: absolute; top: -100px` のため通常時は画面外。`:focus` 時のみ可視化されるが、可視化時にちょうど 44px の高さになる。レイアウト崩れリスク低
- 対応コミット: Commit 1

### Type B — `.link-primary` / `.link-secondary` (36件 / news/activist-shareholder-proposals-japan のみ)

- 該当 violation: `.link-primary` 20件 + `.link-secondary` 16件 = 36件
- 計測サイズ: 86〜193 × 34〜36 (高さ不足 8〜10px)
- 定義箇所: `news/activist-shareholder-proposals-japan.html` の `<style>` ブロック (line 173)
- 既存ルール: `display: inline-flex; align-items: center; padding: 8px 18px; font-size: 0.75rem; font-weight: 600;`
- 修正方針: 既存セレクタに `min-height: 44px;` を追加
- 影響: `.case-links` flex container 内のボタン群。flex-wrap: wrap で並ぶため、min-height 44 にしても外周レイアウト変化は微小
- 対応コミット: Commit 2

### Type C — `.cta-link` (.cta-link-gold / .cta-link-outline) (2件 / news/activist-shareholder-proposals-japan のみ)

- 該当 violation: `.cta-link.cta-link-gold` 1件 + `.cta-link.cta-link-outline` 1件 = 2件
- 計測サイズ: 110×36 / 185×34 (高さ不足 8〜10px)
- 定義箇所: `news/activist-shareholder-proposals-japan.html` の `<style>` ブロック (line 195)
- 既存ルール: `display: inline-flex; align-items: center; padding: 8px 18px; font-size: 0.75rem; font-weight: 600;`
- 修正方針: 既存セレクタに `min-height: 44px;` を追加
- 影響: `.cta-links` flex container 内の trend-card-cta CTA ボタン。同じく min-height 化で安全
- 対応コミット: Commit 2 （Type B と同ファイル・同パターンのため統合）

### Type D — `.filter-btn` (5件 / news/index.html のみ)

- 該当 violation: 5件 (filter-btn 4 + filter-btn.active 1)
- 計測サイズ: 80〜187 × 40 (高さ不足 4px のみ)
- 定義箇所: `news/index.html` の `<style>` ブロック (line 147、`@media (max-width: 768px)` 内)
- 既存ルール: `.filter-btn { min-height: 40px; padding: 8px 16px; }` (mobile override)
- 修正方針: `min-height: 40px` → `min-height: 44px` に変更
- 影響: モバイル幅でのみ適用。デスクトップは line 181 の base ルール `padding: 7px 20px` のままで影響なし
- 対応コミット: Commit 3

### Type E — `.resource-list a` (10件 / news/activist-shareholder-proposals-japan のみ)

- 該当 violation: `.resource-list > li > a` 10件
- 計測サイズ: 158〜304 × 20 or 41 (高さ不足 3〜24px)
- 定義箇所: `news/activist-shareholder-proposals-japan.html` の `<style>` ブロック (line 210)
- 既存ルール: `font-size: 0.85rem; font-weight: 600; color: var(--navy); text-decoration: none;` (display 未指定 → inline)
- 修正方針: `display: inline-block; padding: 12px 0; min-height: 44px;` を追加。`<li>` 内で `<a>` の下に `<span class="resource-desc">` が続くため、`<a>` を block 化して li 全体を埋めるのは避ける
- 影響: 各 `<li>` の上下マージンが現状より +24px 増える可能性。意図的な余白増として許容
- 対応コミット: Commit 4

### Type F — food-service の inline-styled CTA `<a>` (1件)

- 該当 violation: 1件 ("SaaSレポートを見る →" ボタン)
- 計測サイズ: 197×39 (高さ不足 5px)
- 定義箇所: `food-service.html` line 813、`food-service/index.html` line 792（同一の inline `style` 属性）
- 既存スタイル: `style="background:#1a2d4f;color:#fff;text-decoration:none;padding:10px 28px;font-size:0.8rem;letter-spacing:0.5px;transition:opacity 0.2s;white-space:nowrap;"`
- 修正方針: `style` に `display:inline-flex;align-items:center;min-height:44px;` を追加
- 影響: 単一要素のみ。レイアウト変化は数ピクセル
- 対応コミット: Commit 5

### Type G (KNOWN-ISSUE) — Inline 散文中のリンク (8件)

WCAG 2.5.5 Enhanced (44×44) には明示的な例外あり:
> **Inline**: The target is in a sentence or its size is otherwise constrained by the line-height of non-target text.

これに該当する prose-inline citation links は本タスクでは修正しない（修正すると line-height が崩れて文章レイアウトが歪むため）。

- activist-dashboard: 1件 (`<details><p>` 内のリンク "アクティビストの定義｜...")
- activist-screener: 1件 (`<details><p>` 内のリンク "EDINET")
- news/activist-shareholder-proposals-japan: 6件
  - `.report-meta` 内の出典リンク (× 2): "三井住友信託銀行調べ（日経）" / "大和総研"
  - `.trend-card > p` 内のインライン引用 (× 4): "114社・399議案" / "EY-Parthenon" / "大和総研 2025年10月" / "選択"

合計 **8件は known-issue として残置**。修正後 violation 件数は 75 - (A+B+C+D+E+F = 67) = **8件**を上限とする。

## 修正前後の見込み

| 型 | 件数 | 修正対象ファイル |
|----|------|-----------------|
| A | 13 | 21 HTML (`.skip-link` インライン定義) |
| B+C | 38 | news/activist-shareholder-proposals-japan.html |
| D | 5 | news/index.html |
| E | 10 | news/activist-shareholder-proposals-japan.html |
| F | 1 | food-service.html, food-service/index.html |
| **修正対象計** | **67** | |
| G (known-issue) | 8 | (修正しない) |
| **合計** | **75** | |

修正後想定: 75件 → 8件 (known-issue のみ残置)。Playwright 再計測は Step 3a-recheck 完了後に実施予定。

## 並列実行コンテキスト

- 2026-05-07: Step 3a-recheck (perf 中央値再計測、13 URL × 3 ラン × Lighthouse mobile) が並列実行中（30分）
- Step 3b の Step 1〜4 (本ファイル含む classification + CSS/HTML 修正) は並列で進行
- Step 5 (Playwright 再計測) は Step 3a-recheck 完了報告まで保留 (Chrome 同時起動による計測ブレ回避)
