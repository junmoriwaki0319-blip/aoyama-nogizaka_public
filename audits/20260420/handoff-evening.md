# Handoff — 2026-04-20 evening

## Housekeeping (フェーズ1〜2)

### Stash

| stash | 作成時刻 | 内容 | drop 日時 | 保全 patch |
|---|---|---|---|---|
| stash@{0} | 2026-04-19 以前 | `playwright.config.js` の AUDIT_BASE 対応 (HEAD と一致・取り込み済み) | **2026-04-20 drop 済 (SHA 37175dd)** | [stash-0.patch](./stash-0.patch) |
| stash@{1} | 2026-04-19 switchover 前 | 44→48 CSS + playwright.config + tap-targets-report.json | **2026-04-20 drop 済 (SHA 0f477f5)** | [stash-1.patch](./stash-1.patch) |

drop 後 `git stash list` 空確認済み。判断根拠は [stash-assessment.md](./stash-assessment.md) 参照。

### ルート散乱ファイル退避

以下を `audits/20260419/raw/` へ退避 (untracked のまま `mv`):

- `lh-run-1.json` / `lh-run-2.json` / `lh-run-3.json`
- `lighthouse-mobile.json` / `lighthouse-mobile-2.json` / `lighthouse-mobile-3.json`

### .gitignore housekeeping

| branch | commit | push 済み | 備考 |
|---|---|---|---|
| `chore/gitignore-housekeeping-20260420` | **5c0035e** `chore(gitignore): ignore root lighthouse outputs and minified css` | ✅ origin | `/lh-run-*.json`, `/lighthouse-mobile*.json`, `**/*.min.css` を追加。PR 未作成 |

## 本日のベースブランチ

**`chore/audit-20260420-baseline`** (main 9b5d82c から派生)

### Commit 一覧

| # | Task | Commit | Message |
|---|---|---|---|
| 0 | 既知44px修正 | **c7c8b87** | `fix(a11y): bump table link tap target from 44px to 48px (known residual)` |
| 1 | T1 Lighthouse Mobile baseline | **c214052** | `chore(audit): lighthouse mobile baseline 20260420 (13 urls)` |
| 2 | T5 last-updated 統一 | **0b609ca** | `feat(ui): unify last-updated display across sector pages` |
| 3 | T2 violations by selector | **4e95cfe** | `chore(audit): tap/contrast/viewport violations by selector 20260420` |
| 4 | T3 data dropdown a11y test | **6f76deb** | `test(a11y): data dropdown keyboard/aria coverage` |
| 5 | T4 next sector decision memo | **da3acf2** | `docs(sector): next sector decision memo 20260420` |
| 6 | T6 ogp/canonical/sitemap | **9fde213** | `chore(audit): ogp/canonical/sitemap coverage 20260420` |

Push 状況: **push 予定 (本ドキュメント push と同時に upstream 設定)** — prod は触らない / `vercel --prod` 未実行。

---

## T1 Lighthouse Mobile — スコア表要約

Form factor: mobile / Chrome headless / Lighthouse CLI (v13 系)
詳細: [summary-mobile.md](./summary-mobile.md)

| 指標 | avg | min | <90 件数 |
|---|---:|---:|---:|
| Performance | 68.9 | 47 | **13/13 (全滅)** |
| Accessibility | 89.4 | 86 | 7/13 |
| Best Practices | 95.4 | 92 | 0/13 |
| SEO | 100 | 100 | 0/13 |

### Perf 赤信号ページ (特に要注意)

- `news/activist-shareholder-proposals-japan` perf=47 / LCP 4.3s (poor) / TBT 3,470ms (poor)
- `home` perf=52 / LCP 11.1s (poor) / CLS 0.26 (poor)
- `activist-dashboard` perf=52 / TBT 1,220ms (poor)
- `risk-assessment` perf=59 / CLS 0.451 (poor)

### A11y <90 ページ

activist-dashboard (87) / ad-agency (88) / digital-media (88) / entertainment-sector-dashboard (88) / food-service (88) / risk-assessment (86) / saas (88)

### 備考

- Lighthouse CLI の Chrome temp-dir cleanup で Windows EPERM が発生したが、レポート本体は全 13 書き出し済み。

---

## T2 Violations by Page — 件数内訳

詳細: [violations-by-page.md](./violations-by-page.md)

### ページ別 (上位)

| slug | violations |
|---|---:|
| activist-dashboard | **259** |
| news-activist-shareholder-proposals-japan | **222** |
| risk-assessment | 70 |
| activist-screener | 39 |
| food-service | 32 |
| news | 30 |
| team | 29 |
| ad-agency | 28 |
| digital-media | 28 |
| saas | 27 |
| entertainment-sector-dashboard | 25 |
| privacy | 22 |
| home | 7 |
| **TOTAL** | **818** |

### audit 別

| audit | count | 注 |
|---|---:|---|
| color-contrast | **637** | axe-core + Lighthouse 重複含む |
| region | 139 | axe-core (ランドマーク外 content) |
| landmark-one-main | 11 | axe-core |
| skip-link | 10 | axe-core |
| heading-order | 7 | axe-core |
| select-name | 5 | axe-core |
| scrollable-region-focusable | 5 | axe-core |
| empty-heading | 3 | axe-core |
| page-has-heading-one | 1 | axe-core |

### ⚠️ カバレッジ注意

Lighthouse v13 系で **`tap-targets` / `viewport` / `font-size`** が audit として JSON に含まれない状態。
`meta-viewport` のみ動作 (全13 URL score=1)。タップターゲット系は本レポート外で Playwright 実測が必要。

---

## T3 Data Dropdown a11y — 失敗数

詳細: [dropdown-a11y-result.md](./dropdown-a11y-result.md)

**失敗テスト数: 4 / 5**

- ❌ aria-expanded が click で true/false に切り替わる
- ❌ Tab でメニュー項目を順に巡回できる
- ❌ Shift+Tab でトグルへ戻る
- ❌ Esc でメニューが閉じてトグルへフォーカス復帰
- ✅ 外部クリックでメニューが閉じる (偽陽性: そもそも click では開かない hover 実装)

根本原因: `.nav-dropdown` は CSS `:hover` のみで JS ハンドラ・ARIA 属性が一切無い。

---

## T4 Next Sector — 推奨1位

詳細: [../../decision-memos/20260420-next-sector.md](../../decision-memos/20260420-next-sector.md)

**推奨1位: 金融セクター (銀行 / 地銀中心)** — スコア 19/20

理由要約: 4 観点（大量保有件数・低 PBR 社数・テンプレ流用率・営業導線親和性）全てで高評価の唯一の候補。
地銀再編期のアクティビスト大量保有が最多、東証低 PBR 開示要請への IR コンサル需要と合致、
投資事業（エンゲージメント候補発掘）とコンサル事業（IR・PBR 改善助言）を同時に取れる。

URL slug 案 (推奨): `/financial.html`  
想定 KPI 7 個: ROE / PBR / BIS / 貸出金成長率 / NPL ratio / NIM / OHR

---

## T5 Last-Updated 統一 — 追加ページ数

詳細: [last-updated-status.md](./last-updated-status.md)

**追加ページ数: 5** (food-service / saas / ad-agency / digital-media / entertainment-sector-dashboard)

全5ページの `.report-meta` 内に `<span class="last-updated">最終更新: 2026-04-20</span>` を追加。
CSS は新規 [assets/css/sector-common.css](../../assets/css/sector-common.css) に集約、
各ページの `<head>` で `<link rel="stylesheet" href="/assets/css/sector-common.css">` を追加。

---

## T6 Meta / Sitemap Coverage — 欠落数

詳細: [meta-coverage.md](./meta-coverage.md) / [sitemap-coverage.md](./sitemap-coverage.md)

### Meta (OGP / Twitter / canonical)

- 対象 13 URL × 7 フィールド = 91 項目
- **欠落: 0 (全項目埋まっている)**

### Sitemap.xml

- sitemap.xml 内の `<loc>` 数: **10**
- 対象 13 URL のうち包含: **9** (.html 拡張子の有無を正規化して比較)
- **欠落 URL 数: 4**
  - `/ad-agency.html`
  - `/digital-media.html`
  - `/entertainment-sector-dashboard.html`
  - `/news/activist-shareholder-proposals-japan.html`

sitemap 側にあって対象 13 に無い URL: `/news/activist-report` (1 件)。

---

## 申し送り・注意事項

### 本日のブランチ派生時点の既知課題 (main に未 merge)

`chore/tap-targets-audit` ブランチ（PR #4 merged, PR #5 ready で前日作業）で 44→48 の tap-target
修正が複数 commit されているが、**main には未 merge**。結果として本日の `chore/audit-20260420-baseline`
派生時点で `css/mobile-touch-font-fix.css` に `min-(width|height): 44px` が 40+ 箇所残存している。
本日 commit `c7c8b87` で修正したのは 613-622 行の `table a, .activist-table a` ブロック 1 箇所のみ。
残りは `chore/tap-targets-audit` / 関連 PR のマージで解消される想定。

### Lighthouse の tap-targets 監査の取り扱い

Lighthouse v13 系で `tap-targets` が audit 出力から消えた。Playwright ベースの
tap-targets 計測（前日の [../20260419/tap-targets-preview-after.json](../20260419/tap-targets-preview-after.json)）を
今後の baseline として運用する必要がある。本日は再計測未実施。

### データドロップダウンの a11y 改修

T3 の結果から JS + ARIA の実装が必要。別 PR で対応推奨。
推奨実装スケッチは [dropdown-a11y-result.md](./dropdown-a11y-result.md) の末尾参照。

### 完了フラグ

- [x] Phase 1: stash 中身確認 + assessment
- [x] Phase 2A: stash drop (両方)
- [x] Phase 2B: ルート散乱退避 + .gitignore + 別ブランチ push
- [x] Phase 3A: chore/audit-20260420-baseline 派生
- [x] Phase 3B: 既知44px修正
- [x] T1 / T2 / T3 / T4 / T5 / T6 全実施
- [x] chore/audit-20260420-baseline の push (2026-04-21 00:00 頃完了 / 全 8 commit upstream 反映済み)
