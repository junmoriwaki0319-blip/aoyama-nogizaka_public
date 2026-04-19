# Playwright Test Suites

## tap-targets-audit.spec.js — モバイル tap-target 監査

### 目的

Lighthouse の `tap-targets` 監査を Playwright に置き換えた実装。Windows 環境で
`npx lighthouse` が Chrome launcher の tmpdir EPERM で実行できない問題を回避しつつ、
Lighthouse と同じしきい値で「48×48 CSS px 以上」を検証する。

### 測定軸（固定基準）

| 項目 | 値 | 根拠 |
|---|---|---|
| Viewport | **360 × 640 CSS px** | Lighthouse モバイル emulation の古典的寸法 |
| Threshold | **48 × 48 CSS px** | Lighthouse default / Google Mobile 推奨 |
| Target selector | `a[href], button, input:not([type="hidden"]), select, textarea, [role="button"], [role="link"], [onclick], [tabindex]:not([tabindex="-1"])` | Lighthouse audit の対象と同等 |
| Page list | 下記 12 ルート | ルート直下の主要ページ（/news/ サブディレクトリは対象外） |

対象ページ:
```
/  /team  /saas  /food-service  /ad-agency  /digital-media
/entertainment-sector-dashboard  /activist-dashboard
/activist-screener  /activist-campaigns  /risk-assessment  /privacy
```

### Category B 除外フィルタ（false positive 除外）

Lighthouse と挙動を合わせるため、以下を violation カウントから除外し、
透明性のため `excluded.offscreen` / `excluded.inline` として別記録:

1. **オフスクリーン要素**
   - `rect.bottom <= 0` または `rect.right <= 0` または `rect.top >= viewport.height`
   - `position: absolute` かつ `rect.top < -50` (skip-link パターン)
   - `visibility:hidden` / `display:none`

2. **Running-text 内のインラインリンク (WCAG 2.5.5 Inline 例外)**
   - 親が `<p>`, `<li>`, `<details>` の直系 `<a>` 要素
   - 本文中の参照リンクは単独のタップターゲットではなく、文脈上の文字列の一部として扱う

### 実行

```bash
# 本番 (aoyama-nogizaka.com) に対して
npx playwright test --project=tap-targets

# Vercel Preview に対して (AUDIT_BASE 環境変数で上書き)
AUDIT_BASE="https://aoyama-nogizakapublic-git-XXXX-junmoriwaki0319-blips-projects.vercel.app" \
  npx playwright test --project=tap-targets
```

### 出力

- `tests/tap-targets-report.json` — 最新実行の詳細レポート
  - `summary` ... ページ別集計
  - `pages[path].items` ... 違反要素一覧（selector / 寸法 / href / label）
  - `pages[path].excluded.{offscreen,inline}` ... 除外記録（監査結果の透明性担保）
- `audits/YYYYMMDD/tap-targets-*.json` — 日付別スナップショット

### 過去の測定結果との差異について

| 日付 | ブランチ | 閾値 | Viewport | フィルタ | 違反数 | 備考 |
|---|---|---:|---|---|---:|---|
| 2026-04-18 | main | 44×44 (WCAG 2.5.5 AAA) | 375×667 (iPhone SE) | skip-link / inline 除外 | 21 | Playwright spec 不在、ad-hoc script |
| 2026-04-19 (before) | main | 48×48 (Lighthouse) | 360×640 | 無 | 151 | 初回 Playwright spec 実行 |
| 2026-04-19 (B only) | main | 48×48 | 360×640 | **Cat B** | 23 | フィルタ改善のみ |
| 2026-04-19 (A+B) | chore/tap-targets-audit | 48×48 | 360×640 | **Cat B** | **0** | CSS 修正 + フィルタ |

**基準を固定する理由**: 閾値 44 ⇄ 48、viewport 375 ⇄ 360、フィルタ適用有無の組合せで数字は
大きく変動する。今後は本 README の「測定軸」を基準として、過去レポートとの比較は
「同じ軸」で取り直した数字でのみ行うこと。異なる軸同士の比較は意味を持たない。
