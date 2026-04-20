# Last-Updated 表示の統一状況 — 2026-04-20

対象は以下の 5 セクター分析ダッシュボード:

| page | 変更前 (`最終更新` grep 件数) | 変更後 対応 | 追加位置 |
|---|---|---|---|
| [food-service.html](food-service.html) | 0 (欠落) | `<span class="last-updated">最終更新: 2026-04-20</span>` 追加 | `.report-meta` 内 |
| [saas.html](saas.html) | 0 (欠落) | 同上 追加 | `.report-meta` 内 |
| [ad-agency.html](ad-agency.html) | 0 (欠落) | 同上 追加 | `.report-meta` 内 |
| [digital-media.html](digital-media.html) | 0 (欠落) | 同上 追加 | `.report-meta` 内 |
| [entertainment-sector-dashboard.html](entertainment-sector-dashboard.html) | 0 (欠落) | 同上 追加 | `.report-meta` 内 |

## CSS 集約

- 新規ファイル: [assets/css/sector-common.css](assets/css/sector-common.css)
  - `.last-updated` セレクタのみ定義（枠線バッジ風 / モバイルでも可読な font-size）
- 5 ページそれぞれの `<head>` で `/css/mobile-nav.css` の直後に
  `<link rel="stylesheet" href="/assets/css/sector-common.css">` を追加

## 付記（対象外だが既存で `最終更新` を持つページ）

| page | grep 件数 | 備考 |
|---|---|---|
| activist-dashboard.html | 1 | `.data-source` 内に EDINET 最終更新時刻を動的表示（`#lastUpdated`）|
| activist-dashboard/index.html | 1 | 上と同じエントリ |
| news/activist-report.html | 1 | 記事本文内の「最終更新」記述 |

これらは既存実装で既に対応済みのため T5 の対象外。フォーマット統一は段階的対応を推奨（本日は対応外）。

## 追加ページ数

**5 ページ** (food-service, saas, ad-agency, digital-media, entertainment-sector-dashboard)
