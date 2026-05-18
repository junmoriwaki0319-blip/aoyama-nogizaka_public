# canonical 監査: スコープ外 TODO（2026-05-18 抽出）

本ドキュメントは fix/news-canonical-alignment ブランチでの本タスク（/news canonical 修正）の
スコープ外として、別ブランチで対応すべき canonical 関連の発見事項を記録する。

## 監査対象
- HEAD: origin/main `9a6595b data: update EDINET reports 2026-05-17T23:21:36Z`
- 監査時刻: 2026-05-18
- 対象: リポジトリ内の全 *.html ファイルの `<link rel="canonical">` と `sitemap.xml` の `<loc>` 全数照合

## 結果サマリ
| 種別 | 件数 | 備考 |
|---|---|---|
| sitemap loc 件数 | 15 | |
| canonical タグ宣言ファイル | 26 | (.html 直 + /index.html ペアを含む) |
| sitemap-canonical 不整合 | 1 | news/index.html (本タスクで修正) |
| sitemap 未登録だが canonical あり | 2 | 下記参照 |
| canonical が .html 拡張子付きURL を指す | 1 | 下記参照 |

## スコープ外発見事項

### A. game-content.html: canonical が他ページ・かつ .html 拡張子付きURL を指している
- 該当行: `./game-content.html: <link rel="canonical" href="https://aoyama-nogizaka.com/entertainment-sector-dashboard.html">`
- 問題点:
  1. 自ページ (game-content) ではなく別ページ (entertainment-sector-dashboard) を canonical 指定 → 意図的な統合なら OK、そうでなければ修正必要
  2. canonical 先が `.html` 拡張子付き URL（本来は拡張子なしの正規 URL を指すべき）
- 推奨対処: game-content ページの存在意義を確認した上で
  - 統合先として残すなら canonical を `https://aoyama-nogizaka.com/entertainment-sector-dashboard` (拡張子なし) に修正
  - 独立ページとして残すなら canonical を `https://aoyama-nogizaka.com/game-content` に変更し sitemap にも追加検討

### B. activist-campaigns.html: sitemap 未登録だが canonical 宣言あり
- 該当行: `./activist-campaigns.html: <link rel="canonical" href="https://aoyama-nogizaka.com/activist-campaigns">`
- 状況: ページは 200 応答するが sitemap.xml に loc がない
- 推奨対処: 公開ページとして扱うなら sitemap に追加。非公開なら noindex 検討

### C. news/index.html 内に末尾スラッシュ付き URL 言及が残存（canonical と乖離）
本タスク (canonical 修正) 中に発見した、`/news/` 末尾スラッシュ付き表現の残存:
- L29: `<meta property="og:url" content="https://aoyama-nogizaka.com/news/">`
- L37: `<link rel="alternate" hreflang="ja" href="https://aoyama-nogizaka.com/news/">`
- L46 周辺: JSON-LD `"url": "https://aoyama-nogizaka.com/news/"`

タスク指示は canonical 行のみ明記のためスコープ外。
- 推奨対処: 別ブランチで `/news` (末尾スラッシュなし) に統一し canonical と完全整合
- 緊急度: 低（Google は canonical を最優先するため、これらは警告誘発の主因ではない見込み）

## 本タスクで対処する1件
| ファイル | 現在の canonical | sitemap の loc | 修正 |
|---|---|---|---|
| news/index.html | https://aoyama-nogizaka.com/news/ | https://aoyama-nogizaka.com/news | canonical を末尾スラッシュなしへ |
