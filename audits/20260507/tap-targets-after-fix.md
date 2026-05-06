# WCAG 2.5.5 Enhanced (44×44px) タッチターゲット違反レポート

- 計測日: 2026-04-18
- Viewport: 375×667 (iPhone SE), Chromium headless
- 閾値: width < 44 OR height < 44
- 対象セレクタ: `a,button,input[type="button"],input[type="submit"],input[type="reset"],[role="button"],[role="link"],[onclick],label[for]`
- 除外: display:none, visibility:hidden/collapse, opacity:0, 0×0矩形

## サマリ

| ページ | 違反数 |
|---|---:|
| / | 0 |
| /team | 0 |
| /privacy | 0 |
| /news/ | 0 |
| /activist-dashboard.html | 1 |
| /risk-assessment.html | 0 |
| /activist-screener.html | 2 |
| /food-service.html | 0 |
| /saas.html | 0 |
| /ad-agency.html | 0 |
| /digital-media.html | 0 |
| /entertainment-sector-dashboard.html | 0 |
| /news/activist-shareholder-proposals-japan.html | 6 |
| **合計** | **9** |

## ページ別詳細

### /activist-dashboard.html (1件)

#### link (1件)

| # | selector | text | 現在(w×h) | 不足(w/h) | href/for |
|---:|---|---|---:|---:|---|
| 1 | `html > body > div:nth-of-type(5) > details > p:nth-of-type(3) > a` | アクティビストの定義｜アクティビストのことがわかるブログ | 283.95×42.39 | 0/1.6099999999999994 | https://activist-blog.com/definition/ |

### /activist-screener.html (2件)

#### link (2件)

| # | selector | text | 現在(w×h) | 不足(w/h) | href/for |
|---:|---|---|---:|---:|---|
| 1 | `html > body > a.skip-link` | メインコンテンツへスキップ | 195×21 | 0/23 | #main |
| 2 | `div#tab-individual > div.s-card:nth-of-type(1) > details > p > a` | EDINET | 48.73×20 | 0/24 | https://disclosure2.edinet-fsa.go.jp |

### /news/activist-shareholder-proposals-japan.html (6件)

#### link (6件)

| # | selector | text | 現在(w×h) | 不足(w/h) | href/for |
|---:|---|---|---:|---:|---|
| 1 | `body > div.report-cover:nth-of-type(1) > div.report-cover-inner > div.report-meta:nth-of-type(5) > span:nth-of-type(3) > a:nth-of-type(1)` | 三井住友信託銀行調べ（日経） | 189×20 | 0/24 | https://www.nikkei.com/article/DGXZQOUB141S60U5A610C2000000/ |
| 2 | `body > div.report-cover:nth-of-type(1) > div.report-cover-inner > div.report-meta:nth-of-type(5) > span:nth-of-type(3) > a:nth-of-type(2)` | 大和総研 | 56×20 | 0/24 | https://www.dir.co.jp/report/consulting/governance/20251031_025387.html |
| 3 | `div#main > section.article-section:nth-of-type(4) > div.trend-grid > article.trend-card:nth-of-type(3) > p > a:nth-of-type(1)` | 114社・399議案 | 102.63×20 | 0/24 | https://www.nikkei.com/article/DGXZQOUB141S60U5A610C2000000/ |
| 4 | `div#main > section.article-section:nth-of-type(4) > div.trend-grid > article.trend-card:nth-of-type(3) > p > a:nth-of-type(2)` | EY-Parthenon | 88.03×20 | 0/24 | https://www.ey.com/ja_jp/newsroom/2025/09/ey-japan-news-release-2025-09-16 |
| 5 | `div#main > section.article-section:nth-of-type(4) > div.trend-grid > article.trend-card:nth-of-type(3) > p > a:nth-of-type(3)` | 大和総研 2025年10月 | 133.77×20 | 0/24 | https://www.dir.co.jp/report/consulting/governance/20251031_025387.html |
| 6 | `div#main > section.article-section:nth-of-type(4) > div.trend-grid > article.trend-card-cta:nth-of-type(8) > p:nth-of-type(1) > a` | 選択 | 243.83×41 | 0/3 | https://www.sentaku.co.jp/articles/view/25892 |

