# Lighthouse Mobile Baseline — 2026-04-20

Form factor: mobile / Chrome headless / Lighthouse CLI (npx @latest)

## Scores (perf / a11y / best / seo) + Core Web Vitals

スコア <90 と Core Web Vitals の NI / poor は赤太字マーク。

| URL | perf | a11y | best | seo | LCP | CLS | TBT | INP |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| [activist-dashboard](https://aoyama-nogizaka.com/activist-dashboard.html) | **<span style="color:#d00">52</span>** | **<span style="color:#d00">87</span>** | 96 | 100 | **<span style="color:#d00">3.2 s (NI)</span>** | **<span style="color:#d00">0.216 (NI)</span>** | **<span style="color:#d00">1,220 ms (poor)</span>** | - |
| [activist-screener](https://aoyama-nogizaka.com/activist-screener.html) | **<span style="color:#d00">62</span>** | 92 | 96 | 100 | **<span style="color:#d00">2.7 s (NI)</span>** | **<span style="color:#d00">0.38 (poor)</span>** | **<span style="color:#d00">460 ms (NI)</span>** | - |
| [ad-agency](https://aoyama-nogizaka.com/ad-agency.html) | **<span style="color:#d00">79</span>** | **<span style="color:#d00">88</span>** | 96 | 100 | **<span style="color:#d00">3.0 s (NI)</span>** | **<span style="color:#d00">0.12 (NI)</span>** | **<span style="color:#d00">270 ms (NI)</span>** | - |
| [digital-media](https://aoyama-nogizaka.com/digital-media.html) | **<span style="color:#d00">71</span>** | **<span style="color:#d00">88</span>** | 96 | 100 | **<span style="color:#d00">3.3 s (NI)</span>** | 0.093 | **<span style="color:#d00">430 ms (NI)</span>** | - |
| [entertainment-sector-dashboard](https://aoyama-nogizaka.com/entertainment-sector-dashboard.html) | **<span style="color:#d00">66</span>** | **<span style="color:#d00">88</span>** | 96 | 100 | **<span style="color:#d00">3.1 s (NI)</span>** | **<span style="color:#d00">0.15 (NI)</span>** | **<span style="color:#d00">640 ms (poor)</span>** | - |
| [food-service](https://aoyama-nogizaka.com/food-service.html) | **<span style="color:#d00">76</span>** | **<span style="color:#d00">88</span>** | 96 | 100 | **<span style="color:#d00">3.1 s (NI)</span>** | 0.094 | **<span style="color:#d00">390 ms (NI)</span>** | - |
| [home](https://aoyama-nogizaka.com/) | **<span style="color:#d00">52</span>** | 90 | 96 | 100 | **<span style="color:#d00">11.1 s (poor)</span>** | **<span style="color:#d00">0.26 (poor)</span>** | 150 ms | - |
| [news-activist-shareholder-proposals-japan](https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html) | **<span style="color:#d00">47</span>** | 94 | 92 | 100 | **<span style="color:#d00">4.3 s (poor)</span>** | 0.001 | **<span style="color:#d00">3,470 ms (poor)</span>** | - |
| [news](https://aoyama-nogizaka.com/news/) | **<span style="color:#d00">82</span>** | 91 | 92 | 100 | **<span style="color:#d00">2.7 s (NI)</span>** | **<span style="color:#d00">0.135 (NI)</span>** | **<span style="color:#d00">270 ms (NI)</span>** | - |
| [privacy](https://aoyama-nogizaka.com/privacy) | **<span style="color:#d00">86</span>** | 91 | 96 | 100 | 2.5 s | 0.066 | **<span style="color:#d00">420 ms (NI)</span>** | - |
| [risk-assessment](https://aoyama-nogizaka.com/risk-assessment.html) | **<span style="color:#d00">59</span>** | **<span style="color:#d00">86</span>** | 96 | 100 | **<span style="color:#d00">2.6 s (NI)</span>** | **<span style="color:#d00">0.451 (poor)</span>** | **<span style="color:#d00">470 ms (NI)</span>** | - |
| [saas](https://aoyama-nogizaka.com/saas.html) | **<span style="color:#d00">80</span>** | **<span style="color:#d00">88</span>** | 96 | 100 | **<span style="color:#d00">3.0 s (NI)</span>** | **<span style="color:#d00">0.119 (NI)</span>** | **<span style="color:#d00">240 ms (NI)</span>** | - |
| [team](https://aoyama-nogizaka.com/team) | **<span style="color:#d00">84</span>** | 91 | 96 | 100 | **<span style="color:#d00">2.9 s (NI)</span>** | 0.074 | **<span style="color:#d00">350 ms (NI)</span>** | - |

## Aggregates

- **Performance**: avg 68.9, min 47, <90 count: 13/13
- **Accessibility**: avg 89.4, min 86, <90 count: 7/13
- **Best Practices**: avg 95.4, min 92, <90 count: 0/13
- **SEO**: avg 100, min 100, <90 count: 0/13

## Notes

- Lighthouse CLI の Chrome temp-dir cleanup で Windows の EPERM が発生したが、レポート本体 (JSON) は全 13 URL 書き出し済み。スコアは上表の通り取得できている。
- INP は Lighthouse CLI では計測されない (`interaction-to-next-paint` audit は field data ベース)。列には audit の `displayValue` をそのまま出しているため `-` になっているものは未計測。
- Thresholds:
  - Score: <50 poor / 50-89 NI / ≥90 good
  - LCP: ≤2.5s good / ≤4.0s NI / >4.0s poor
  - CLS: ≤0.1 good / ≤0.25 NI / >0.25 poor
  - TBT: ≤200ms good / ≤600ms NI / >600ms poor

Raw JSON: `./audits/20260420/lh-mobile-*.json`
Log: `./audits/20260420/lh-baseline.log`
