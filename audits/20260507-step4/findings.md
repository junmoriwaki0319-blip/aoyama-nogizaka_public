# Step 4 — bisect 27ccbfc..main: 結論「bisect 不実施」

実行日時: 2026-05-07 (UTC 07:06–07:50)
ブランチ: investigate/perf-regression-20260505-recheck-h4c-bisect
レンジ: 27ccbfc..origin/main (89 commits, うち HTML 触る commit = 58)
戦略選定: bisect (commit count > 8)

## TL;DR
- 計測ヘルパー (measure_commit) を組む前に、両端 anchor を 3 ラン × saas で再計測したところ、**Step3F が BASE_421 saas=86 と報告した値が再現せず、median=56 だった**。
- 同時刻に main HEAD preview saas を 3 ラン計測 → median=54。**BASE_421 と main は同じ ~55 帯**。
- 同じ preview URL の同じページで perf が 52→83 (Δ=31pp) 振れる。`--throttling-method=simulate` の 1 ラン不安定性 (LH 公式既知) がそのまま出ている。
- → **27ccbfc..main の commit 範囲には bisect で見つけ出せる「劣化境界」commit が存在しない**。Step3F の「avg 17pp 劣化」は cross-session の試行ばらつきだった可能性が高い (Step3F は BASE_421 を 1 ラン計測、PR6 5/5 main は production の別時刻計測 — apples-to-oranges 比較)。

## 計測結果テーブル

### saas (1 page, simulate-throttling, mobile)

| Anchor | URL kind | Run 1 | Run 2 | Run 3 | Median | Note |
|---|---|---:|---:|---:|---:|---|
| BASE_421 (`27ccbfc`) | Vercel preview (Step3F の URL を再利用) | 55 | 57 | 56 | **56** | Step3F 報告 86 とは別物 |
| main HEAD (`97c26a7`)¹ | Vercel preview (本ブランチの auto-deploy) | 52 | 56 | 54 | **54** | – |
| production | aoyama-nogizaka.com/saas | 80 | – | – | 80² | 別 sequence で計測した 1 ラン |

¹ 97c26a7 は origin/main の e20ed4c に audits commit を 1 つ載せたもの (HTML/CSS/JS は touch していないので main と挙動同等)。
² 直前に他ページを連続計測した直後の "warm" な計測。

### Sanity check: 1ラン値が示す preview≈production

| Page | BASE_421 preview (Step3F 1run) | main HEAD preview (今 1run) | production (今 1run) |
|---|---:|---:|---:|
| saas | 86 | 83 | 80 |
| food-service | 87 | 88 | 83 |
| entertainment-sector-dashboard | 89 | 93 | 86 |
| digital-media | 86 | 86 | 84 |
| activist-dashboard | 84 | 76 | 79 |
| home (root) | 55 | 56 (3-run median) | 56 (3-run median) |

→ どのページでも BASE_421 ≈ main ≈ production (差 ≤ 8pp、概ねノイズ域)。

## 何が起きていたのか

Step3F の本文（audits/20260507-step3f/step3f-section-draft.md）は、

- BASE_421 on 5/7 infra: avg perf = 72.0 (n=13, 1ラン平均)
- 5/5 baseline (PR #6 系): avg perf = 54.8 (n=13, 別時刻の production)

これを「-17.2pp の commit 起因劣化」と結論付けたが、

1. BASE_421 saas=86 は **同じ URL で今 56**。1 ラン計測の 30pp 級ノイズで「たまたま高く出た」だけ。
2. PR6 5/5 baseline は production を別時刻に計測した値。Step3F BASE_421 は preview の別時刻計測。比較対象が「同じインフラ・同じ時刻」を満たしていない。
3. Step3C 20260507-recheck の 3-run median (production avg=57.2) は安定していたが、その 3 ラン全部が「同じ低い帯」にいた可能性が高い (= measurement window の状態に依存)。

つまり Step3F の結論「commit 起因」は前提が崩れた。

## なぜ bisect を始めなかったか (判断ログ)

bisect は「2 endpoint が明確に分離している」前提が必要。本タスクの想定は:
- BASE_421 saas ≥ 70 (good 帯)
- main HEAD saas < 65 (bad 帯)

実測:
- BASE_421 saas median = 56 (good 帯に届かない)
- main HEAD saas median = 54 (bad 帯にいるが BASE_421 と区別不能)

threshold 65 で good/bad 判定すると、両 anchor とも bad。bisect は first bad commit を返せない。
→ 計測関数を 1 commit でも回せば結論が変わるわけではない (anchor 不成立)。bisect 開始前に停止するのが合理的。

## 推奨次アクション

### A. 計測方法の根本見直し（最優先）
- LH `--throttling-method=simulate` の 1 ラン値は比較に使えない。**3-run 以上 + median + 同一セッション内で連続計測** を最低条件にする。
- できれば `--throttling-method=devtools` で実スロットリングに切り替え (より重いが variance 小さい)。
- 同一 LH バージョン・同一 Chrome version・同一 Node version で揃える。

### B. perf 改善の優先度
- home / team / news / privacy / activist-screener 等が依然 ~55 帯で低い (BASE_421 でも main でも)。
- 「劣化を戻す」のではなく「最初から低い」前提で **改善タスク** として扱うのが筋。
- LCP が 11-19s と異常に長い: hero 画像 / Chart.js bundle / フォント初期描画ブロックいずれかが疑い濃。Step3 の network/coverage analysis が直接的に効く。

### C. 監視運用
- production 監視の場合は **WebPageTest / PageSpeed Insights (Field Data, CrUX)** を併用。Lab data 1 ラン値で警報設計しない。
- CI に LH を入れるなら `--quiet --output=json --max-wait-for-load=45000 --output-path=...` を 3-5 回回して `lighthouse-batch` 系で median をとる構成にする。

### D. 本タスクのスコープ
- 本ブランチは fact-gathering only。修正コミットは作らない。bisect log もないため `bisect-log.txt` は作らない (最初の measure_commit 前に停止したため bisect セッションを start すらしていない)。
- `analysis.txt` 相当のものは本 findings.md にまとめた。

## アーティファクト

```
audits/20260507-step4/
├── _run.log                                — タイムスタンプ付き実行ログ
├── base421.env                             — BASE_421=27ccbfc
├── commit-range.txt                        — 89 commits の oneline list
├── commits-html-only.txt                   — そのうち HTML 触ってる 58 commit
├── strategy.txt                            — 当初戦略 (bisect)
├── perf-by-commit.txt                      — 計測値サマリ (実質 anchor のみ)
├── findings.md                             — 本ファイル
└── lh-anchors/
    ├── main-97c26a7/                       — main HEAD preview, 6 ページ × 1 run
    ├── production-now/                     — aoyama-nogizaka.com 5 ページ × 1 run
    ├── production-home-3run/               — production / 3 run
    ├── main-home-3run/                     — main HEAD preview / 3 run
    ├── main-saas-3run/                     — main HEAD preview saas / 3 run
    └── base421-recheck/                    — Step3F の preview URL を saas で 3 run
```

## AV 干渉 / 所要時間メモ

- AV 干渉再発なし。各 LH ラン直後に Defender quarantine 等で .json が消える事象は観測されず (Step3 で入れた除外設定が効いている)。
- 1 LH ラン所要時間 (saas page): ~12 秒。ただし `npx --yes` のキャッシュ初回で 30 秒程度かかる。
- chrome-launcher のクリーンアップで EPERM が末尾に出るが、JSON 生成は完了しており perf 値の取得には影響しない。
