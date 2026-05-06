# LH 13.1.0 強制再計測 比較（2026-05-07）— 仮説 4a 確定検証

## エグゼクティブ・サマリ

| 検証項目 | 結果 |
|---|---|
| 5/7 LH 13.1.0 強制再計測 平均 mobile-perf | **56.8** |
| 4/21 baseline (LH 13.1.0) 平均 | 72.2 |
| 5/5 baseline (LH 13.2.0) 平均 | 54.8 |
| 5/7 中央値 (LH 13.2.0) | 57.2 |
| **5/7 LH13.1.0 vs 4/21** | **−15.4** |
| **5/7 LH13.1.0 vs 5/7 LH13.2.0 中央値** | **−0.4** |

→ **仮説 4a (LH バージョン drift) 反証**。同じ LH 13.1.0 を使っても 5/7 計測値は **4/21 baseline と +5 以内に収束しなかった**（−15.4 の乖離が残る）。LH 13.1.0 と 13.2.0 で計測した結果はほぼ同一（差 −0.4）であり、LH バージョンが perf スコアに与える影響は事実上ゼロ。

→ **真の原因は仮説 4c（Vercel Edge / インフラ動作変更）**または**未特定の隠れた要因**に絞り込まれた。

## 計測条件

| 項目 | 値 |
|---|---|
| 取得日時 | 2026-05-07T04:31:42 〜 04:35:04 JST（深夜帯、3 分 22 秒） |
| Lighthouse バージョン | **13.1.0**（`npx --yes lighthouse@13.1.0` で強制指定）|
| Chrome | HeadlessChrome/147.0.0.0（4/21 と完全一致） |
| Throttling | rttMs=150 / cpuSlowdownMultiplier=4 / throttlingMethod=simulate（4/21 と一致） |
| benchmarkIndex 範囲 | min=4125.5 / max=4489.0 / **avg=4336.8** / median=4359.0 |
| 取得回数 | **1 ラン × 13 URL** |
| 5/7 中央値（LH13.2.0）の benchmarkIndex（参考） | home=4088.5 |
| 4/21 baseline の benchmarkIndex（参考） | home=4103.5 |

→ 環境条件は **4/21 とほぼ完全に一致**。CPU 競合や Chrome 版数の影響を排除した上での測定値である。

## 4 baseline 比較表

| URL | 4/21 LH13.1.0 | 5/5 LH13.2.0 | 5/7 LH13.2.0 中央値 | **5/7 LH13.1.0 (本タスク)** | 5/7 LH13.1.0 − 4/21 Δ | 判定 |
|---|---:|---:|---:|---:|---:|---|
| [home](https://aoyama-nogizaka.com/) | 56 | 26 | 55 | **66** | **+10** | 改善 |
| [team](https://aoyama-nogizaka.com/team) | 87 | 63 | 56 | **56** | **−31** | 真の回帰 |
| [news](https://aoyama-nogizaka.com/news/) | 82 | 57 | 56 | **56** | **−26** | 真の回帰 |
| [privacy](https://aoyama-nogizaka.com/privacy) | 56 | 55 | 57 | **57** | +1 | 一致 |
| [activist-dashboard](https://aoyama-nogizaka.com/activist-dashboard.html) | 60 | 55 | 56 | **56** | −4 | 一致 |
| [risk-assessment](https://aoyama-nogizaka.com/risk-assessment.html) | 73 | 77 | 54 | **37** | **−36** | 真の回帰（極大、要再計測） |
| [activist-screener](https://aoyama-nogizaka.com/activist-screener.html) | 72 | 46 | 46 | **56** | **−16** | 真の回帰 |
| [food-service](https://aoyama-nogizaka.com/food-service.html) | 77 | 56 | 56 | **55** | **−22** | 真の回帰 |
| [saas](https://aoyama-nogizaka.com/saas.html) | 78 | 56 | 57 | **55** | **−23** | 真の回帰 |
| [ad-agency](https://aoyama-nogizaka.com/ad-agency.html) | 83 | 60 | 83 | **56** | **−27** | 真の回帰 |
| [digital-media](https://aoyama-nogizaka.com/digital-media.html) | 82 | 56 | 56 | **56** | **−26** | 真の回帰 |
| [entertainment-sector-dashboard](https://aoyama-nogizaka.com/entertainment-sector-dashboard.html) | 78 | 51 | 57 | **55** | **−23** | 真の回帰 |
| [news-activist-shareholder-proposals-japan](https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html) | 54 | 54 | 54 | **78** | **+24** | 改善 |
| **平均** | **72.2** | **54.8** | **57.2** | **56.8** | **−15.4** | **真の回帰** |

## 結果分岐の判定: **Case 1b（4a 反証）**

判定基準:
- 平均が 4/21 (72.2) ± 5 → **不一致** (−15.4)
- 平均が 5/7 LH13.2.0 中央値 (57.2) ± 5 → **一致** (差 −0.4)

→ **仮説 4a 反証確定**。**LH 13.2.0 へのバージョンアップは perf スコアに事実上影響を与えていない**。原因は別にある。

## ページ別の傾向観察

### 真の回帰群 vs ノイズ群の振る舞い

Step 3a-recheck で定義した「真の回帰群 8」「ノイズ群 5」を、5/7 LH13.1.0 値で再評価:

| 群 | ページ数 | 4/21 LH13.1.0 と一致したページ（Δ ≥ −5） |
|---|---:|---|
| 真の回帰群 (8) | 8 | **0 / 8**（全ページで Δ < −5、−16 〜 −36 の範囲） |
| ノイズ群 (5) | 5 | **3 / 5**（home +10、privacy +1、activist-dashboard −4） |

→ **真の回帰群は LH 13.1.0 で再計測しても全ページが回帰したまま**。これは LH バージョン由来ではなく、サイト or インフラに何かが起きていることを強く示唆。

### 個別の異常値

- **risk-assessment: 5/7 LH13.1.0 = 37**（4/21=73 から −36、5/7 中央値 54 からも −17）
  - 1 ラン だけだと特に振れが大きい URL（前回 5/7 中央値計測時も run-1=51, run-2=54, run-3=55 でばらつきあり）
  - 信頼区間を取るには 3 ラン中央値が必要
- **home: 5/7 LH13.1.0 = 66 (+10)** および **news-activist-shareholder-proposals-japan: 78 (+24)**
  - これらは「改善」しているが、計測ノイズによる上振れの可能性大
  - 5/7 中央値ではそれぞれ 55 と 54 で別の結果

つまり 1 ラン だけでは個別ページの判定はノイジー。**平均 56.8 という総合値**が判定に最も信頼できる数値。

## benchmarkIndex の評価（CPU 競合の有無）

| baseline | benchmarkIndex |
|---|---:|
| 4/21 home | 4103.5 |
| 5/5 home | **2064.0**（異常） |
| 5/7 home（中央値時 run-1） | 4088.5 |
| **5/7 LH13.1.0 home（本タスク）** | **4489.0** |

→ 本タスクの benchmarkIndex は **4/21 よりやや高い**（CPU により余裕がある状態）。それでも perf 平均は 56.8 で 4/21 (72.2) に届かない。
→ **計測時の CPU 状態が 4/21 と同等以上にもかかわらずスコアが届かない**ということは、**ホスト側の問題ではなく、サイト/インフラ側に何かが起きている**。

## LH 13.2.0 changelog の主要変更（参考）

公式リリースノート (https://github.com/GoogleChrome/lighthouse/releases/tag/v13.2.0、2026-05-01 公開) より抜粋:

### 新規 audit
- `webmcp-form-coverage` / `webmcp-registered-tools` / `webmcp-schema-validity`（WebMCP 関連、新カテゴリ）
- `agentic` カテゴリ追加（AI エージェント向け）

### Core 変更
- `implement UKM Invalidate fallback for LCP`（LCP 計算の fallback ロジック追加 — perf に影響しうるが、4/21 と 5/5 の間の主因なら 5/7 LH13.1.0 で改善するはずで、しなかったため棄却可）
- accessibility audits の default config への追加・分離（agentic 向け）

### Deps
- `trace_engine` 0.0.64 へ upgrade（trace 解析の microtuning）
- `web-features` 3.24.0 へ upgrade

### perf scoring の重み変更
- **明示的な重み変更や閾値変更は changelog に記載なし**
- 13.1.0 → 13.2.0 は minor リリースで、Performance category の構造的変更は無し

→ **changelog の事実が「LH バージョン変更は perf スコアに影響しない」という本タスクの観測結果と完全に整合**。

## 残仮説と次フェーズ

### 確定事項
- ✅ 4a (LH version drift): **反証**（LH 13.1.0 でも 5/7 値は 4/21 から −15.4）
- ✅ 4b (Chrome UA): 反証（全 baseline で Chrome 147.0.0.0）
- ✅ 4d (throttling): 反証（設定完全一致）
- ✅ 仮説 2 (reports.json): 反証（接点ページなし、Step 3c で確定）
- ✅ 仮説 3 (CDN/圧縮/TTFB): 反証（Brotli ○、HIT ○、TTFB 差なし、Step 3c で確定）

### 残る仮説
- 🔴 **仮説 4c (Vercel Edge / インフラ動作変更)**: 検証必要、本タスクスコープ外
- 🔴 **仮説 5（新規）: Chrome 内部 rendering / trace 解釈の subtle drift**: Chrome は 147.0.0.0 で同じだが、minor patch 内での trace 出力変化や、`trace_engine` 0.0.64 の解釈変化が複合的に効いている可能性
- 🔴 **計測ノイズ（既知）**: 1 ラン だと特にバラつくが、5/5・5/7（3 ラン）・5/7 LH13.1.0（1 ラン）の **平均がいずれも 56-57 に収束**しているため、ノイズ単独では説明不可

### 推奨次アクション

**仮説 4c (Vercel Edge / インフラ動作変更) の検証**が最も筋が通る。サイト側のコード/設定が完全に不変で、計測環境も同等なのに −15 が出る以上、**「実際に届いている content / network 経路」が何か変わっている** と考えるのが自然。

具体的には:

1. **Vercel ダッシュボードでの確認**（森脇さん側操作必須）
   - 4/21 〜 5/5 の deployment 履歴
   - region / edge node 構成変更の有無
   - `vercel.json` の effective config が変わったか（ファイル不変でも Vercel 側の解釈が更新されることがある）
   - Build cache の無効化 / 再生成タイミング

2. **Vercel API 経由の deployment 詳細取得**（プログラマブル）
   ```
   curl -H "Authorization: Bearer $VERCEL_TOKEN" \
     "https://api.vercel.com/v6/deployments?projectId=...&since=2026-04-20&until=2026-05-07"
   ```

3. **別ホスティング/別 CDN での A/B 検証**（最終手段）
   - サイトを Cloudflare Pages や Netlify に並行デプロイし、同 LH13.1.0 で計測
   - Vercel と差が出れば 4c 確定、差が出なければサイト自体の問題

### 追加観点: 4/21 baseline 自体の異常値の可能性

逆説的だが、4/21 baseline (72.2) が **異常に高かった**可能性も検討すべき:
- 4/21 計測時に何か特殊な条件（CDN cache がフルに warm、ネットワーク経路が好条件など）が揃った可能性
- 5/5、5/7 中央値、5/7 LH13.1.0 の 3 つの計測 (LH バージョン異なる、時刻帯異なる、ラン数異なる) すべてが **56-57 帯に収束** していることは、これが **新しい "本当のスコア"** であることを示唆

この場合、回帰として捉えるのではなく **「4/21 が外れ値、現状が定常状態」** と再解釈することも可能。判断は森脇さん側に委ねる。

## 入力ファイル

- `audits/20260507-recheck-lh131/run-1/lh-mobile-*.json` × 13（本タスク）
- `audits/20260507-recheck-lh131/scores.json`（perf/a11y/bp/seo の集計）
- `audits/20260507-recheck-lh131/lh-13.2.0-changelog.md`（公式 release notes）
- `audits/20260507-recheck-lh131/_run.log`（計測ログ）

## 結論

**仮説 4a 反証確定**。Lighthouse バージョン (13.1.0 vs 13.2.0) は perf スコアに事実上影響を与えていない（差 −0.4）。

**真の原因は引き続き未特定**だが、選択肢は:
1. **仮説 4c (Vercel Edge / インフラ)**: 残仮説の最有力。Vercel ダッシュボード調査必要
2. **4/21 baseline 外れ値説**: 5/5・5/7・5/7 LH13.1.0 が 56-57 に収束している事実を踏まえ、4/21 が異常値だった可能性を再検討

どちらに振るかは森脇さんの判断。**本タスクではここで停止**。
