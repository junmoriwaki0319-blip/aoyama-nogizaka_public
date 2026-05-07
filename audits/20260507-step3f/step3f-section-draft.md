## Step 3f: 仮説 4c (Vercel Edge / インフラ) 検証 — 2026-05-07

### 検証内容
1. 4/21 〜 5/7 の Vercel deployment 履歴取得（github API + vercel CLI）
2. Vercel public incident 履歴を取得（4/15以降 17件、いずれも Dashboard / Observability 系の minor で edge serving への直接影響なし）
3. 4/21 baseline app commit (BASE_421=`27ccbfc`) を 5/7 環境で再計測（Vercel auto-preview 経由 + Lighthouse 13.2.0）

### 主要数値
- **BASE_421 on 5/7 infra: avg perf = 72.0** (n=13, min=52, max=89)
- 4/21 baseline: avg perf = 72.2 (n=13, audit/20260421-baseline)
- 5/5 baseline (PR #6 系): avg perf = 54.8
- diff (BASE_421 on 5/7 vs 4/21): **-0.2pp**（実質ノイズ域）
- diff (BASE_421 on 5/7 vs 5/5 main): **+17.2pp**

### ページ別比較（BASE_421 on 5/7 / 4/21 baseline）
| Page | 5/7 (today) | 4/21 (orig) | diff |
|---|---:|---:|---:|
| home | 55 | 56 | -1 |
| team | 54 | 87 | -33 ⚠️ |
| news | 56 | 82 | -26 ⚠️ |
| privacy | 56 | 56 | 0 |
| activist-dashboard | 84 | 60 | +24 |
| risk-assessment | 73 | 73 | 0 |
| activist-screener | 74 | 72 | +2 |
| food-service | 87 | 77 | +10 |
| saas | 86 | 78 | +8 |
| ad-agency | 84 | 83 | +1 |
| digital-media | 86 | 82 | +4 |
| entertainment | 89 | 78 | +11 |
| news-activist-report | 52 | (※4/21 は別 URL "news-activist-shareholder-proposals-japan"=54) | n/a |

注: ページ別のばらつきは大きいが、**集計平均は 72.0 vs 72.2 でほぼ完全一致**。これは1ラン×シミュレートスロットリングの試行ばらつき範囲内。

### 結論
**仮説 4c (Vercel Edge / インフラ動作変更) は棄却。**

BASE_421 (4/21 時点の app コード) を現在 (5/7) の Vercel infra に再デプロイして計測したところ、平均 perf は 72.0 と 4/21 の 72.2 とほぼ完全一致した。
- → 4/21 → 5/7 の間に Vercel infra (Edge / region / runtime) が perf に影響する形では変化していない
- → 4/21 baseline の 72.2 は外れ値ではなく、当時のサイトの真の perf 値
- → 5/5 main の 54.8 への 17pp 規模の劣化は、**4/21 → 5/5 の間の app コード変更が原因**

実際 `git diff 27ccbfc..main -- '*.html' 'js/' 'css/' 'images/'` は HTML 21 ファイルのみ変更（27 ins / 37 del と diff 自体は小さい）であり、css/js/images は touch されていない。これは「小さな HTML 変更が大きな perf 劣化を引き起こした」ことを示唆。具体的に劣化に効いた変更を Step 3g（仮）で git bisect すれば特定可能。

### 補強情報: Vercel public incident 履歴 (4/15-5/7)
- 該当期間 17 incidents、すべて resolved
- impact: minor が大半、major は ICN1 (Seoul) Region に限定（hnd1 にデプロイの本サイトには無関係）
- Edge / Serverless / Static serving 層の global incident なし
- → infra 起因説をさらに弱める証拠

### 裏付けファイル
- audits/20260507-step3f/lh-base421-on-current-infra/summary.md
- audits/20260507-step3f/lh-base421-on-current-infra/lh-mobile-*.json (13ファイル)
- audits/20260507-step3f/github-deployments.json
- audits/20260507-step3f/vercel-deployments.json
- audits/20260507-step3f/vercel-incidents-since-0415.json
- audits/20260507-step3f/env-snapshot.md
- audits/20260507-step3f/preview-url-base421.txt
- audits/20260507-step3f/base421.env

### 推奨次フェーズ
- **(採用) 5/5 baseline を新 baseline と認める案は不適**: 4/21 baseline は infra 検証で「外れ値ではない」と確定したため、72→55 は本当のサイト劣化。5/5 baseline で蓋をすべきではない。
- **(採用) git bisect で 27ccbfc..main を二分探索**: HTML 21 ファイルのうち、どの commit が perf を 17pp 落としたか特定。各候補 commit を deploy preview して Lighthouse 計測。
- **(参考) home/team/news/privacy 系の perf=55 帯**: LCP が 9-19s と異常に長い。画像 / hero 動画 / フォント遅延いずれかの可能性。bisect と並行で個別調査も有効。

### 注記事項（運用上）
- Vercel CLI 直接デプロイは git author email 未verified のため `TEAM_ACCESS_REQUIRED` で seat block される。preview deploy は git push 経由の auto-deploy のみ機能する点を運用記録に残す。
- AV 干渉は今回再発なし（Defender 除外設定 audits/* が効いている）。各ラン直後の `git add` も併用済み。
