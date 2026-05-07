## 但し書き — Step 3 系検証結果（2026-05-07 追記）

PR #6 では 4/21 → 5/5 で mobile-perf 平均 -17.4 の全面回帰を baseline として記録した。
原因究明のため Step 3a-recheck → 3c → 3d → 3e → 3f の検証を実施し、以下が確認された。

### 反証された仮説

| 仮説 | 検証 Step | 結果 |
|---|---|---|
| 1. 計測ノイズ | Step 3a-recheck | 反証（5/7 中央値 57.2 で持続） |
| 2. reports.json 直接 fetch | Step 3c | 反証（真の回帰 8 ページと接点 0/8） |
| 3. CDN/圧縮/TTFB | Step 3c | 反証（br 圧縮 ✓、HIT ✓、TTFB 差なし） |
| 4a. LH version drift | Step 3d/3e | 反証（5/7 LH 13.1.0 でも 56.8） |
| 4b. Chrome UserAgent | Step 3d | 反証（全 baseline で 147.0.0.0 一致） |
| 4d. throttling preset | Step 3d | 反証（設定完全一致） |
| **4c. Vercel Edge / インフラ動作変更** | **Step 3f** | **反証**（BASE_421=27ccbfc を 5/7 環境で計測 = 72.0、4/21 と同値） |
| **4/21 baseline 外れ値説** | **Step 3f** | **反証**（同上、4/21 = 72.2 は真値） |

### Step 3f の決定的検証

4/21 baseline 計測時の app commit (`BASE_421=27ccbfc`) を 5/7 環境（現 Vercel infra）で再計測:

- BASE_421 on 5/7 infra: **avg perf 72.0**
- 4/21 baseline: 72.2
- diff: **-0.2pp**（ノイズ域）

→ **インフラは変化していない**。4/21 baseline は外れ値ではなく真値。

### 真の原因: app コード変更

5/5 = 54.8 は同じ infra での 17pp 劣化。原因は **app コード変更にある**ことが確定。

- commit 範囲: `27ccbfc..main`（4/21 baseline 取得時の本番 commit から 5/5 baseline まで）
- diff: HTML 21 ファイル / +27 / -37（css/js/images は touch なし）
- 小さな HTML 変更が大幅劣化を引き起こした可能性高

### 次フェーズ Step 4: git bisect

`27ccbfc..main` を二分探索し劣化原因 commit を特定。
結果は別 PR で incidents/ に追記予定。

### Step 3e エビデンス保全に関する補足

Step 3e で計測直後の Lighthouse JSON 13 ファイルが 2 回連続で消失する事故が発生。
Windows Defender 隔離履歴ゼロ → ASR / CFA / Search Indexer / Chrome cleanup 派生プロセスが原因と推定。
2026-05-07 に Defender 除外設定（4 worktree）+ 計測直後 git add の二重対策で解決済。
Step 3f では再発せず、13 JSON すべて保存に成功。

### 裏付けファイル

- `audits/20260507-step3f/lh-base421-on-current-infra/summary.md`
- `audits/20260507-step3f/github-deployments.json`（253 件、4/15 以降）
- `audits/20260507-step3f/vercel-deployments.json`
- `audits/20260507-step3f/vercel-incidents-since-0415.json`（17 件、edge serving 影響なし）
- `audits/20260507-step3f/env-snapshot.md`（Node v24.14.0 / LH 13.2.0 / Vercel CLI 50.32.5）
- `audits/20260507-step3f/preview-url-base421.txt`
- `audits/20260507-step3f/base421.env`
- `audits/20260507-step3f/step3f-section-draft.md`
- `audits/20260507-step3f/lh-base421-on-current-infra/lh-mobile-*.json`（13 ファイル）

### 関連ブランチ

- `investigate/perf-regression-20260505-recheck-h4c` (Step 3f の主作業ブランチ)
- `investigate/perf-regression-20260505-recheck-h4c-base421` (BASE_421 revert ブランチ)

### この baseline の利用方針

PR #6 の数値（5/5 = 54.8）は **5/7 環境で再現性が確認された "現状値"** として正しく利用可能。
**「4/21 比 -17.4 の悪化」表現は確定的に正しい**（infra 不変が確定したため）。
原因 commit は Step 4 で特定後、別 PR で対処する。
