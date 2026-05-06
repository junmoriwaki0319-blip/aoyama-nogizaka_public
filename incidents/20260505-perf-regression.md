# 5/5 mobile-perf 全面回帰 RCA（2026-05-07）

## サマリー

- 5/5 baseline 監査で **mobile-perf 平均が 4/21 比 −17.4（72.2 → 54.8）と全面回帰**を検出
- 12/13 ページで悪化（−1〜−30）、唯一改善は risk-assessment（+4）、唯一同点は news-activist-shareholder-proposals-japan（±0）
- a11y は 89.4 で完全に変動なし
- **本 RCA の結論: 4/21 → main の commit range には HTML / CSS / JS / 画像の変更は一切存在しない**。フロントエンド資産は完全に同一であり、コード起因の回帰は理論上発生し得ない
- 唯一の変更は `data/reports.json` の **+1.12 MB（+8.6%）膨張**だが、これを参照するページは activist-dashboard 系のみで、しかも当該ページの悪化幅は −5 と全 13 ページ中最小クラス。データ仮説では他ページの −25〜−30 を説明できない
- 最有力仮説は **Lighthouse 計測ノイズ／環境差**。修正コミット投入前に **3 回連続計測の中央値で再検証する**ことを Top 1 推奨修正方針とする

## Commit Range

| 項目 | 値 |
|---|---|
| BASE_4_21（4/21 baseline 取得時点）| `27ccbfc314a1a2d9d2202608a3f72cf80f43db77` |
| main HEAD（5/5 直前） | `ef2a686eba5a43de90af19a7827de53f03e220cb` |
| 件数 | **63 コミット**（merge 除く） |
| 期間 | 2026-04-21 ～ 2026-05-06 |

### ファイル変更（diff --stat）

```
 data/edinet-financials.json |     2 +-
 data/reports.json           | 32432 +++++++++++++++++++++++++++++++++++-------
 2 files changed, 27031 insertions(+), 5403 deletions(-)
```

### 変更ファイル（拡張子別カウント）

| 拡張子 | 件数 |
|---|---:|
| `.json` | **2** |
| `.html` | 0 |
| `.css` | 0 |
| `.js` | 0 |
| `.png/jpg/svg/webp/avif` | 0 |
| **フロントエンド総計** | **0** |

### コミットの内訳

63 コミット全てが GitHub Actions による自動 EDINET データ更新:

```
ef2a686 data: update EDINET reports 2026-05-06T11:48:42Z   ← main HEAD
06b716e data: update EDINET financials cache 2026-05-06T11:03:26Z
... (60 commits omitted)
04a30b0 data: update EDINET financials cache 2026-04-21T10:18:47Z
fef1621 data: update EDINET reports 2026-04-21T08:58:09Z   ← BASE_4_21+1
```

すべて `data/edinet-financials.json` または `data/reports.json` の更新のみ。手動コミットは存在しない。

## 共通アセット変更

| 観点 | 結果 |
|---|---|
| 3rd-party タグ追加削除 | **変動 0**（HTML 不変のため `<script src="http...">` / `<link rel=stylesheet href="http...">` は完全同一） |
| 共通テンプレート | `assets/`, `includes/`, `components/`, `partials/` ディレクトリは存在しない（HTML 直書き構成）。HTML 不変のため共通テンプレ変更も 0 |
| グローバル CSS / JS | `css/`, `js/` ディレクトリ配下に変更ファイル 0 |

## 上位 3 ページ個別 diff

| ページ | 5/5 perf | 4/21 perf | Δ | git diff 結果 |
|---|---:|---:|---:|---|
| `index.html` (home) | 26 | 56 | −30 | **空（変更なし）** |
| `entertainment-sector-dashboard.html` | 51 | 78 | −27 | **空（変更なし）** |
| `activist-screener.html` | 46 | 72 | −26 | **空（変更なし）** |

`git diff --stat ${BASE_4_21}..main -- <file>` で全 3 ページとも出力空。1 行も変更されていない。

## `data/reports.json` 詳細調査

| 観点 | 値 |
|---|---|
| サイズ 4/21 (`27ccbfc`) | 12,971,075 bytes（12.97 MB） |
| サイズ 5/4 (`0b83ef4`, 5/5 計測直前) | 14,004,708 bytes（14.00 MB） |
| サイズ 5/6 (`ef2a686`, main HEAD) | 14,087,979 bytes（14.09 MB） |
| 増加量 | **+1,116,904 bytes（+1.12 MB / +8.6%）** |
| client-side fetch している箇所 | `js/activist-dashboard-page.js:6` の `DATA_URL` のみ |
| 読み込まれる HTML | `activist-dashboard.html` および `activist-dashboard/index.html` のみ |
| 読み込みパターン | Firebase 認証ありなら `/api/premium-reports?type=reports` 経由、認証なしなら `/data/reports.json` を直接 fetch（fallback） |

LH 計測は認証なし → fallback で 14 MB を直接 fetch する。これが activist-dashboard の Δperf に効く可能性はあるが、観測値は **−5（4/21 60 → 5/5 55）と全ページ中最小級**。逆に `reports.json` を一切参照しない home が **−30 で最大悪化**しており、データ膨張仮説とは観測パターンが矛盾する。

## 仮説 3 つ

### 仮説 1（最有力）: Lighthouse mobile-perf の計測ノイズ／環境差

- **一行サマリ**: フロントエンド変更ゼロかつ reports.json 参照外のページが軒並み −25〜−30 悪化していることから、本回帰は LH の run-to-run variance（特に Windows + headless Chrome 環境）由来である可能性が極めて高い
- **裏付け**:
  - `git diff --stat 27ccbfc..main` の出力（HTML/CSS/JS 変更 0 ファイル）
  - reports.json を参照しない home (`index.html`) が −30、entertainment-sector-dashboard.html が −27、digital-media.html が −26
  - reports.json を参照する activist-dashboard.html は −5 のみ（参照しないページより悪化が小さい）
  - 5/5 baseline 計測ログ `audits/20260505/_run.log` に記録された全 LH 実行で `chrome-launcher` の `Launcher.kill` 例外あり（既知ノイズ）
  - LH 公式ドキュメント上、mobile perf score は run-to-run で ±10 以上のブレが報告されている
- **修正方針**: コード変更は不要。3 回連続計測の中央値で再ベースラインを取得し、真の回帰の有無を確定する。具体策は「推奨修正方針 Top 1」参照
- **推定インパクト**: 真の perf スコアは 4/21 と同水準（平均 70 前後）に戻る可能性が高い。**+15〜+20 点回復見込み**（再計測のみで）

### 仮説 2: `data/reports.json` 1.12 MB 膨張による activist-dashboard 系の局所的劣化

- **一行サマリ**: reports.json が 12.97 MB → 14.09 MB に膨張し、activist-dashboard.html / activist-dashboard/index.html での fetch + JSON.parse コストが Δ−5 を生んでいる
- **裏付け**:
  - `js/activist-dashboard-page.js:6` `const DATA_URL = '/data/reports.json';`
  - `js/activist-dashboard-page.js:8-11` `DOMContentLoaded` で同期的に `loadData() → renderAll()` を実行
  - `js/activist-dashboard-page.js:21-32` 認証なしフローでは fallback として 14 MB ファイルを直接 fetch
  - activist-dashboard.html: 4/21 60 → 5/5 55（−5）。他ページの −25〜−30 と比べ局所的影響に留まる
- **修正方針**: 
  - 短期: `<link rel="preload" as="fetch" href="/data/reports.json">` の検討（既にあれば不要）または fetch を `defer` 化
  - 中期: pagination 化／必要分のみ取得する API 経路化（`api/premium-reports.js` 経由は実装済、認証必須化で fallback 廃止が筋）
  - 長期: gzip / Brotli 圧縮確認、Vercel 側の Cache-Control 設定確認
- **推定インパクト**: activist-dashboard 系のみ +5〜+8 点。**他ページには影響なし**

### 仮説 3: Vercel Edge / 外部 CDN の一時的劣化（5/5 計測時刻特有）

- **一行サマリ**: 5/5 計測は 09:51〜16:46 の長時間に渡り、その時間帯に Vercel Edge / Google Fonts / Firebase / GA4 / jsdelivr 等の外部経路が遅延した可能性
- **裏付け**:
  - HTML 不変なので外部依存は 4/21 と完全同一
  - 5/5 計測ログ `audits/20260505/_run.log`、`_dry-run.log` が示す通り、本日（5/5）の長時間連続計測
  - 外部 CDN の品質低下は数時間スパンでの再現性が低い（=4/21 と異なる時間帯の影響を受けている可能性）
- **修正方針**: 別日・別時間帯で再計測。サーバ側修正なし
- **推定インパクト**: 仮説 1 と同様、再計測で全面改善する可能性。+10〜+15 点

## 推奨修正方針 Top 1

**仮説 1（計測ノイズ／環境差）を最有力として採用。修正コミット投入前に 3 回連続計測の中央値で再ベースラインを取得する。**

### 実行ブランチ案

`investigate/perf-regression-20260505-recheck`（fix ではなく re-measure 用）

### 具体策

1. 4/21 baseline 取得時と同じ手順で **同時刻帯**（午前中／午後など条件を揃える）に Lighthouse mobile を 3 回連続実行する
2. 13 URL × 3 ラン = 39 ジョブ。`audits/20260507-recheck/run-1/`, `run-2/`, `run-3/` 配下に分けて保存
3. 各 URL × 各カテゴリで中央値（median）を取り、`audits/20260507-recheck/summary.md` を生成
4. 中央値ベースで 4/21 vs 5/7 を再比較
5. 結果分岐:
   - 平均 −5 以内に収まれば **真の回帰なし**（5/5 baseline は計測ノイズ）と判定し、PR #6 baseline には但し書きを追加して closing
   - 平均 −10 以上で残れば **真の回帰あり**として仮説 2（reports.json）または別の隠れた要因を再調査（fix ブランチへ）

### 実装コスト見積

**S（Small）** — 既存 `audits/20260421/run-lighthouse.sh` を流用、対象 URL 13、ラン数 3。1 回あたり 13 URL × 約 25 秒 = 約 5.5 分。3 回計 ~17 分 + summary 生成 ~3 分。**1 セッション（30 分以内）で完結**

### 推定 perf 回復

中央値 baseline で **+15〜+20 点回復見込み**（4/21 平均 72.2 への回帰）。回復しなければ仮説 2/3 を順に検証

### ロールバック容易性

**極めて高い**。コード変更ゼロのため、どの段階で停止してもロールバック不要。出力 JSON / md は `audits/20260507-recheck/` 配下に閉じる

### 副次的に推奨される作業（別 PR）

- **`api/premium-reports.js` 経由を必須化**して `/data/reports.json` 直接 fetch fallback を廃止する（仮説 2 の予防）
- `data/reports.json` のサイズ増加トレンドを Vercel CDN ヘッダで圧縮配信されているか確認（Brotli / gzip 効果測定）

## 補足: なぜ「修正なし」を Top 1 にしたか

通常 perf 回帰の RCA は「何を直すか」を提示するが、本ケースは:

1. 4/21 → main のフロントエンド変更が **物理的にゼロ**（git の事実）
2. 唯一のデータ変更（reports.json）と最も影響を受けるべきページ（activist-dashboard）の悪化幅（−5）が、関係ないページ（home: −30）より遥かに小さい

という客観事実から、コード起因の回帰として説明できる経路が存在しない。**修正前に「本当に回帰しているのか」を確定させる**のが最優先で、これが Top 1 となる。

仮説 2（reports.json 膨張）については、仮に真の回帰がなくても **将来の予防策**として有効なので別 PR として推奨する。
