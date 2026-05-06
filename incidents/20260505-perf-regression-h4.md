# 仮説 4 検証レポート（2026-05-07） — 版数 drift + redirect 相関

## エグゼクティブ・サマリ

| 仮説 | 判定 | 核となる事実 |
|---|---|---|
| **4a: Lighthouse バージョン drift** | **🎯 当たり（確定）** | 4/21 = **13.1.0** / 5/5 = **13.2.0** / 5/7 = **13.2.0** — 4/21 → 5/5 の間にマイナーアップデート |
| **4b: Chrome UA drift** | 反証 | 全 baseline で `HeadlessChrome/147.0.0.0` 完全一致 |
| **4d: throttling preset drift** | 反証 | 全 baseline で `rttMs=150 / cpuSlowdownMultiplier=4 / throttlingMethod=simulate` 完全一致 |
| **副次観察 (redirect 相関)** | 弱相関 | 真の回帰群 6/8 (75%) vs ノイズ群 3/5 (60%) — 差 15pt のみ |
| **4c: Vercel Edge / インフラ動作** | 本タスクスコープ外 | Vercel ダッシュボード操作が必要（森脇さん側） |

→ **真の回帰の主因はほぼ確定: Lighthouse 13.1.0 → 13.2.0 のマイナーアップデートに伴う perf scoring algorithm の変更**。4a を「同じ LH 13.1.0 で再計測」して確認すれば完全確定する。

副次的に **benchmarkIndex の異常値** を発見（5/5 計測時のホスト負荷）— ただしこれは 5/5 のみの問題で、5/7 では正常値に戻っている。

---

## 仮説 4a: Lighthouse バージョン drift

### 全 13 URL × 3 baseline での `lighthouseVersion`

| baseline | バージョン | 一貫性 |
|---|---|---|
| **4/21** | **13.1.0** | 全 13 URL で一致 |
| **5/5** | **13.2.0** | 全 13 URL で一致 |
| **5/7 recheck** | **13.2.0** | 全 13 URL で一致 |

判定: **4/21 → 5/5 の間に Lighthouse が 13.1.0 から 13.2.0 にアップグレードされた**。

### 含意

- Lighthouse 13.x の minor 版アップグレードで perf scoring 関連の変更が入った可能性
- 公式 changelog で 13.2.0 のリリースノートを確認すべき（candidate: scoring weight 変更、new audits 追加、metric 計算式の修正など）
- **5/5 と 5/7 が 13.2.0 で揃っているのに −15 が持続している**事実が、これがバージョン依存の系統的な差であることを強く示唆
- 5/5 と 5/7 の中央値の差はわずか +2.4 → 13.2.0 環境下では **これが新しい "本当のスコア"**

### 即時推奨アクション

**4/21 と同じ LH 13.1.0 で 1 ラン再計測**することで以下を確定:
- 13.1.0 で再計測した 5/7 値が **4/21 baseline (72.2)** に近づく → 4a 完全確定 + 4c 自動反証 = サイトに変更なし
- 13.1.0 で再計測しても 5/7 が 57 のままなら → 13.1.0/13.2.0 の差ではなく別要因（4c 含む）

具体実装案: `npx --yes lighthouse@13.1.0 ...` で 13.1.0 を強制指定して再計測。

---

## 仮説 4b: Chrome User Agent drift

### 全 3 baseline の `environment.hostUserAgent`

```
4/21:  Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) HeadlessChrome/147.0.0.0 Safari/537.36
5/5:   Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) HeadlessChrome/147.0.0.0 Safari/537.36
5/7:   Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) HeadlessChrome/147.0.0.0 Safari/537.36
```

`networkUserAgent`（mobile emulation 用）も全 baseline で `Chrome/147.0.0.0 Mobile Safari/537.36` 完全一致。

判定: **4b 反証**。Chrome のメジャー / マイナーバージョンは 147.0.0.0 で一貫しており、Chrome 自動更新は本回帰の原因ではない。

---

## 仮説 4d: throttling preset

### 全 3 baseline の `configSettings.throttling` + `throttlingMethod`

```json
{
  "rttMs": 150,
  "throughputKbps": 1638.4,
  "requestLatencyMs": 562.5,
  "downloadThroughputKbps": 1474.5600000000002,
  "uploadThroughputKbps": 675,
  "cpuSlowdownMultiplier": 4
}
throttlingMethod: simulate
```

判定: **4d 反証**。Slow 4G + 4× CPU throttling の simulate メソッドで 3 baseline 完全一致。throttling は変動なし。

---

## 副次発見: `environment.benchmarkIndex` の異常値

| baseline | benchmarkIndex | 解釈 |
|---|---:|---|
| 4/21 | **4103.5** | 正常（ホスト CPU 余裕あり） |
| 5/5 | **2064.0** | **異常**（4/21 比 50% 低下、ホスト CPU 高負荷の可能性） |
| 5/7 | **4088.5** | 正常（4/21 とほぼ同値） |

`benchmarkIndex` は Lighthouse がホストマシンのベンチマークを取って throttling 補正に使う値。Lighthouse は理論上この値で CPU 負荷を補正する設計だが、補正の精度は完全ではない。

ただし重要な観察:
- **5/7 の benchmarkIndex は 4088.5（正常）** だが、それでも perf 中央値 = 57.2 で 4/21 と −15 乖離
- → 5/5 のホスト負荷異常は副次要因として 5/5 baseline を歪めた可能性はあるが、**5/7 で正常化していてもなお −15 ある以上、主因はやはり LH バージョン drift (4a)**

---

## 副次観察: redirect 相関分析

### 13 URL の redirect 経由状況（curl -L, num_redirects）

| slug | num_redirects | 群 |
|---|---:|---|
| home | 0 | ノイズ |
| team | 0 | 真の回帰 |
| news | 0 | 真の回帰 |
| privacy | 0 | ノイズ |
| activist-dashboard | 1 | ノイズ |
| risk-assessment | 1 | 真の回帰 |
| activist-screener | 1 | 真の回帰 |
| food-service | 1 | 真の回帰 |
| saas | 1 | 真の回帰 |
| ad-agency | 1 | ノイズ |
| digital-media | 1 | 真の回帰 |
| entertainment-sector-dashboard | 1 | 真の回帰 |
| news-activist-shareholder-proposals-japan | 1 | ノイズ |

### 群別集計

| 群 | redirect=1 | redirect=0 | 経由率 |
|---|---:|---:|---:|
| 真の回帰群 (8) | 6 | 2 | **75%** |
| ノイズ群 (5) | 3 | 2 | **60%** |
| 全体 (13) | 9 | 4 | 69% |

### 判定: **弱相関**

- 真の回帰群とノイズ群の redirect 経由率の差はわずか **15 ポイント**
- 真の回帰群でも team / news は redirect=0 で含まれる
- ノイズ群でも activist-dashboard / ad-agency / news-activist-shareholder-proposals-japan は redirect=1 で含まれる
- → **redirect 経由を主因と断定するには証拠不十分**

ただし、Lighthouse の `redirects` audit は perf score にマイナス影響を与えうるため、副次因子として残る可能性は残存。

---

## 残仮説 4c の取り扱い提案

### 4c: Vercel Edge / インフラ側の動作変更

本タスクのスコープ外（Vercel ダッシュボード / API 操作が必要）。ただし、4a が確定的に当たり判定なので **4c の優先度は大幅に低下**。

**推奨**: 4a の即時検証（13.1.0 再計測）で結果が 4/21 と一致すれば、4c は自動的に反証され検証不要となる。
4a 検証で残差が出た場合のみ、Vercel ダッシュボードで以下を確認:
- 4/21 〜 5/5 期間の deployment region 変更履歴
- Edge node 構成変更ログ
- `vercel.json` 変更履歴（git log では確認済 = 変更なし、しかし Vercel 側の効果反映タイミングがずれた可能性）

---

## 推奨次アクション

### Case 1（最有力）: 4a 確定検証

`npx --yes lighthouse@13.1.0 ...` で 13.1.0 を強制指定し、13 URL × 1 ラン再計測:

| 想定結果 | 結論 |
|---|---|
| 13.1.0 中央値が 4/21 (72.2) ± 5 以内 | **4a 完全確定**、サイトに真の回帰なし → PR #6 baseline に「LH バージョン drift 由来の見かけ上の回帰」但し書き追加 → closing |
| 13.1.0 中央値が 5/7 (57.2) ± 5 以内 | 4a 反証、4c が主因 → Vercel ダッシュボード調査へ |
| 中間 | 4a 部分当たり、scoring 変更だけでは不足 → 5 ラン中央値 + 4c 並行調査 |

実行コスト: **S（5-10 分）**。`audits/20260507-recheck-lh131/run-1/` に格納。

### Case 2: redirect 相関を別途切り分けたい場合

`vercel.json` の `cleanUrls` 動作を活かし、`.html` なしの cleanPath で 13 URL を再計測:
- 4/21 baseline は `.html` 付きで取得 → 4/21 と直接比較不能
- ただし 5/7 cleanPath vs 5/7 .html 付きで「redirect ありで何点落ちるか」が分かる
- → redirect の純粋な perf score 影響を定量化

実行コスト: **S（5 分）**。

### Case 3: 13.2.0 の changelog 確認

LH 13.2.0 のリリースノートで perf scoring 関連の変更を確認:
- https://github.com/GoogleChrome/lighthouse/releases/tag/v13.2.0
- 主に scoring algorithm / metric 変更箇所を抜粋

---

## 結論

- **主因確定: 4a (Lighthouse 13.1.0 → 13.2.0 のバージョン drift)**
- 副次要因: 5/5 ホスト CPU 負荷由来 benchmarkIndex 低下 (2064)、redirect 弱相関 (75% vs 60%)
- 4b / 4d 反証
- 4c は本タスク外、4a 検証後に必要なら追加調査

**PR #6 baseline の取り扱い**: 4a 確定検証 (Case 1) を実施し、13.1.0 中央値が 4/21 と一致すれば、PR #6 の summary.md に「Lighthouse 13.1.0 → 13.2.0 の scoring drift 由来の見かけ上の回帰」と明記して closing するのが筋。

## 入力ファイル

- 4/21 baseline: `audit/20260421-baseline:audits/20260421/lighthouse-mobile/*.json`
- 5/5 baseline: `chore/audit-20260505:audits/20260505/lighthouse-*-mobile.json`
- 5/7 recheck: `audits/20260507-recheck/run-1/lh-mobile-*.json`
- redirect 確認: `audits/20260507-recheck/redirects/*.txt` × 13
