# 仮説 2 / 仮説 3 事実調査レポート（2026-05-07）

## エグゼクティブ・サマリ

| 仮説 | 判定 | 根拠の核 |
|---|---|---|
| **仮説 2**: `data/reports.json` 1.12 MB 膨張が真の回帰の主因 | **反証** | 真の回帰 8 ページの reports.json 接点 = **0/8** |
| **仮説 3a**: Brotli/gzip 圧縮が効いていない | **反証** | 全 13 ページで `Content-Encoding: br`、reports.json も br 圧縮済 |
| **仮説 3b**: 真の回帰群の TTFB が大きい | **反証** | 真の回帰群 平均 ~0.263s、ノイズ群 平均 ~0.233s — 有意差なし |
| **仮説 3c**: Cache MISS / STALE が多発 | **反証** | 全 13 ページ `x-vercel-cache: HIT`、age 値も連続増加で cache 安定 |

→ **仮説 2 / 3 は両方とも反証**。コード起因・データ起因・配信起因のいずれでも回帰を説明できない。

副次的に新仮説の手がかりが得られた:
- **副次観察 A**: 9/13 URL が `cleanUrls: true` 由来の 308 redirect を経由（4/21 当時から同じ、state change ではない）
- **副次観察 B**: HTML 不変・vercel.json 不変・3rd-party 不変なのに −15 が持続 → 「**変わったのはサイト側ではない外部要因**」の可能性が強まる（Vercel Edge 動作変更 / Chrome 自動更新 / Lighthouse 13.x 内部 scoring drift など）

---

## 仮説 2 詳細

### Step H2-1: コードベース全体の `/data/reports.json` 参照

```
$ git grep -n 'data/reports\.json' -- '*.html' '*.js' '*.mjs' '*.ts'
js/activist-dashboard-page.js:6:const DATA_URL = '/data/reports.json';
scripts/check_status.js:2:const data = JSON.parse(fs.readFileSync('data/reports.json', 'utf8'));
```

**該当ファイル数: 2**
- `js/activist-dashboard-page.js`: クライアント JS（fetch する）
- `scripts/check_status.js`: Node 管理スクリプト（クライアント不参照）

→ **クライアントから fetch するのは `js/activist-dashboard-page.js` の 1 ファイルのみ**。

### Step H2-2: `js/activist-dashboard-page.js` の fetch ロジック（全文 Read 済）

主要な分岐（line 12-40）:

```js
async function loadData() {
  if (dataLoaded) return;
  let loaded = false;
  const user = window.currentUser;
  if (user) {
    try {
      const idToken = await user.getIdToken();
      const resp = await fetch(PREMIUM_API + '?type=reports', { ... });
      if (resp.ok) { reportData = await resp.json(); loaded = true; }
    } catch (e) { /* フォールバックへ */ }
  }
  // フォールバック: 直接ファイル取得
  if (!loaded) {
    const cacheBuster = '?t=' + Date.now();
    const resp = await fetch(DATA_URL + cacheBuster, { cache: 'no-store' });
    if (resp.ok) { reportData = await resp.json(); }
  }
  ...
}
```

**重要な観察**:
- DOMContentLoaded で **無条件に `loadData()` を呼ぶ** (line 8-11)
- 認証なしでは fallback 経路で `/data/reports.json?t=${Date.now()}` を fetch（**cache buster + `cache: 'no-store'`** で CDN cache を毎回バイパス）
- → **Lighthouse 計測（認証なし）では毎回 14MB をオリジン fetch + パース**

**`js/activist-dashboard-page.js` を読み込む HTML**:
```
$ git grep -n 'activist-dashboard-page\.js' -- '*.html'
activist-dashboard.html:1066:<script src="/js/activist-dashboard-page.js"></script>
activist-dashboard/index.html:995:<script src="/js/activist-dashboard-page.js"></script>
```

→ **2 ファイルのみ**（どちらも activist-dashboard 系）。

### Step H2-3: 真の回帰確定 8 ページとの接点

真の回帰 8 ページ: team, news, risk-assessment, activist-screener, food-service, saas, digital-media, entertainment-sector-dashboard

**直接接点**: 全 8 ページとも `activist-dashboard-page.js` を**読み込まない**。

**間接接点**: 共通 JS（≥3 ページで読まれるもの）:

| 共通 JS | 利用ページ数 | reports.json 参照？ |
|---|---:|---|
| `/js/nav-handler.js` | 8 | なし |
| `/js/css-loader.js` | 8 | なし |
| `/js/ga-loader.js` | 7 | なし |
| `/js/auth.js` | 4 | なし |
| `chart.js / chart プラグイン × 3` (CDN) | 4 | なし |

→ **真の回帰 8 ページの reports.json 接点 = 0/8**。

### Step H2-4: reports.json サイズ・形式

```
data/reports.json: 14M, 368,874 行, CRLF, JSON
last_updated: 2026-05-06T20:48:42+09:00
total_reports: 24,940
activist_reports: 1,242
```

サイズは確かに 14MB だが、**真の回帰 8 ページのいずれもこのファイルを fetch しない**。

### 仮説 2 判定: **反証**

唯一影響を受けうる activist-dashboard.html は 5/7 中央値 56（4/21 比 −4）= 計測ノイズ判定で、**真の回帰群には属さない**。
仮に reports.json が原因なら逆の現象（activist-dashboard だけ大幅悪化）になるはずで、観測パターンと矛盾する。

---

## 仮説 3 詳細

### Step H3-1: 13 URL レスポンスヘッダ（`-L` で redirect 追跡後の最終応答）

| slug | Content-Encoding | Cache-Control | x-vercel-cache | Age |
|---|---|---|---|---:|
| home | **br** | public, max-age=0, must-revalidate | **HIT** | 109 |
| team | **br** | 〃 | **HIT** | 108 |
| news | **br** | 〃 | **HIT** | 107 |
| privacy | **br** | 〃 | **HIT** | 107 |
| activist-dashboard | **br** | 〃 | **HIT** | 6 |
| risk-assessment | **br** | 〃 | **HIT** | 0 |
| activist-screener | **br** | 〃 | **HIT** | 0 |
| food-service | **br** | 〃 | **HIT** | 0 |
| saas | **br** | 〃 | **HIT** | 11 |
| ad-agency | **br** | 〃 | **HIT** | 0 |
| digital-media | **br** | 〃 | **HIT** | 0 |
| entertainment-sector-dashboard | **br** | 〃 | **HIT** | 0 |
| news-activist-shareholder-proposals-japan | **br** | 〃 | **HIT** | 16 |

→ **全 13 ページで Brotli (br) 圧縮が効き、Vercel CDN cache HIT で配信**されている。

### Step H3-2: data/reports.json のヘッダ

```
Content-Encoding: br
Content-Type: application/json; charset=utf-8
X-Vercel-Cache: HIT
Cache-Control: public, max-age=0, must-revalidate
```

→ reports.json も Brotli 圧縮 + cache HIT。**14MB は wire 上で大幅に縮む**（Brotli で JSON は 80-90% 圧縮）。

ただし `js/activist-dashboard-page.js` が `cache: 'no-store'` で fetch するため、CDN cache が効いてもブラウザ側では毎回 origin から取得する形になる（ただし wire 上は br 圧縮）。

### Step H3-3: TTFB 計測

| 群 | 平均 TTFB | 内訳 |
|---|---:|---|
| 真の回帰 8 ページ | **~0.263s** | team 0.365 / news 0.281 / risk-assessment 0.238 / activist-screener 0.269 / food-service 0.238 / saas 0.260 / digital-media 0.220 / entertainment-sector-dashboard 0.234 |
| 計測ノイズ 5 ページ | **~0.233s** | home 0.350 / privacy 0.158 / activist-dashboard 0.192 / ad-agency 0.225 / news-activist-shareholder-proposals-japan 0.240 |

差: +30ms 程度（真の回帰群の方がわずかに大きいが、Lighthouse の perf score を 15 点削るには **数倍以上の差**が必要）。

→ **TTFB の有意差なし**。

### Step H3-4: Cache HIT/MISS 安定性

3 URL × 3 連続 HEAD で確認:
- home: HIT/HIT/HIT (age 183 → 185 → 186)
- saas: HIT/HIT/HIT (age 79 → 81 → 83)
- team: HIT/HIT/HIT (age 191 → 193 → 194)

→ **全 HIT、age が秒単位で増加する正常動作**。MISS / STALE への切替なし。CDN cache は完全に健全。

### 仮説 3 判定: **反証**

3a (圧縮)、3b (TTFB)、3c (Cache) いずれも問題なし。Vercel Edge の配信品質は 5/7 時点で**完全に正常**。

---

## 副次観察: cleanUrls による 308 redirect

`vercel.json` の `cleanUrls: true` により、`.html` 拡張子付き URL は `/path` への 308 redirect を返す:

```
$ curl -sI -L https://aoyama-nogizaka.com/saas.html
HTTP/1.1 308 Permanent Redirect
Location: /saas
HTTP/1.1 200 OK
```

13 URL のうち **9 URL が num_redirects=1**（cleanUrls 由来）:
- redirect=1 群: activist-dashboard, risk-assessment, activist-screener, food-service, saas, ad-agency, digital-media, entertainment-sector-dashboard, news-activist-shareholder-proposals-japan
- redirect=0 群: home, team, news, privacy（既に clean パスまたは `/team` 形式の rewrite が vercel.json で別定義されているため）

Lighthouse は redirect を perf 上不利に評価しうる（"Avoid multiple page redirects" audit）。ただし:
- 4/21 baseline 計測時にも `cleanUrls: true` は有効（`vercel.json` は 4/21 → main 期間で**変更なし**）
- → redirect は state change ではなく、4/21 → 5/7 の差分要因にはならない

ただし、**redirect ありのページが 5/7 で軒並み 50 台に張り付いている**のは興味深い。redirect の影響が「ベース疲弊」として常時存在し、何か別の付加要因と組み合わさったときに顕在化する可能性は残る。

---

## 推奨次アクション（仮説 2/3 反証を踏まえて）

### 仮説 4（隠れた要因）の探索が必要

仮説 2/3 がいずれも反証された以上、**コード・データ・配信のどれでもない外部要因**を疑う段階。候補:

#### 仮説 4a: Lighthouse 13.x scoring algorithm のバージョン依存
- 4/21 baseline と 5/5 baseline の Lighthouse バージョンを再確認（両方 13.2.0 か？）
- 13.x の minor / patch リリースで perf scoring の閾値が変わった可能性
- 検証: `audits/20260421/lighthouse-mobile/home.json` の `lighthouseVersion` を grep
- 修正: 同 minor で揃えるか、Lighthouse 12 系で再計測してクロスチェック

#### 仮説 4b: ローカル Chrome のバージョン更新
- 4/21 〜 5/7 の間に Windows 側で Chrome が自動更新された可能性
- 検証: `chrome.exe --version` を実行（4/21 計測時は不明、現状のみ取得可）
- 影響: Chrome の rendering / network throttling 実装変更が perf score に間接的に影響

#### 仮説 4c: Vercel Edge / Vercel 側のインフラ変更
- HTML / CSS / JS / vercel.json すべて不変だが、**Vercel 側で Edge ノードや配信戦略が変わった可能性**
- 検証: Vercel ダッシュボードでアナリティクスを確認、5/5 前後にデプロイメント設定の変更があったか
- 修正: Vercel サポートに問い合わせるか、別ホスティングで一時的に検証

#### 仮説 4d: Mobile network throttling 変更
- Lighthouse mobile は CPU 4× slowdown + Slow 4G throttling を適用
- 13.x で throttling のデフォルト値が変わった可能性
- 検証: 4/21 と 5/7 の JSON で `throttlingMethod` / `throttling` 設定を比較

### 即時推奨アクション

1. **`audits/20260421/lighthouse-mobile/home.json` の `lighthouseVersion` を grep**で 4/21 baseline の LH バージョンを確定
2. **`audits/20260507-recheck/run-1/lh-mobile-home.json` の `environment.hostUserAgent`** も確認（Chrome バージョン情報含む）
3. これらが一致していたら 4a/4d は反証 → 4b/4c へ
4. 一致していなければ 4a/4d を主仮説に格上げ

### 追加で実施可能な観点

- **`dataLoaded` の無条件 fetch を condition 付きにする**（仮説 2 の予防策）: たとえ仮説 2 が反証でも、Lighthouse 計測中に 14MB を毎回 fetch するのは回避すべき将来リスク。`activist-dashboard.html` を ad-agency / digital-media など他ページと比較したとき、唯一 perf を大きく下げる要素になり得る
- **redirect=1 のページを cleanUrl で叩くよう監査スクリプトを修正**（cosmetic）: 4/21 → 5/7 baseline 比較の精度を上げるため、`.html` 拡張子なしでも別 baseline を取り、redirect 影響を切り分け

---

## 入力データ

- ヘッダ: `audits/20260507-recheck/headers/{slug}-final.txt` × 13 + `data-reports-json.txt`
- TTFB: `audits/20260507-recheck/ttfb.txt`
- 既存 incidents: `incidents/20260505-perf-regression.md`（Step 2、仮説 1〜3 の初期分析）

## 結論

仮説 2 / 3 とも反証。次フェーズは **仮説 4（隠れた要因）の探索**で、Lighthouse バージョン / Chrome バージョン / Vercel Edge 動作の確認から開始すべき。
