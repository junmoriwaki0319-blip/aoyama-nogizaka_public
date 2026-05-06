// Compute 3-run median for each URL × category from audits/20260507-recheck/run-{1,2,3}/lh-mobile-*.json
// Outputs:
//   - median-summary.json: { "<slug>": { "url": "...", "performance": <median>, ..., "raw": [r1, r2, r3] } }
//   - summary.md (4/21 baseline + 5/5 baseline + 5/7 median + Δ)
//
// 4/21 and 5/5 baseline values are hardcoded from audits/20260421/lighthouse-mobile/summary.md
// and audits/20260505/summary.md (per Step 2 incidents.md).
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DIR = __dirname;

const PAGES = [
  ['home',                                       'https://aoyama-nogizaka.com/'],
  ['team',                                       'https://aoyama-nogizaka.com/team'],
  ['news',                                       'https://aoyama-nogizaka.com/news/'],
  ['privacy',                                    'https://aoyama-nogizaka.com/privacy'],
  ['activist-dashboard',                         'https://aoyama-nogizaka.com/activist-dashboard.html'],
  ['risk-assessment',                            'https://aoyama-nogizaka.com/risk-assessment.html'],
  ['activist-screener',                          'https://aoyama-nogizaka.com/activist-screener.html'],
  ['food-service',                               'https://aoyama-nogizaka.com/food-service.html'],
  ['saas',                                       'https://aoyama-nogizaka.com/saas.html'],
  ['ad-agency',                                  'https://aoyama-nogizaka.com/ad-agency.html'],
  ['digital-media',                              'https://aoyama-nogizaka.com/digital-media.html'],
  ['entertainment-sector-dashboard',             'https://aoyama-nogizaka.com/entertainment-sector-dashboard.html'],
  ['news-activist-shareholder-proposals-japan',  'https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html'],
];

// 4/21 baseline (mobile only): [perf, a11y, bp, seo]
const BASE_4_21 = {
  'home':                                       [56, 90, 96, 100],
  'team':                                       [87, 91, 96, 100],
  'news':                                       [82, 91, 92, 100],
  'privacy':                                    [56, 91, 96, 100],
  'activist-dashboard':                         [60, 87, 96, 100],
  'risk-assessment':                            [73, 86, 96, 100],
  'activist-screener':                          [72, 92, 96, 100],
  'food-service':                               [77, 88, 96, 100],
  'saas':                                       [78, 88, 96, 100],
  'ad-agency':                                  [83, 88, 96, 100],
  'digital-media':                              [82, 88, 96, 100],
  'entertainment-sector-dashboard':             [78, 88, 96, 100],
  'news-activist-shareholder-proposals-japan':  [54, 94, 92, 100],
};

// 5/5 baseline (mobile only)
const BASE_5_5 = {
  'home':                                       [26, 90, 96, 100],
  'team':                                       [63, 91, 96, 100],
  'news':                                       [57, 91, 92, 100],
  'privacy':                                    [55, 91, 96, 100],
  'activist-dashboard':                         [55, 87, 96, 100],
  'risk-assessment':                            [77, 86, 96, 100],
  'activist-screener':                          [46, 92, 96, 100],
  'food-service':                               [56, 88, 96, 100],
  'saas':                                       [56, 88, 96, 100],
  'ad-agency':                                  [60, 88, 96, 100],
  'digital-media':                              [56, 88, 96, 100],
  'entertainment-sector-dashboard':             [51, 88, 96, 100],
  'news-activist-shareholder-proposals-japan':  [54, 94, 92, 100],
};

const readJSON = (p) => JSON.parse(fs.readFileSync(p, 'utf-8'));

function lhScores(file) {
  const j = readJSON(file);
  const c = j.categories || {};
  const s = (k) => {
    const v = c[k]?.score;
    return v == null ? null : Math.round(v * 100);
  };
  return {
    perf: s('performance'),
    a11y: s('accessibility'),
    bp:   s('best-practices'),
    seo:  s('seo'),
  };
}

function median(xs) {
  const a = [...xs].filter(x => x != null).sort((p, q) => p - q);
  if (!a.length) return null;
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : Math.round((a[mid-1] + a[mid]) / 2);
}

function range(xs) {
  const a = [...xs].filter(x => x != null);
  if (!a.length) return null;
  return Math.max(...a) - Math.min(...a);
}

const results = {};
const rows = [];

for (const [slug, url] of PAGES) {
  const runs = [1, 2, 3].map(n => lhScores(path.join(DIR, `run-${n}`, `lh-mobile-${slug}.json`)));
  const perfs = runs.map(r => r.perf);
  const a11ys = runs.map(r => r.a11y);
  const bps   = runs.map(r => r.bp);
  const seos  = runs.map(r => r.seo);

  const med = {
    perf: median(perfs),
    a11y: median(a11ys),
    bp:   median(bps),
    seo:  median(seos),
  };

  results[slug] = {
    url,
    performance: med.perf,
    accessibility: med.a11y,
    'best-practices': med.bp,
    seo: med.seo,
    raw_perfs: perfs,
    raw_a11ys: a11ys,
    raw_bps: bps,
    raw_seos: seos,
    perf_range: range(perfs),
    a11y_range: range(a11ys),
  };

  rows.push({ slug, url, runs, med, perfs, a11ys, bps, seos });
}

fs.writeFileSync(path.join(DIR, 'median-summary.json'), JSON.stringify(results, null, 2));
console.log('WROTE median-summary.json:', Object.keys(results).length, 'pages');

// === Build summary.md ===
const avg = (xs) => xs.length ? Math.round(xs.reduce((a, b) => a + b, 0) / xs.length * 10) / 10 : null;
const fmt = (n) => n == null ? '—' : String(n);
const sgn = (d) => d == null ? '—' : (d === 0 ? '±0' : (d > 0 ? `+${d}` : String(d)));

let md = '';
md += `# Perf Recheck Summary — 5/7 中央値ベースライン (2026-05-07)\n\n`;
md += `Lighthouse 13.2.0 / mobile only / 13 URL × 3 run = 39 jobs / 中央値ベース。\n`;
md += `Generated by: \`audits/20260507-recheck/calc-median.mjs\`\n\n`;
md += `## 計測条件\n\n`;
md += `| 項目 | 値 |\n|---|---|\n`;
md += `| Lighthouse | 13.2.0（npx --yes、npm cache 経由） |\n`;
md += `| Form factor | mobile |\n`;
md += `| Chrome flags | \`--headless --no-sandbox\` |\n`;
md += `| Max wait for load | 60,000ms |\n`;
md += `| Total jobs | **39**（13 URL × 3 ラン） |\n`;
md += `| 5/7 recheck 開始時刻 | 2026-05-07 03:01 JST（深夜帯） |\n`;
md += `| 4/21 baseline 時刻帯 | 2026-04-21 10:16 JST（午前帯） |\n`;
md += `| 5/5 baseline 時刻帯 | 2026-05-05 09:51 JST 〜 5/6 16:45 JST（混在） |\n`;
md += `| 同時刻帯一致 | **不一致**（詳細は \`timing-context.md\` 参照） |\n\n`;

md += `## 中央値ベースの 4/21 vs 5/5 vs 5/7 比較（mobile-perf 中心）\n\n`;
md += `| URL | 4/21 perf | 5/5 perf | 5/7 median | 5/7−4/21 Δ | 5/7−5/5 Δ | 判定（5/7 vs 4/21） |\n`;
md += `|---|---:|---:|---:|---:|---:|---|\n`;

const judge = (d) => {
  if (d == null) return '—';
  if (d >= -5) return '計測ノイズ';
  if (d <= -10) return '真の回帰';
  return '中間';
};

let perfDeltas = [];
for (const r of rows) {
  const b421 = BASE_4_21[r.slug]?.[0];
  const b55 = BASE_5_5[r.slug]?.[0];
  const m = r.med.perf;
  const d421 = (m == null || b421 == null) ? null : m - b421;
  const d55  = (m == null || b55 == null) ? null : m - b55;
  perfDeltas.push(d421);
  md += `| [${r.slug}](${r.url}) | ${fmt(b421)} | ${fmt(b55)} | **${fmt(m)}** | ${sgn(d421)} | ${sgn(d55)} | ${judge(d421)} |\n`;
}

const avg421 = avg(rows.map(r => BASE_4_21[r.slug]?.[0]).filter(x => x != null));
const avg55  = avg(rows.map(r => BASE_5_5[r.slug]?.[0]).filter(x => x != null));
const avg57  = avg(rows.map(r => r.med.perf).filter(x => x != null));
const dAvg421 = (avg57 - avg421).toFixed(1);
const dAvg55  = (avg57 - avg55).toFixed(1);
md += `| **平均** | **${avg421}** | **${avg55}** | **${avg57}** | **${dAvg421 >= 0 ? '+' : ''}${dAvg421}** | **${dAvg55 >= 0 ? '+' : ''}${dAvg55}** | **${judge(parseFloat(dAvg421))}** |\n\n`;

md += `## 結果分岐の判定\n\n`;
md += `- **平均 −5 以内**: 真の回帰なし（5/5 baseline は計測ノイズ）\n`;
md += `- **平均 −10 以上**: 真の回帰あり、仮説 2/3 を再調査\n`;
md += `- **中間（−5 〜 −10）**: 追加ランか別環境クロスチェック\n\n`;

const overall = parseFloat(dAvg421);
let verdict;
if (overall >= -5) verdict = `**Case A: 計測ノイズ判定**（平均 ${dAvg421}）`;
else if (overall <= -10) verdict = `**Case B: 真の回帰判定**（平均 ${dAvg421}）`;
else verdict = `**Case C: 中間**（平均 ${dAvg421}）`;
md += `### 本タスクの判定\n${verdict}\n\n`;

md += `平均 5/7 中央値 = ${avg57} / 4/21 baseline = ${avg421} / 5/5 baseline = ${avg55}\n\n`;

md += `## 5/7 内 3 ラン分散（max−min）— 計測安定性\n\n`;
md += `| URL | run-1 | run-2 | run-3 | range | 分散判定 |\n`;
md += `|---|---:|---:|---:|---:|---|\n`;
const variances = [];
for (const r of rows) {
  const rg = range(r.perfs);
  const variStatus = rg <= 5 ? '安定' : (rg <= 10 ? 'やや不安定' : '不安定');
  variances.push({ slug: r.slug, range: rg });
  md += `| ${r.slug} | ${fmt(r.perfs[0])} | ${fmt(r.perfs[1])} | ${fmt(r.perfs[2])} | ${fmt(rg)} | ${variStatus} |\n`;
}

const sortedVar = [...variances].sort((a, b) => (b.range || 0) - (a.range || 0));
md += `\n### 分散の大きい URL Top 3\n\n`;
md += `| rank | URL | range |\n|---:|---|---:|\n`;
for (let i = 0; i < Math.min(3, sortedVar.length); i++) {
  md += `| ${i+1} | ${sortedVar[i].slug} | ${fmt(sortedVar[i].range)} |\n`;
}

md += `\n## 全カテゴリ中央値（perf 以外も提示）\n\n`;
md += `| URL | perf | a11y | bp | seo |\n|---|---:|---:|---:|---:|\n`;
for (const r of rows) {
  md += `| ${r.slug} | ${fmt(r.med.perf)} | ${fmt(r.med.a11y)} | ${fmt(r.med.bp)} | ${fmt(r.med.seo)} |\n`;
}

md += `\n## 補足: 同時刻帯条件の影響評価\n\n`;
md += `4/21 は午前帯、5/5 は午前〜午後混在、5/7（本 recheck）は深夜帯で計測しており、**完全な同時刻帯一致は得られていない**（\`timing-context.md\` 参照）。`;
md += `ただし本 recheck は同一セッション内で 3 ラン連続実行のため、3 ランの内部分散は時刻帯影響を受けない。\n\n`;
md += `仮に 5/7 中央値が 4/21 と乖離した場合、それが「真の回帰」か「時刻帯由来」かを切り分けるには、4/21 と同じ午前帯で追加 1 ランを実施してクロスチェックが必要。\n\n`;

md += `## 参考: 入力ファイル\n\n`;
md += `- \`audits/20260507-recheck/run-1/lh-mobile-*.json\` × 13\n`;
md += `- \`audits/20260507-recheck/run-2/lh-mobile-*.json\` × 13\n`;
md += `- \`audits/20260507-recheck/run-3/lh-mobile-*.json\` × 13\n`;
md += `- \`audits/20260507-recheck/median-summary.json\`（本スクリプト出力）\n`;
md += `- \`audits/20260507-recheck/timing-context.md\`（時刻帯コンテキスト）\n`;

fs.writeFileSync(path.join(DIR, 'summary.md'), md);
console.log('WROTE summary.md:', md.length, 'bytes');

console.log('\n=== Headline ===');
console.log(`5/7 median avg perf: ${avg57}`);
console.log(`vs 4/21 (${avg421}): ${dAvg421}`);
console.log(`vs 5/5 (${avg55}): ${dAvg55}`);
console.log(`Verdict: ${verdict.replace(/\*\*/g, '')}`);
