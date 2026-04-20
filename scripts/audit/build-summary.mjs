import fs from 'node:fs';

const DIR = './audits/20260420';
const files = fs.readdirSync(DIR).filter(f => /^lh-mobile-[^.]+\.json$/.test(f)).sort();

// slug -> URL
const slugToUrl = {
  'home': 'https://aoyama-nogizaka.com/',
  'team': 'https://aoyama-nogizaka.com/team',
  'news': 'https://aoyama-nogizaka.com/news/',
  'privacy': 'https://aoyama-nogizaka.com/privacy',
  'activist-dashboard': 'https://aoyama-nogizaka.com/activist-dashboard.html',
  'risk-assessment': 'https://aoyama-nogizaka.com/risk-assessment.html',
  'activist-screener': 'https://aoyama-nogizaka.com/activist-screener.html',
  'food-service': 'https://aoyama-nogizaka.com/food-service.html',
  'saas': 'https://aoyama-nogizaka.com/saas.html',
  'ad-agency': 'https://aoyama-nogizaka.com/ad-agency.html',
  'digital-media': 'https://aoyama-nogizaka.com/digital-media.html',
  'entertainment-sector-dashboard': 'https://aoyama-nogizaka.com/entertainment-sector-dashboard.html',
  'news-activist-shareholder-proposals-japan': 'https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html',
};

// Lighthouse mobile thresholds (2024+)
// score: <50 poor, 50-89 NI, >=90 good
// LCP (mobile, 4G): good <=2.5s, NI <=4.0s, poor >4.0s
// CLS: good <=0.1, NI <=0.25, poor >0.25
// TBT (mobile): good <=200, NI <=600, poor >600
// INP: good <=200, NI <=500, poor >500

function markScore(s) {
  if (s === '-') return '-';
  if (s < 90) return `**<span style="color:#d00">${s}</span>**`;
  return `${s}`;
}

function numLCP(s) {
  if (!s || s === '-') return null;
  const m = /([\d.]+)\s*s/.exec(s);
  return m ? parseFloat(m[1]) : null;
}
function markLCP(s) {
  const n = numLCP(s);
  if (n == null) return s || '-';
  if (n > 4.0) return `**<span style="color:#d00">${s} (poor)</span>**`;
  if (n > 2.5) return `**<span style="color:#d00">${s} (NI)</span>**`;
  return s;
}

function numCLS(s) {
  if (!s || s === '-') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function markCLS(s) {
  const n = numCLS(s);
  if (n == null) return s || '-';
  if (n > 0.25) return `**<span style="color:#d00">${s} (poor)</span>**`;
  if (n > 0.1) return `**<span style="color:#d00">${s} (NI)</span>**`;
  return s;
}

function numTBT(s) {
  if (!s || s === '-') return null;
  const m = /([\d,]+)\s*ms/.exec(s);
  return m ? parseInt(m[1].replace(/,/g,''),10) : null;
}
function markTBT(s) {
  const n = numTBT(s);
  if (n == null) return s || '-';
  if (n > 600) return `**<span style="color:#d00">${s} (poor)</span>**`;
  if (n > 200) return `**<span style="color:#d00">${s} (NI)</span>**`;
  return s;
}

const rows = [];
for (const f of files) {
  const slug = f.replace(/^lh-mobile-/,'').replace(/\.json$/,'');
  try {
    const j = JSON.parse(fs.readFileSync(`${DIR}/${f}`));
    const c = j.categories || {};
    const a = j.audits || {};
    const sc = k => c[k] && c[k].score != null ? Math.round(c[k].score*100) : '-';
    const v = k => a[k] && a[k].displayValue ? a[k].displayValue : '-';
    rows.push({
      slug,
      url: slugToUrl[slug] || slug,
      perf: sc('performance'),
      a11y: sc('accessibility'),
      best: sc('best-practices'),
      seo: sc('seo'),
      lcp: v('largest-contentful-paint'),
      cls: v('cumulative-layout-shift'),
      tbt: v('total-blocking-time'),
      inp: v('interaction-to-next-paint'),
    });
  } catch (e) {
    rows.push({ slug, url: slugToUrl[slug]||slug, perf:'-',a11y:'-',best:'-',seo:'-',lcp:'-',cls:'-',tbt:'-',inp:'-' });
  }
}

let md = `# Lighthouse Mobile Baseline — 2026-04-20

Form factor: mobile / Chrome headless / Lighthouse CLI (npx @latest)

## Scores (perf / a11y / best / seo) + Core Web Vitals

スコア <90 と Core Web Vitals の NI / poor は赤太字マーク。

| URL | perf | a11y | best | seo | LCP | CLS | TBT | INP |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
`;
for (const r of rows) {
  md += `| [${r.slug}](${r.url}) | ${markScore(r.perf)} | ${markScore(r.a11y)} | ${markScore(r.best)} | ${markScore(r.seo)} | ${markLCP(r.lcp)} | ${markCLS(r.cls)} | ${markTBT(r.tbt)} | ${r.inp} |\n`;
}

// Aggregate
const perfs = rows.map(r => typeof r.perf==='number'?r.perf:null).filter(x=>x!=null);
const a11ys = rows.map(r => typeof r.a11y==='number'?r.a11y:null).filter(x=>x!=null);
const best = rows.map(r => typeof r.best==='number'?r.best:null).filter(x=>x!=null);
const seos = rows.map(r => typeof r.seo==='number'?r.seo:null).filter(x=>x!=null);
const avg = arr => arr.length ? Math.round(arr.reduce((a,b)=>a+b,0)/arr.length*10)/10 : '-';

md += `\n## Aggregates

- **Performance**: avg ${avg(perfs)}, min ${Math.min(...perfs)}, <90 count: ${perfs.filter(x=>x<90).length}/${perfs.length}
- **Accessibility**: avg ${avg(a11ys)}, min ${Math.min(...a11ys)}, <90 count: ${a11ys.filter(x=>x<90).length}/${a11ys.length}
- **Best Practices**: avg ${avg(best)}, min ${Math.min(...best)}, <90 count: ${best.filter(x=>x<90).length}/${best.length}
- **SEO**: avg ${avg(seos)}, min ${Math.min(...seos)}, <90 count: ${seos.filter(x=>x<90).length}/${seos.length}

## Notes

- Lighthouse CLI の Chrome temp-dir cleanup で Windows の EPERM が発生したが、レポート本体 (JSON) は全 13 URL 書き出し済み。スコアは上表の通り取得できている。
- INP は Lighthouse CLI では計測されない (\`interaction-to-next-paint\` audit は field data ベース)。列には audit の \`displayValue\` をそのまま出しているため \`-\` になっているものは未計測。
- Thresholds:
  - Score: <50 poor / 50-89 NI / ≥90 good
  - LCP: ≤2.5s good / ≤4.0s NI / >4.0s poor
  - CLS: ≤0.1 good / ≤0.25 NI / >0.25 poor
  - TBT: ≤200ms good / ≤600ms NI / >600ms poor

Raw JSON: \`./audits/20260420/lh-mobile-*.json\`
Log: \`./audits/20260420/lh-baseline.log\`
`;

fs.writeFileSync(`${DIR}/summary-mobile.md`, md);
console.log('wrote', `${DIR}/summary-mobile.md`);
