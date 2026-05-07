import fs from 'node:fs';
import path from 'node:path';

const dir = 'audits/20260507-step3f/lh-base421-on-current-infra';
const files = fs.readdirSync(dir).filter(f => f.startsWith('lh-mobile-') && f.endsWith('.json'));

const rows = [];
for (const f of files) {
  try {
    const data = JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8'));
    const perf = data.categories?.performance?.score;
    rows.push({
      page: f.replace('lh-mobile-', '').replace('.json', ''),
      perf: perf == null ? null : Math.round(perf * 100),
      lcp: data.audits?.['largest-contentful-paint']?.numericValue ?? null,
      cls: data.audits?.['cumulative-layout-shift']?.numericValue ?? null,
      tbt: data.audits?.['total-blocking-time']?.numericValue ?? null,
    });
  } catch (e) {
    rows.push({ page: f, perf: null, error: String(e).slice(0, 80) });
  }
}

const valid = rows.filter(r => typeof r.perf === 'number');
const avg = valid.length ? (valid.reduce((s, r) => s + r.perf, 0) / valid.length) : 0;
const min = valid.length ? Math.min(...valid.map(r => r.perf)) : 0;
const max = valid.length ? Math.max(...valid.map(r => r.perf)) : 0;

let md = '# BASE_421 on 5/7 infra — Perf Summary\n\n';
md += `Generated: ${new Date().toISOString()}\n\n`;
md += '| Page | Perf | LCP (ms) | CLS | TBT (ms) |\n|---|---:|---:|---:|---:|\n';
for (const r of rows) {
  md += `| ${r.page} | ${r.perf ?? 'N/A'} | ${r.lcp != null ? Math.round(r.lcp) : 'N/A'} | ${r.cls != null ? r.cls.toFixed(3) : 'N/A'} | ${r.tbt != null ? Math.round(r.tbt) : 'N/A'} |\n`;
}
md += `\n**Aggregate (Performance):** avg=${avg.toFixed(1)} min=${min} max=${max} valid=${valid.length}/${rows.length}\n`;

fs.writeFileSync(path.join(dir, 'summary.md'), md);
console.log(JSON.stringify({ avg: +avg.toFixed(1), min, max, valid: valid.length, total: rows.length }));
