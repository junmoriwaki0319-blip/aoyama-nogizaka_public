import fs from 'node:fs';

const DIR = './audits/20260420';
const files = fs.readdirSync(DIR);

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

// ヒューリスティック優先度: footer/copy/address => P2, main content => P0
function priority(selector, snippet='') {
  const s = (selector||'').toLowerCase();
  const sn = (snippet||'').toLowerCase();
  if (/footer|\.footer|address|copy/.test(s) || /<footer|footer-/.test(sn)) return 'P2';
  if (/nav|hamburger|nav-logo|nav-link/.test(s)) return 'P0';
  if (/hero|cta|register|login|btn-primary|ranking|tab/.test(s)) return 'P0';
  return 'P1';
}

function suggestFix(audit, selector) {
  switch (audit) {
    case 'tap-targets': return 'min-height/min-width を 48px 以上に拡大、または近接要素との間隔を 10px 以上確保';
    case 'color-contrast': return 'WCAG 2 AA (4.5:1 通常 / 3:1 大) を満たす配色に変更。rgba 透過の親要素との合成を確認';
    case 'viewport': return '<meta name="viewport" content="width=device-width, initial-scale=1"> を追加';
    case 'meta-viewport': return 'viewport meta の user-scalable=no / maximum-scale を削除（ズーム許可）';
    case 'font-size': return '小サイズテキスト (<12px) を拡大。モバイルは最低 14-16px を基準';
    case 'select-name': return 'select 要素に aria-label または関連 <label for> を付与';
    default: return '—';
  }
}

// slug list from filesystem
const slugs = Array.from(new Set(
  files
    .map(f => {
      let m = f.match(/^lh-mobile-only-(.+)\.json$/);
      if (m) return m[1];
      m = f.match(/^axe-(.+)\.json$/);
      if (m) return m[1];
      return null;
    })
    .filter(Boolean)
)).sort();

const rows = [];
// per-page counts
const pageCount = {};
const auditCount = {};

for (const slug of slugs) {
  const url = slugToUrl[slug] || slug;
  // Lighthouse only-audits
  const lhFile = `${DIR}/lh-mobile-only-${slug}.json`;
  if (fs.existsSync(lhFile)) {
    try {
      const j = JSON.parse(fs.readFileSync(lhFile));
      const aud = j.audits || {};
      for (const [key, a] of Object.entries(aud)) {
        if (a.score === 1 || a.score === null) continue; // passed or NA
        // detail items
        const items = (a.details && (a.details.items || [])) || [];
        if (items.length === 0) {
          // audit failed but no per-element items (e.g. meta-viewport)
          rows.push({ url, slug, source: 'lighthouse', audit: key, selector: '(page-level)', snippet: a.title || '', fix: suggestFix(key), prio: 'P0' });
          pageCount[slug] = (pageCount[slug]||0) + 1;
          auditCount[key] = (auditCount[key]||0) + 1;
          continue;
        }
        for (const it of items) {
          const sel = (it.node && (it.node.selector || it.node.path)) || it.selector || '(unknown)';
          const snip = (it.node && it.node.snippet) || '';
          rows.push({ url, slug, source: 'lighthouse', audit: key, selector: sel, snippet: snip.slice(0,160), fix: suggestFix(key, sel), prio: priority(sel, snip) });
          pageCount[slug] = (pageCount[slug]||0) + 1;
          auditCount[key] = (auditCount[key]||0) + 1;
        }
      }
    } catch (e) {
      console.error('lh parse fail', slug, e.message);
    }
  }
  // axe-core
  const axeFile = `${DIR}/axe-${slug}.json`;
  if (fs.existsSync(axeFile)) {
    try {
      const j = JSON.parse(fs.readFileSync(axeFile));
      // axe-core CLI wraps in array
      const report = Array.isArray(j) ? j[0] : j;
      const violations = (report && report.violations) || [];
      for (const v of violations) {
        for (const n of (v.nodes||[])) {
          const sel = (n.target||[]).join(' > ') || '(unknown)';
          const snip = (n.html || '').slice(0,160);
          rows.push({ url, slug, source: 'axe', audit: v.id, selector: sel, snippet: snip, fix: suggestFix(v.id, sel), prio: priority(sel, snip) });
          pageCount[slug] = (pageCount[slug]||0) + 1;
          auditCount[v.id] = (auditCount[v.id]||0) + 1;
        }
      }
    } catch (e) {
      console.error('axe parse fail', slug, e.message);
    }
  }
}

// sort: page -> prio -> audit
const prioRank = { P0: 0, P1: 1, P2: 2 };
rows.sort((a,b) => a.slug.localeCompare(b.slug) || (prioRank[a.prio]-prioRank[b.prio]) || a.audit.localeCompare(b.audit));

let md = `# Violations by Page — 2026-04-20

Source: Lighthouse (\`--only-audits=tap-targets,color-contrast,viewport,meta-viewport,font-size\`) + @axe-core/cli

- 13 URL × 2 ツール = 26 レポートを機械集計
- 優先度ヒューリスティック: nav/hero/cta/btn/ranking/tab 等の主要導線 = **P0**, footer/copy/address = **P2**, それ以外 = **P1**

## ページ別違反件数

| slug | violations |
|---|---:|
`;
for (const slug of slugs) {
  md += `| [${slug}](${slugToUrl[slug]||slug}) | ${pageCount[slug]||0} |\n`;
}
const total = Object.values(pageCount).reduce((a,b)=>a+b,0);
md += `| **TOTAL** | **${total}** |\n`;

md += `\n## audit 別違反件数\n\n| audit | count |\n|---|---:|\n`;
for (const [k,v] of Object.entries(auditCount).sort((a,b)=>b[1]-a[1])) {
  md += `| ${k} | ${v} |\n`;
}

md += `\n## 違反セレクタ一覧 (URL × audit × selector)\n\n`;
md += `| URL | source | audit | prio | selector | snippet | suggested-fix |\n`;
md += `|---|---|---|---|---|---|---|\n`;
for (const r of rows) {
  const sn = (r.snippet||'').replace(/\|/g,'\\|').replace(/\n/g,' ');
  const sel = (r.selector||'').replace(/\|/g,'\\|').replace(/\n/g,' ');
  md += `| [${r.slug}](${r.url}) | ${r.source} | ${r.audit} | ${r.prio} | \`${sel}\` | \`${sn}\` | ${r.fix} |\n`;
}

fs.writeFileSync(`${DIR}/violations-by-page.md`, md);
console.log('wrote', `${DIR}/violations-by-page.md`);
console.log('total violations', total);
console.log('per page', pageCount);
console.log('per audit', auditCount);
