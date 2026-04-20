#!/usr/bin/env node
// T6: OGP/canonical/sitemap coverage checker
// Fetches 13 URLs in parallel and extracts meta tags; also fetches /sitemap.xml
// and diffs against the target URL list.

import fs from 'node:fs';

const OUT_DIR = './audits/20260420';
fs.mkdirSync(OUT_DIR, { recursive: true });

const BASE = 'https://aoyama-nogizaka.com';
const URLS = [
  `${BASE}/`,
  `${BASE}/team`,
  `${BASE}/news/`,
  `${BASE}/privacy`,
  `${BASE}/activist-dashboard.html`,
  `${BASE}/risk-assessment.html`,
  `${BASE}/activist-screener.html`,
  `${BASE}/food-service.html`,
  `${BASE}/saas.html`,
  `${BASE}/ad-agency.html`,
  `${BASE}/digital-media.html`,
  `${BASE}/entertainment-sector-dashboard.html`,
  `${BASE}/news/activist-shareholder-proposals-japan.html`,
];

const META_KEYS = [
  { key: 'og:title', pattern: /<meta[^>]+property=["']og:title["'][^>]*content=["']([^"']*)["']/i },
  { key: 'og:description', pattern: /<meta[^>]+property=["']og:description["'][^>]*content=["']([^"']*)["']/i },
  { key: 'og:image', pattern: /<meta[^>]+property=["']og:image["'][^>]*content=["']([^"']*)["']/i },
  { key: 'og:url', pattern: /<meta[^>]+property=["']og:url["'][^>]*content=["']([^"']*)["']/i },
  { key: 'og:type', pattern: /<meta[^>]+property=["']og:type["'][^>]*content=["']([^"']*)["']/i },
  { key: 'twitter:card', pattern: /<meta[^>]+name=["']twitter:card["'][^>]*content=["']([^"']*)["']/i },
  { key: 'canonical', pattern: /<link[^>]+rel=["']canonical["'][^>]*href=["']([^"']*)["']/i },
];

async function fetchHTML(url) {
  try {
    const res = await fetch(url, { redirect: 'follow' });
    if (!res.ok) return { url, error: `HTTP ${res.status}` };
    const html = await res.text();
    return { url, html, status: res.status };
  } catch (e) {
    return { url, error: e.message };
  }
}

function extractMeta(html) {
  const row = {};
  for (const { key, pattern } of META_KEYS) {
    const m = pattern.exec(html);
    row[key] = m ? m[1] : null;
  }
  return row;
}

async function buildMetaCoverage() {
  const results = await Promise.all(URLS.map(fetchHTML));
  let md = `# Meta (OGP / Twitter / canonical) Coverage — 2026-04-20

Source: 13 URL 並列 fetch → HTML 内の meta/link タグを正規表現抽出

✅ = 埋まっている (文字列あり) / ❌ = 欠落

| URL | og:title | og:description | og:image | og:url | og:type | twitter:card | canonical |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
`;
  const details = [];
  let totalMissing = 0;
  for (const r of results) {
    if (r.error) {
      md += `| ${r.url} | ❌ fetch: ${r.error} | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ |\n`;
      totalMissing += 7;
      continue;
    }
    const row = extractMeta(r.html);
    const mark = v => v ? '✅' : '❌';
    md += `| ${r.url} | ${mark(row['og:title'])} | ${mark(row['og:description'])} | ${mark(row['og:image'])} | ${mark(row['og:url'])} | ${mark(row['og:type'])} | ${mark(row['twitter:card'])} | ${mark(row['canonical'])} |\n`;
    for (const [k, v] of Object.entries(row)) { if (!v) totalMissing++; }
    details.push({ url: r.url, ...row });
  }

  md += `\n## 欠落合計\n\n**${totalMissing}** fields missing (across ${URLS.length} URLs × 7 meta fields = ${URLS.length*7} total).\n\n`;
  md += `## 詳細 (取得できた値)\n\n`;
  for (const d of details) {
    md += `### ${d.url}\n`;
    for (const k of META_KEYS.map(x=>x.key)) {
      const v = d[k];
      md += `- **${k}**: ${v ? '`'+(v.length>120?v.slice(0,120)+'...':v)+'`' : '❌ missing'}\n`;
    }
    md += `\n`;
  }
  fs.writeFileSync(`${OUT_DIR}/meta-coverage.md`, md);
  return totalMissing;
}

async function buildSitemapCoverage() {
  const res = await fetch(`${BASE}/sitemap.xml`);
  let sitemapText = null;
  let sitemapError = null;
  if (res.ok) {
    sitemapText = await res.text();
  } else {
    sitemapError = `HTTP ${res.status}`;
  }
  const locs = [];
  if (sitemapText) {
    const re = /<loc>\s*([^<\s]+)\s*<\/loc>/gi;
    let m;
    while ((m = re.exec(sitemapText)) != null) locs.push(m[1].trim());
  }
  function normalize(u) {
    // Vercel cleanUrls 対応: .html 拡張子を除去。trailing slash も除去 (ただし root '/' は残す)。
    try {
      const x = new URL(u);
      let path = x.pathname;
      if (path.endsWith('.html')) path = path.slice(0, -5);
      if (path.length > 1 && path.endsWith('/')) path = path.slice(0, -1);
      return x.origin + path;
    } catch { return u; }
  }
  const sitemapSet = new Set(locs.map(normalize));
  const targetSet = new Set(URLS.map(normalize));

  const included = URLS.filter(u => sitemapSet.has(normalize(u)));
  const missing = URLS.filter(u => !sitemapSet.has(normalize(u)));
  const extras = [...sitemapSet].filter(u => !targetSet.has(u));

  let md = `# Sitemap Coverage — 2026-04-20

Source: ${BASE}/sitemap.xml

`;
  if (sitemapError) {
    md += `❌ sitemap.xml の取得に失敗: **${sitemapError}**\n\n`;
  } else {
    md += `- sitemap.xml 内の <loc> 数: **${locs.length}**\n`;
    md += `- 対象 13 URL のうち包含: **${included.length}** / 欠落: **${missing.length}**\n`;
    md += `- sitemap にあるが対象13 URL に含まれない (参考) : ${extras.length}\n\n`;
  }

  md += `## 13 URL 包含状況\n\n| URL | in sitemap |\n|---|:-:|\n`;
  for (const u of URLS) {
    md += `| ${u} | ${sitemapSet.has(normalize(u)) ? '✅' : '❌'} |\n`;
  }

  if (missing.length) {
    md += `\n## 欠落 URL\n\n`;
    for (const u of missing) md += `- ${u}\n`;
  }

  if (extras.length) {
    md += `\n## sitemap にのみ存在 (参考, 先頭30件まで)\n\n`;
    for (const u of extras.slice(0,30)) md += `- ${u}\n`;
    if (extras.length > 30) md += `- ... +${extras.length-30} more\n`;
  }

  fs.writeFileSync(`${OUT_DIR}/sitemap-coverage.md`, md);
  return { sitemapError, missingCount: missing.length, sitemapEntryCount: locs.length };
}

const [totalMissing, sitemap] = await Promise.all([
  buildMetaCoverage(),
  buildSitemapCoverage(),
]);

console.log('meta total missing:', totalMissing);
console.log('sitemap:', sitemap);
