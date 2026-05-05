const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const THRESHOLD = 44;
const BASE = 'https://aoyama-nogizaka.com';
const PAGES = [
  '/',
  '/team',
  '/privacy',
  '/news/',
  '/activist-dashboard.html',
  '/risk-assessment.html',
  '/activist-screener.html',
  '/food-service.html',
  '/saas.html',
  '/ad-agency.html',
  '/digital-media.html',
  '/entertainment-sector-dashboard.html',
];

const CLICKABLE_SELECTOR = [
  'a',
  'button',
  'input[type="button"]',
  'input[type="submit"]',
  'input[type="reset"]',
  '[role="button"]',
  '[role="link"]',
  '[onclick]',
  'label[for]',
].join(',');

function buildCssPath(el) {
  const parts = [];
  let node = el;
  while (node && node.nodeType === 1 && parts.length < 6) {
    let seg = node.tagName.toLowerCase();
    if (node.id) {
      seg += '#' + node.id;
      parts.unshift(seg);
      break;
    }
    if (node.className && typeof node.className === 'string') {
      const cls = node.className.trim().split(/\s+/).slice(0, 2).join('.');
      if (cls) seg += '.' + cls;
    }
    const parent = node.parentNode;
    if (parent) {
      const siblings = Array.from(parent.children).filter(
        (c) => c.tagName === node.tagName
      );
      if (siblings.length > 1) {
        seg += `:nth-of-type(${siblings.indexOf(node) + 1})`;
      }
    }
    parts.unshift(seg);
    node = node.parentElement;
  }
  return parts.join(' > ');
}

async function measurePage(context, url) {
  const page = await context.newPage();
  page.on('pageerror', () => {});
  const violations = [];
  try {
    await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(800);
    const evalArgs = { selector: CLICKABLE_SELECTOR, threshold: THRESHOLD };
    const result = await page.evaluate((args) => {
      const SELECTOR = args.selector;
      const T = args.threshold;

      function isVisible(el) {
        if (!el.isConnected) return false;
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) return false;
        let node = el;
        while (node && node.nodeType === 1) {
          const cs = window.getComputedStyle(node);
          if (cs.display === 'none' || cs.visibility === 'hidden' || cs.visibility === 'collapse') return false;
          if (parseFloat(cs.opacity) === 0) return false;
          node = node.parentElement;
        }
        return true;
      }

      function cssPath(el) {
        const parts = [];
        let node = el;
        let depth = 0;
        while (node && node.nodeType === 1 && depth < 6) {
          let seg = node.tagName.toLowerCase();
          if (node.id) {
            seg += '#' + node.id;
            parts.unshift(seg);
            break;
          }
          if (node.className && typeof node.className === 'string') {
            const cls = node.className.trim().split(/\s+/).filter(Boolean).slice(0, 2).join('.');
            if (cls) seg += '.' + cls;
          }
          const parent = node.parentNode;
          if (parent && parent.children) {
            const sameTag = Array.from(parent.children).filter(
              (c) => c.tagName === node.tagName
            );
            if (sameTag.length > 1) {
              seg += ':nth-of-type(' + (sameTag.indexOf(node) + 1) + ')';
            }
          }
          parts.unshift(seg);
          node = node.parentElement;
          depth++;
        }
        return parts.join(' > ');
      }

      const results = [];
      const nodes = document.querySelectorAll(SELECTOR);
      nodes.forEach((el) => {
        if (!isVisible(el)) return;
        const r = el.getBoundingClientRect();
        const w = Math.round(r.width * 100) / 100;
        const h = Math.round(r.height * 100) / 100;
        if (w >= T && h >= T) return;
        const text = (el.innerText || el.textContent || el.value || '').trim().replace(/\s+/g, ' ').slice(0, 60);
        const tag = el.tagName.toLowerCase();
        const role = el.getAttribute('role') || '';
        const href = el.getAttribute('href') || '';
        const onclickFlag = el.hasAttribute('onclick');
        const labelFor = el.getAttribute('for') || '';
        const id = el.id || '';
        const classList = (el.className && typeof el.className === 'string') ? el.className.trim() : '';
        results.push({
          tag,
          id,
          classes: classList,
          role,
          href,
          onclick: onclickFlag,
          labelFor,
          text,
          current: { w, h },
          deficit: { w: Math.max(0, T - w), h: Math.max(0, T - h) },
          selector: cssPath(el),
          outerHTML: el.outerHTML.slice(0, 240).replace(/\s+/g, ' '),
        });
      });
      return results;
    }, evalArgs);

    for (const item of result) {
      violations.push({ page: url, ...item });
    }
  } catch (e) {
    console.error('PAGE_ERROR', url, e.message);
    violations.push({ page: url, error: e.message });
  } finally {
    await page.close();
  }
  return violations;
}

function classifyKind(v) {
  const cls = (v.classes || '').toLowerCase();
  const sel = (v.selector || '').toLowerCase();
  const tag = v.tag;
  if (tag === 'button' || v.role === 'button') return 'button';
  if (v.tag === 'a' && /nav|menu|header|hamburger/.test(cls + ' ' + sel)) return 'nav-link';
  if (/footer/.test(sel)) return 'footer-link';
  if (v.tag === 'a') return 'link';
  if (v.tag === 'input') return 'input-' + (v.href || '');
  if (v.labelFor) return 'label';
  return tag;
}

function groupBy(arr, fn) {
  const m = new Map();
  for (const x of arr) {
    const k = fn(x);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(x);
  }
  return m;
}

(async () => {
  const iphone = devices['iPhone SE'] || {
    viewport: { width: 375, height: 667 },
    userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1',
    deviceScaleFactor: 2,
    isMobile: true,
    hasTouch: true,
  };
  // Force viewport to 375x667 as user requested
  const ctxOptions = { ...iphone, viewport: { width: 375, height: 667 } };

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext(ctxOptions);
  const all = [];
  const perPage = {};

  for (const p of PAGES) {
    const url = BASE + p;
    process.stdout.write(`Measuring ${url} ... `);
    const v = await measurePage(context, url);
    const errs = v.filter((x) => x.error);
    const real = v.filter((x) => !x.error);
    console.log(`${real.length} violations${errs.length ? ' (errors: ' + errs.length + ')' : ''}`);
    perPage[url] = { total: real.length, errors: errs };
    all.push(...real);
  }

  await browser.close();

  const outDir = path.join(__dirname, '20260418');
  fs.mkdirSync(outDir, { recursive: true });
  const jsonPath = path.join(outDir, 'tap-targets-enhanced-44px.json');
  const mdPath = path.join(outDir, 'tap-targets-enhanced-44px.md');
  fs.writeFileSync(jsonPath, JSON.stringify(all, null, 2));

  // Markdown: grouped by page → kind
  let md = `# WCAG 2.5.5 Enhanced (44×44px) タッチターゲット違反レポート\n\n`;
  md += `- 計測日: 2026-04-18\n`;
  md += `- Viewport: 375×667 (iPhone SE), Chromium headless\n`;
  md += `- 閾値: width < 44 OR height < 44\n`;
  md += `- 対象セレクタ: \`${CLICKABLE_SELECTOR}\`\n`;
  md += `- 除外: display:none, visibility:hidden/collapse, opacity:0, 0×0矩形\n\n`;
  md += `## サマリ\n\n| ページ | 違反数 |\n|---|---:|\n`;
  for (const p of PAGES) {
    const url = BASE + p;
    md += `| ${p} | ${perPage[url].total} |\n`;
  }
  md += `| **合計** | **${all.length}** |\n\n`;

  md += `## ページ別詳細\n\n`;
  for (const p of PAGES) {
    const url = BASE + p;
    const pageViolations = all.filter((x) => x.page === url);
    if (pageViolations.length === 0) continue;
    md += `### ${p} (${pageViolations.length}件)\n\n`;
    const byKind = groupBy(pageViolations, classifyKind);
    for (const [kind, arr] of byKind) {
      md += `#### ${kind} (${arr.length}件)\n\n`;
      md += `| # | selector | text | 現在(w×h) | 不足(w/h) | href/for |\n`;
      md += `|---:|---|---|---:|---:|---|\n`;
      arr.forEach((v, i) => {
        const hrefOrFor = v.href || v.labelFor || (v.onclick ? '[onclick]' : '');
        const textCell = v.text ? v.text.replace(/\|/g, '\\|') : '(no text)';
        md += `| ${i + 1} | \`${v.selector}\` | ${textCell} | ${v.current.w}×${v.current.h} | ${v.deficit.w}/${v.deficit.h} | ${hrefOrFor} |\n`;
      });
      md += `\n`;
    }
  }

  fs.writeFileSync(mdPath, md);
  console.log('\nWROTE:', jsonPath);
  console.log('WROTE:', mdPath);
  console.log('TOTAL violations:', all.length);
})();
