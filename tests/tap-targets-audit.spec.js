// @ts-check
const { test } = require('@playwright/test');
const fs = require('fs');
const path = require('path');

// ============================================================
// Tap-Targets Audit (Playwright based replacement for Lighthouse)
// Lighthouse's default threshold: tap targets must be >= 48x48 CSS px.
// Mobile viewport: Moto G Power (Lighthouse default) = 412x823 @ DPR 1.75
// We use 360x640 (the classic Lighthouse mobile emulation CSS px).
//
// 実行: npx playwright test tests/tap-targets-audit.spec.js --project=smoke
// 出力: tests/tap-targets-report.json  (pages × failing targets)
// ============================================================

const PAGES = [
  '/',
  '/team',
  '/saas',
  '/food-service',
  '/ad-agency',
  '/digital-media',
  '/entertainment-sector-dashboard',
  '/activist-dashboard',
  '/activist-screener',
  '/activist-campaigns',
  '/risk-assessment',
  '/privacy',
];

const MIN_SIZE = 48;
const VIEWPORT = { width: 360, height: 640 };

const results = {};

test.describe('tap-targets audit (mobile 360x640, min 48x48)', () => {
  test.describe.configure({ mode: 'serial' });

  for (const route of PAGES) {
    test(`audit ${route}`, async ({ page }) => {
      await page.setViewportSize(VIEWPORT);
      const res = await page.goto(route, { waitUntil: 'networkidle' });
      if (!res || !res.ok()) {
        results[route] = { error: `HTTP ${res ? res.status() : 'no-response'}` };
        return;
      }

      const failing = await page.evaluate((minSize) => {
        const SELECTOR = 'a[href], button, input:not([type="hidden"]), select, textarea, [role="button"], [role="link"], [onclick], [tabindex]:not([tabindex="-1"])';
        const nodes = Array.from(document.querySelectorAll(SELECTOR));
        const out = [];
        for (const el of nodes) {
          const rect = el.getBoundingClientRect();
          if (rect.width === 0 && rect.height === 0) continue;
          const style = getComputedStyle(el);
          if (style.visibility === 'hidden' || style.display === 'none') continue;
          if (rect.width < minSize || rect.height < minSize) {
            const label = (el.getAttribute('aria-label') || el.textContent || el.getAttribute('title') || '').trim().replace(/\s+/g, ' ').slice(0, 80);
            const path = (() => {
              const parts = [];
              let cur = el;
              while (cur && cur !== document.body && parts.length < 4) {
                let sel = cur.tagName.toLowerCase();
                if (cur.id) { sel += '#' + cur.id; parts.unshift(sel); break; }
                if (cur.className && typeof cur.className === 'string') {
                  const cls = cur.className.trim().split(/\s+/).slice(0, 2).join('.');
                  if (cls) sel += '.' + cls;
                }
                parts.unshift(sel);
                cur = cur.parentElement;
              }
              return parts.join(' > ');
            })();
            out.push({
              selector: path,
              label,
              width: Math.round(rect.width),
              height: Math.round(rect.height),
              top: Math.round(rect.top),
              left: Math.round(rect.left),
              href: el.getAttribute('href') || null,
            });
          }
        }
        return out;
      }, MIN_SIZE);

      results[route] = { count: failing.length, items: failing };
      console.log(`\n--- ${route}: ${failing.length} failing ---`);
      failing.slice(0, 20).forEach((f, i) => {
        console.log(`  #${i + 1} ${f.width}x${f.height} | ${f.selector} | "${f.label}"`);
      });
    });
  }

  test.afterAll(async () => {
    const reportPath = path.join(__dirname, 'tap-targets-report.json');
    fs.writeFileSync(reportPath, JSON.stringify({
      viewport: VIEWPORT,
      minSize: MIN_SIZE,
      timestamp: new Date().toISOString(),
      pages: results,
    }, null, 2));
    const totals = Object.entries(results).map(([r, v]) => `${r}: ${v.count ?? v.error}`).join('\n');
    console.log(`\n=== tap-targets summary ===\n${totals}\nreport: ${reportPath}`);
  });
});
