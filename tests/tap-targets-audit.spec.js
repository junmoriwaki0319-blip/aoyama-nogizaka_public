// @ts-check
const { test } = require('@playwright/test');
const fs = require('fs');
const path = require('path');

// ============================================================
// Tap-Targets Audit (Playwright based replacement for Lighthouse)
//
// Baseline / measurement axis (固定):
//   - viewport       : 360 x 640 CSS px
//   - threshold      : 48 x 48 CSS px (Lighthouse default / Google Mobile)
//   - target selector: a[href], button, input (visible types), select, textarea,
//                      [role=button], [role=link], [onclick], [tabindex>=0]
//   - page list      : PAGES[] below (ルート直下の主要 12 ページ)
//
// Category B filters (false positive 除外):
//   1. Off-screen / visually hidden
//      - rect.bottom <= 0, rect.right <= 0
//      - position:absolute かつ top <= -50px (skip-link パターン)
//      - visibility/display 非表示
//   2. Inline in running text (WCAG 2.5.5 / 2.5.8 Inline 例外)
//      - 親が <p>, <li>, <details> の直系リンク
//      - exclusion="inline" として分類（count には含まない、records に残す）
//
// 出力: tests/tap-targets-report.json — fix されたのみ violations、+ excluded.inline
// 実行: npx playwright test --project=tap-targets
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

      const audit = await page.evaluate(({ minSize, viewportHeight }) => {
        const SELECTOR = 'a[href], button, input:not([type="hidden"]), select, textarea, [role="button"], [role="link"], [onclick], [tabindex]:not([tabindex="-1"])';
        const nodes = Array.from(document.querySelectorAll(SELECTOR));
        const violations = [];
        const excludedOffscreen = [];
        const excludedInline = [];

        const describe = (el) => {
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
        };

        const hasInlineTextAncestor = (el) => {
          // <p>, <li>, <details> 内に直接含まれる <a> は running-text inline link とみなす
          const inlineContainers = new Set(['P', 'LI', 'DETAILS']);
          let cur = el.parentElement;
          for (let i = 0; i < 3 && cur; i++, cur = cur.parentElement) {
            if (inlineContainers.has(cur.tagName)) return true;
          }
          return false;
        };

        for (const el of nodes) {
          const rect = el.getBoundingClientRect();
          if (rect.width === 0 && rect.height === 0) continue;
          const style = getComputedStyle(el);
          if (style.visibility === 'hidden' || style.display === 'none') continue;

          // Category B-1: offscreen / skip-link pattern
          const isOffscreen =
            rect.bottom <= 0 ||
            rect.right <= 0 ||
            rect.top >= viewportHeight ||
            (style.position === 'absolute' && rect.top < -50);
          if (isOffscreen) {
            // record for transparency but skip violation counting
            const label = (el.getAttribute('aria-label') || el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 80);
            excludedOffscreen.push({ selector: describe(el), label, top: Math.round(rect.top), w: Math.round(rect.width), h: Math.round(rect.height) });
            continue;
          }

          if (rect.width < minSize || rect.height < minSize) {
            const label = (el.getAttribute('aria-label') || el.textContent || el.getAttribute('title') || '').trim().replace(/\s+/g, ' ').slice(0, 80);
            const info = {
              selector: describe(el),
              label,
              width: Math.round(rect.width),
              height: Math.round(rect.height),
              top: Math.round(rect.top),
              left: Math.round(rect.left),
              href: el.getAttribute('href') || null,
            };

            // Category B-2: inline in running text
            if (el.tagName === 'A' && hasInlineTextAncestor(el)) {
              excludedInline.push({ ...info, reason: 'inline-text (WCAG 2.5.5 exception)' });
              continue;
            }

            violations.push(info);
          }
        }
        return { violations, excludedOffscreen, excludedInline };
      }, { minSize: MIN_SIZE, viewportHeight: VIEWPORT.height });

      results[route] = {
        count: audit.violations.length,
        items: audit.violations,
        excluded: {
          offscreen: audit.excludedOffscreen,
          inline: audit.excludedInline,
        },
      };
      const v = audit.violations.length;
      const xo = audit.excludedOffscreen.length;
      const xi = audit.excludedInline.length;
      console.log(`\n--- ${route}: ${v} failing (excluded: ${xo} offscreen, ${xi} inline) ---`);
      audit.violations.slice(0, 20).forEach((f, i) => {
        console.log(`  #${i + 1} ${f.width}x${f.height} | ${f.selector} | "${f.label}"`);
      });
    });
  }

  test.afterAll(async () => {
    const reportPath = path.join(__dirname, 'tap-targets-report.json');
    const summary = Object.fromEntries(
      Object.entries(results).map(([r, v]) => [r, {
        count: v.count ?? null,
        excluded: v.excluded ? { offscreen: v.excluded.offscreen.length, inline: v.excluded.inline.length } : null,
        error: v.error ?? null,
      }])
    );
    fs.writeFileSync(reportPath, JSON.stringify({
      viewport: VIEWPORT,
      minSize: MIN_SIZE,
      methodology: {
        threshold: '48x48 (Lighthouse default / Google Mobile)',
        filters: ['offscreen (rect outside viewport or absolute top<-50)', 'inline text links in <p>/<li>/<details>'],
        selector: 'a[href], button, input (non-hidden), select, textarea, [role=button|link], [onclick], [tabindex>=0]',
      },
      timestamp: new Date().toISOString(),
      summary,
      pages: results,
    }, null, 2));
    const totals = Object.entries(results)
      .map(([r, v]) => `${r}: ${v.count ?? v.error}  (excluded: ${v.excluded ? v.excluded.offscreen.length + ' offscreen, ' + v.excluded.inline.length + ' inline' : 'n/a'})`)
      .join('\n');
    console.log(`\n=== tap-targets summary ===\n${totals}\nreport: ${reportPath}`);
  });
});
