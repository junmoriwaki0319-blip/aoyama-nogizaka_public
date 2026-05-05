const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const PREVIEW = process.argv[2];
const OUT = process.argv[3];
const THRESHOLD = 44;
if (!PREVIEW || !OUT) {
  console.error('usage: node measure-preview-after.js <preview-base-url> <out-json-path>');
  process.exit(1);
}

const CLICKABLE = [
  'a','button','input[type="button"]','input[type="submit"]','input[type="reset"]',
  '[role="button"]','[role="link"]','[onclick]','label[for]'
].join(',');

(async () => {
  const iphone = devices['iPhone SE'];
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ ...iphone, viewport: { width: 375, height: 667 } });
  const page = await ctx.newPage();
  const url = PREVIEW.replace(/\/$/, '') + '/risk-assessment';
  await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(500);

  const result = await page.evaluate((args) => {
    const T = args.threshold;
    function isVisible(el) {
      if (!el.isConnected) return false;
      const r = el.getBoundingClientRect();
      if (r.width === 0 && r.height === 0) return false;
      let n = el;
      while (n && n.nodeType === 1) {
        const cs = getComputedStyle(n);
        if (cs.display === 'none' || cs.visibility === 'hidden' || cs.visibility === 'collapse') return false;
        if (parseFloat(cs.opacity) === 0) return false;
        n = n.parentElement;
      }
      return true;
    }
    const violations = [];
    document.querySelectorAll(args.selector).forEach((el) => {
      if (!isVisible(el)) return;
      const r = el.getBoundingClientRect();
      const w = Math.round(r.width * 100) / 100, h = Math.round(r.height * 100) / 100;
      if (w >= T && h >= T) return;
      violations.push({
        tag: el.tagName.toLowerCase(),
        classes: el.className || '',
        text: (el.innerText || el.textContent || '').trim().replace(/\s+/g,' ').slice(0, 80),
        w, h,
        deficit: { w: Math.max(0, T - w), h: Math.max(0, T - h) },
      });
    });
    // Capture cta-primary/secondary measurements specifically
    const ctas = Array.from(document.querySelectorAll('.cta-links a')).map((el) => {
      const r = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      return {
        text: el.textContent.trim(),
        className: el.className,
        width: Math.round(r.width * 100) / 100,
        height: Math.round(r.height * 100) / 100,
        padding: cs.padding,
        border: cs.border,
        boxSizing: cs.boxSizing,
      };
    });
    return { violations, ctas };
  }, { selector: CLICKABLE, threshold: THRESHOLD });

  const ctaPrimary = result.ctas.find((c) => /\bcta-primary\b/.test(c.className));
  const passed = ctaPrimary && ctaPrimary.width >= THRESHOLD && ctaPrimary.height >= THRESHOLD;

  const payload = {
    measuredAt: new Date().toISOString(),
    previewBase: PREVIEW,
    pageUrl: url,
    viewport: { width: 375, height: 667 },
    device: 'iPhone SE (Playwright preset)',
    threshold: { width: THRESHOLD, height: THRESHOLD },
    wcag: '2.5.5 Enhanced',
    ctaPrimary,
    ctaSecondary: result.ctas.find((c) => /\bcta-secondary\b/.test(c.className)),
    ctaPrimaryPass: passed,
    pageViolations: result.violations,
    summary: {
      totalViolationsOnPage: result.violations.length,
      ctaPrimaryStillViolating: result.violations.some((v) => /\bcta-primary\b/.test(v.classes)),
    },
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(payload, null, 2));
  console.log('WROTE:', OUT);
  console.log('cta-primary:', JSON.stringify(ctaPrimary));
  console.log('cta-primary PASS:', passed);
  console.log('page violations remaining:', result.violations.length);
  result.violations.forEach((v, i) => console.log(`  ${i+1}. <${v.tag}${v.classes?' class="'+v.classes+'"':''}> "${v.text}" ${v.w}×${v.h}`));
  await browser.close();
  process.exit(passed ? 0 : 2);
})();
