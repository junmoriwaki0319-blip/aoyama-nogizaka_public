const { chromium, devices } = require('playwright');
const THRESHOLD = 44;
const url = process.argv[2];
if (!url) { console.error('usage: node verify-single-page.js <url>'); process.exit(1); }
const CLICKABLE_SELECTOR = [
  'a','button','input[type="button"]','input[type="submit"]','input[type="reset"]',
  '[role="button"]','[role="link"]','[onclick]','label[for]'
].join(',');

(async () => {
  const iphone = devices['iPhone SE'];
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ ...iphone, viewport: { width: 375, height: 667 } });
  const page = await ctx.newPage();
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
    const out = [];
    document.querySelectorAll(args.selector).forEach((el) => {
      if (!isVisible(el)) return;
      const r = el.getBoundingClientRect();
      const w = Math.round(r.width * 100) / 100, h = Math.round(r.height * 100) / 100;
      if (w >= T && h >= T) return;
      out.push({
        tag: el.tagName.toLowerCase(),
        classes: el.className || '',
        text: (el.innerText || el.textContent || '').trim().replace(/\s+/g,' ').slice(0, 60),
        w, h,
      });
    });
    return out;
  }, { selector: CLICKABLE_SELECTOR, threshold: THRESHOLD });
  console.log('URL:', url);
  console.log('Violations (<44×44):', result.length);
  result.forEach((v, i) => console.log(`  ${i+1}. <${v.tag}${v.classes?' class="'+v.classes+'"':''}> "${v.text}" ${v.w}×${v.h}`));
  await browser.close();
})();
