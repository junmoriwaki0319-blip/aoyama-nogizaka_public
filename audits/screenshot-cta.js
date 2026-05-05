const { chromium, devices } = require('playwright');
const path = require('path');
const fs = require('fs');

(async () => {
  const label = process.argv[2] || 'before';
  const url = process.argv[3] || 'https://aoyama-nogizaka.com/risk-assessment';
  const outDir = path.join(__dirname, '20260418', 'screenshots');
  fs.mkdirSync(outDir, { recursive: true });

  const iphone = devices['iPhone SE'];
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ ...iphone, viewport: { width: 375, height: 667 } });
  const page = await ctx.newPage();

  await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(500);

  const ctaLocator = page.locator('a.cta-primary').first();
  await ctaLocator.scrollIntoViewIfNeeded();
  await page.waitForTimeout(300);

  const box = await ctaLocator.boundingBox();
  const cardLocator = page.locator('.cta-card').first();
  const cardBox = await cardLocator.boundingBox();

  const ctaFull = path.join(outDir, `cta-card-${label}.png`);
  await cardLocator.screenshot({ path: ctaFull });

  const primaryOnly = path.join(outDir, `cta-primary-${label}.png`);
  await ctaLocator.screenshot({ path: primaryOnly });

  const measurements = await page.evaluate(() => {
    const els = Array.from(document.querySelectorAll('.cta-links a'));
    return els.map((el) => {
      const r = el.getBoundingClientRect();
      const cs = window.getComputedStyle(el);
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
  });

  const measPath = path.join(outDir, `measurements-${label}.json`);
  fs.writeFileSync(measPath, JSON.stringify(measurements, null, 2));

  console.log(JSON.stringify({ label, ctaBox: box, cardBox, ctaFull, primaryOnly, measPath, measurements }, null, 2));
  await browser.close();
})();
