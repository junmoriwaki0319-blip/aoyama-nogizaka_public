const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = process.argv[2];
const OUT = process.argv[3];
if (!BASE || !OUT) { console.error('usage: node smoke-test-preview.js <base> <out-json>'); process.exit(1); }

const GA_PAGES = [
  '/',
  '/risk-assessment',
  '/activist-dashboard',
  '/saas',
  '/food-service',
];
const CSP_CHECK_PAGES = ['/', '/activist-dashboard', '/saas'];

function isGaRequest(url) {
  return /google-analytics\.com\/(g|r)\/collect|googletagmanager\.com\/gtag\/js/.test(url);
}

async function testPage(context, base, pathSeg, opts = {}) {
  const page = await context.newPage();
  const gaHits = [];
  const cspViolations = [];
  const otherConsoleErrors = [];
  page.on('request', (req) => {
    const u = req.url();
    if (isGaRequest(u)) gaHits.push({ url: u.slice(0, 200), method: req.method() });
  });
  page.on('console', (msg) => {
    const text = msg.text();
    const type = msg.type();
    if (/Content Security Policy|CSP|Refused to execute|Refused to load|violates the following/i.test(text)) {
      cspViolations.push({ type, text: text.slice(0, 300) });
    } else if (type === 'error') {
      otherConsoleErrors.push(text.slice(0, 200));
    }
  });
  page.on('pageerror', (err) => {
    otherConsoleErrors.push('[pageerror] ' + String(err).slice(0, 200));
  });

  const url = base.replace(/\/$/, '') + pathSeg;
  const resp = await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
  // ga-loader.js defers GA load by 3s; wait long enough for page_view to fire
  await page.waitForTimeout(5000);

  // Check that the page's CSS print→all swap fired (looking at Noto Sans JP actually applied)
  const bodyFont = await page.evaluate(() => getComputedStyle(document.body).fontFamily);
  const nonPrintStylesheets = await page.evaluate(() => {
    return Array.from(document.querySelectorAll('link[rel="stylesheet"]')).map((l) => ({
      href: l.href,
      media: l.media,
    }));
  });

  const result = {
    pathSeg,
    finalUrl: page.url(),
    status: resp ? resp.status() : null,
    gaHitsCount: gaHits.length,
    gaHitsSample: gaHits.slice(0, 3),
    cspViolations,
    otherConsoleErrors,
    bodyFontFamily: bodyFont,
    nonPrintStylesheets,
  };
  await page.close();
  return result;
}

async function testRedirect(context, base, fromPath) {
  const page = await context.newPage();
  const url = base.replace(/\/$/, '') + fromPath;
  let resp;
  try {
    resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  } catch (e) {
    await page.close();
    return { fromPath, error: e.message };
  }
  await page.waitForTimeout(1500);
  const finalUrl = page.url();
  const status = resp ? resp.status() : null;
  await page.close();
  return { fromPath, finalUrl, status, redirected: finalUrl !== url };
}

async function test404(context, base) {
  const page = await context.newPage();
  const url = base.replace(/\/$/, '') + '/definitely-not-a-real-page-xyz123';
  const resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  const title = await page.title();
  const has404Marker = await page.evaluate(() => {
    const text = document.body.innerText || '';
    return /404|ページが見つか|not found/i.test(text);
  });
  await page.close();
  return { url, status: resp ? resp.status() : null, title, has404Marker };
}

(async () => {
  const iphone = devices['iPhone SE'];
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ ...iphone, viewport: { width: 375, height: 667 } });

  const report = { testedAt: new Date().toISOString(), base: BASE };

  report.gaAndCspChecks = [];
  for (const p of GA_PAGES) {
    process.stdout.write(`GA/CSP check: ${p} ... `);
    const r = await testPage(ctx, BASE, p);
    report.gaAndCspChecks.push(r);
    console.log(`GA=${r.gaHitsCount} CSP_viol=${r.cspViolations.length} other_errs=${r.otherConsoleErrors.length}`);
  }

  console.log('\n--- redirect tests ---');
  report.redirectTests = [];
  for (const p of ['/game-content.html', '/game-content']) {
    const r = await testRedirect(ctx, BASE, p);
    report.redirectTests.push(r);
    console.log(p, '->', JSON.stringify(r));
  }

  console.log('\n--- 404 test ---');
  report.the404Test = await test404(ctx, BASE);
  console.log(JSON.stringify(report.the404Test));

  // Summary
  const cspFailurePages = report.gaAndCspChecks.filter((c) => c.cspViolations.length > 0);
  const gaFailurePages = report.gaAndCspChecks.filter((c) => c.gaHitsCount === 0);
  report.summary = {
    pagesChecked: GA_PAGES.length,
    pagesWithCspViolations: cspFailurePages.map((c) => c.pathSeg),
    pagesWithoutGaHits: gaFailurePages.map((c) => c.pathSeg),
    gameContentRedirectWorking: report.redirectTests.some((r) => r.redirected && /\/activist-dashboard|\/saas|\/food-service|\/ad-agency|\/digital-media|\/entertainment/.test(r.finalUrl || '')),
    the404Recognized: report.the404Test.has404Marker,
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(report, null, 2));
  console.log('\n=== SUMMARY ===');
  console.log(JSON.stringify(report.summary, null, 2));
  console.log('\nWROTE:', OUT);
  await browser.close();
})();
