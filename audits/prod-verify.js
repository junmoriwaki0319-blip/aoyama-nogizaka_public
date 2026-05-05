const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = process.argv[2] || 'https://aoyama-nogizaka.com';
const OUT = process.argv[3] || 'audits/20260419/prod-deploy-verification.json';
const THRESHOLD = 44;

const GA_PAGES = ['/', '/risk-assessment', '/activist-dashboard', '/saas', '/food-service'];
const FIREBASE_PAGES = ['/activist-dashboard', '/saas', '/food-service'];

function isGaRequest(u) {
  return /google-analytics\.com\/(g|r)\/collect|googletagmanager\.com\/gtag\/js/.test(u);
}
function isFirebaseInstallations(u) {
  return /firebaseinstallations\.googleapis\.com\//.test(u);
}

async function checkPage(ctx, base, pathSeg) {
  const page = await ctx.newPage();
  const gaHits = [];
  const firebaseInstallResponses = [];
  const cspViolations = [];
  const otherConsoleErrors = [];
  page.on('request', (req) => {
    if (isGaRequest(req.url())) gaHits.push({ url: req.url().slice(0, 200) });
  });
  page.on('response', async (res) => {
    const u = res.url();
    if (isFirebaseInstallations(u)) {
      firebaseInstallResponses.push({ url: u.slice(0, 200), status: res.status() });
    }
  });
  page.on('console', (msg) => {
    const text = msg.text();
    const type = msg.type();
    if (/Content Security Policy|CSP|Refused to execute|Refused to load|violates the following/i.test(text)) {
      cspViolations.push({ type, text: text.slice(0, 600) });
    } else if (type === 'error') {
      otherConsoleErrors.push(text.slice(0, 300));
    }
  });
  page.on('pageerror', (err) => otherConsoleErrors.push('[pageerror] ' + String(err).slice(0, 300)));

  const url = base.replace(/\/$/, '') + pathSeg;
  let status = null;
  try {
    const resp = await page.goto(url, { waitUntil: 'networkidle', timeout: 45000 });
    status = resp ? resp.status() : null;
  } catch (e) {
    await page.close();
    return { pathSeg, error: e.message };
  }
  // ga-loader.js defers GA by 3s; allow 5-6s for GA + firebase installations roundtrip
  await page.waitForTimeout(6000);

  // Specific cta-primary measurement on /risk-assessment
  let cta = null;
  if (pathSeg === '/risk-assessment') {
    cta = await page.evaluate(() => {
      const el = document.querySelector('a.cta-primary');
      if (!el) return null;
      const r = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      return {
        text: el.textContent.trim(),
        width: Math.round(r.width * 100) / 100,
        height: Math.round(r.height * 100) / 100,
        border: cs.border,
        boxSizing: cs.boxSizing,
      };
    });
  }

  await page.close();
  return { pathSeg, finalUrl: url, status, gaHits, firebaseInstallResponses, cspViolations, otherConsoleErrors, cta };
}

(async () => {
  const iphone = devices['iPhone SE'];
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ ...iphone, viewport: { width: 375, height: 667 } });

  const report = { verifiedAt: new Date().toISOString(), base: BASE };
  report.pages = [];
  for (const p of GA_PAGES) {
    process.stdout.write(`verify ${p} ... `);
    const r = await checkPage(ctx, BASE, p);
    report.pages.push(r);
    const gaN = (r.gaHits || []).length;
    const fbN = (r.firebaseInstallResponses || []).length;
    const fbOk = (r.firebaseInstallResponses || []).filter(x => x.status === 200).length;
    const cspReal = (r.cspViolations || []).filter(v => !/vercel\.live/i.test(v.text) && !/delivered via a <meta>/i.test(v.text)).length;
    console.log(`GA=${gaN} FB=${fbN}(${fbOk}x200) CSP_new=${cspReal}`);
  }
  await browser.close();

  // Evaluate criteria
  const risk = report.pages.find(p => p.pathSeg === '/risk-assessment');
  const ctaPass = risk && risk.cta && risk.cta.width >= THRESHOLD && risk.cta.height >= THRESHOLD;
  const gaAllPass = report.pages.every(p => (p.gaHits || []).length > 0);
  const firebaseResults = FIREBASE_PAGES.map(fp => {
    const pg = report.pages.find(p => p.pathSeg === fp);
    const responses = (pg && pg.firebaseInstallResponses) || [];
    const any200 = responses.some(r => r.status === 200);
    const anyNon200 = responses.some(r => r.status !== 200);
    return { page: fp, responseCount: responses.length, any200, anyNon200, statuses: responses.map(r => r.status) };
  });
  const firebaseAllPass = firebaseResults.every(f => f.responseCount === 0 || (f.any200 && !f.anyNon200));
  // For CSP: new violations = anything NOT vercel.live AND NOT meta-tag delivery warning
  const cspNewViolations = report.pages.flatMap(p =>
    (p.cspViolations || []).filter(v =>
      !/vercel\.live/i.test(v.text) && !/delivered via a <meta>/i.test(v.text)
    ).map(v => ({ page: p.pathSeg, violation: v.text.slice(0, 400) }))
  );
  const cspPass = cspNewViolations.length === 0;

  report.evaluation = {
    ctaPrimaryPass: !!ctaPass,
    ctaPrimary: risk && risk.cta,
    gaAllPagesPass: gaAllPass,
    gaHitCounts: report.pages.map(p => ({ page: p.pathSeg, count: (p.gaHits || []).length })),
    firebaseAllPass,
    firebaseResults,
    cspNewViolationsZero: cspPass,
    cspNewViolations,
    allPass: !!ctaPass && gaAllPass && firebaseAllPass && cspPass,
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(report, null, 2));
  console.log('\n=== EVAL ===');
  console.log('cta-primary PASS:', ctaPass, risk && risk.cta ? `(${risk.cta.width}x${risk.cta.height})` : '');
  console.log('GA all pages PASS:', gaAllPass);
  console.log('Firebase all pass:', firebaseAllPass, JSON.stringify(firebaseResults));
  console.log('CSP new violations zero:', cspPass, 'count=', cspNewViolations.length);
  console.log('ALL PASS:', report.evaluation.allPass);
  console.log('WROTE:', OUT);
  process.exit(report.evaluation.allPass ? 0 : 2);
})();
