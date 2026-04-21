#!/usr/bin/env node
// Capture full-page screenshots for 13 URLs × 3 viewports.
// Output: audits/20260421/mobile-screenshots/<slug>_<w>x<h>.png

import { chromium } from '@playwright/test';
import fs from 'node:fs';

const OUT_DIR = './audits/20260421/mobile-screenshots';
fs.mkdirSync(OUT_DIR, { recursive: true });

const URLS = [
  ['home', 'https://aoyama-nogizaka.com/'],
  ['team', 'https://aoyama-nogizaka.com/team'],
  ['news', 'https://aoyama-nogizaka.com/news/'],
  ['privacy', 'https://aoyama-nogizaka.com/privacy'],
  ['activist-dashboard', 'https://aoyama-nogizaka.com/activist-dashboard.html'],
  ['risk-assessment', 'https://aoyama-nogizaka.com/risk-assessment.html'],
  ['activist-screener', 'https://aoyama-nogizaka.com/activist-screener.html'],
  ['food-service', 'https://aoyama-nogizaka.com/food-service.html'],
  ['saas', 'https://aoyama-nogizaka.com/saas.html'],
  ['ad-agency', 'https://aoyama-nogizaka.com/ad-agency.html'],
  ['digital-media', 'https://aoyama-nogizaka.com/digital-media.html'],
  ['entertainment-sector-dashboard', 'https://aoyama-nogizaka.com/entertainment-sector-dashboard.html'],
  ['news-activist-shareholder-proposals-japan', 'https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html'],
];

const VIEWPORTS = [
  { w: 375, h: 667, label: '375' },
  { w: 414, h: 896, label: '414' },
  { w: 768, h: 1024, label: '768' },
];

const browser = await chromium.launch({ headless: true });
const log = [];

try {
  for (const vp of VIEWPORTS) {
    const context = await browser.newContext({
      viewport: { width: vp.w, height: vp.h },
      deviceScaleFactor: 2,
      isMobile: vp.w < 768,
      hasTouch: true,
    });
    for (const [slug, url] of URLS) {
      const page = await context.newPage();
      const out = `${OUT_DIR}/${slug}_${vp.label}.png`;
      const t0 = Date.now();
      try {
        await page.goto(url, { waitUntil: 'networkidle', timeout: 45000 });
      } catch (e) {
        // networkidle が届かなくても DOM ready 時点で進める
        try { await page.waitForLoadState('domcontentloaded', { timeout: 5000 }); } catch {}
        log.push(`WARN ${slug}_${vp.label}: goto timeout (${e.message})`);
      }
      // CLS が落ち着くのを待つ
      await page.waitForTimeout(1500);
      try {
        await page.screenshot({ path: out, fullPage: true, timeout: 30000 });
        const st = fs.statSync(out);
        const ms = Date.now() - t0;
        log.push(`OK   ${slug}_${vp.label} (${(st.size/1024).toFixed(0)}KB, ${ms}ms)`);
      } catch (e) {
        log.push(`FAIL ${slug}_${vp.label}: ${e.message}`);
      }
      await page.close();
    }
    await context.close();
  }
} finally {
  await browser.close();
}

fs.writeFileSync(`${OUT_DIR}/_run.log`, log.join('\n') + '\n');
console.log(log.join('\n'));
console.log(`\nDone. ${log.filter(l => l.startsWith('OK')).length} screenshots, ${log.filter(l => l.startsWith('FAIL')).length} failed.`);
