// @ts-check
const { test, expect } = require('@playwright/test');

// ============================================================
// スモークテスト: デプロイ後の基本動作確認
// 実行: npx playwright test tests/smoke.spec.js
// ============================================================

// --- トップページ ---
test.describe('トップページ', () => {
  test('ページが正常に表示される', async ({ page }) => {
    const res = await page.goto('/');
    expect(res.status()).toBe(200);
    await expect(page).toHaveTitle(/青山乃木坂/);
  });

  test('ヒーロー画像(WebP)が読み込まれる', async ({ page }) => {
    await page.goto('/');
    const img = page.locator('img.hero-visual-img');
    await expect(img).toHaveAttribute('src', /hero-tokyo-night\.webp/);
    // 画像が実際にロードされたか
    const naturalWidth = await img.evaluate(el => el.naturalWidth);
    expect(naturalWidth).toBeGreaterThan(0);
  });

  test('会社マーク画像(WebP)が読み込まれる', async ({ page }) => {
    await page.goto('/');
    const img = page.locator('img[alt="青山乃木坂パートナーズ マーク"]');
    await expect(img).toHaveAttribute('src', /company-mark\.webp/);
  });

  test('ハンバーガーメニューが開閉する (モバイル)', async ({ page }) => {
    await page.setViewportSize({ width: 375, height: 812 });
    await page.goto('/');
    // ナビ初期化を待つ（並列実行時の低速ロードでハンドラ未装着になる対策）
    await page.locator('[data-nav-toggle]').waitFor({ state: 'visible' });
    const menu = page.locator('#mobileMenu');
    await expect(menu).not.toHaveClass(/open/);

    await page.locator('[data-nav-toggle]').click();
    await expect(menu).toHaveClass(/open/);

    await page.locator('[data-nav-toggle]').click();
    await expect(menu).not.toHaveClass(/open/);
  });

  test('ナビリンクでメニューが閉じる (モバイル)', async ({ page }) => {
    await page.setViewportSize({ width: 375, height: 812 });
    await page.goto('/');
    await page.locator('[data-nav-toggle]').waitFor({ state: 'visible' });
    await page.locator('[data-nav-toggle]').click();
    await expect(page.locator('#mobileMenu')).toHaveClass(/open/);

    // 直下の可視トップレベルリンクを押す。先頭の [data-nav-link] は
    // 「会社情報」アコーディオン内に畳まれ非表示のため対象から外す。
    await page.locator('#mobileMenu > a[data-nav-link]').first().click();
    await expect(page.locator('#mobileMenu')).not.toHaveClass(/open/);
  });

  test('お問い合わせフォームが method=POST である', async ({ page }) => {
    await page.goto('/');
    const form = page.locator('#contactForm');
    await expect(form).toHaveAttribute('method', /post/i);
  });

  test('nav-handler.js が読み込まれる', async ({ page }) => {
    const jsLoaded = [];
    page.on('response', res => {
      if (res.url().includes('nav-handler.js')) jsLoaded.push(res.status());
    });
    await page.goto('/');
    expect(jsLoaded).toContain(200);
  });
});

// --- 認証モーダル (Dashboard) ---
test.describe('認証モーダル (activist-dashboard)', () => {
  test('ログインモーダルが開閉する', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    const modal = page.locator('#loginModal');
    await expect(modal).not.toBeVisible();

    await page.locator('[data-auth-action="openModal:login"]').first().click();
    await expect(modal).toBeVisible();

    await page.locator('#loginModal [data-auth-action="closeModal:login"]').click();
    await expect(modal).not.toBeVisible();
  });

  test('登録モーダルが開閉する', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    const modal = page.locator('#registerModal');
    await expect(modal).not.toBeVisible();

    await page.locator('[data-auth-action="openModal:register"]').first().click();
    await expect(modal).toBeVisible();
  });

  test('ログイン→登録 モーダル切替', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    await page.locator('[data-auth-action="openModal:login"]').first().click();
    await expect(page.locator('#loginModal')).toBeVisible();

    await page.locator('[data-auth-action="switchModal:register"]').click();
    await expect(page.locator('#loginModal')).not.toBeVisible();
    await expect(page.locator('#registerModal')).toBeVisible();
  });

  test('auth.js が読み込まれる', async ({ page }) => {
    const jsLoaded = [];
    page.on('response', res => {
      if (res.url().includes('/js/auth.js')) jsLoaded.push(res.status());
    });
    await page.goto('/activist-dashboard.html');
    expect(jsLoaded).toContain(200);
  });
});

// --- 認証モーダル (Food Service) ---
test.describe('認証モーダル (food-service)', () => {
  test('ログインモーダルが開く', async ({ page }) => {
    await page.goto('/food-service.html');
    await page.locator('[data-auth-action="openModal:login"]').first().click();
    await expect(page.locator('#loginModal')).toBeVisible();
  });
});

// --- 認証モーダル (SaaS) ---
test.describe('認証モーダル (saas)', () => {
  test('ログインモーダルが開く', async ({ page }) => {
    await page.goto('/saas.html');
    await page.locator('[data-auth-action="openModal:login"]').first().click();
    await expect(page.locator('#loginModal')).toBeVisible();
  });
});

// --- 各ページ HTTP 200 ---
test.describe('全ページ HTTP 200', () => {
  const pages = [
    '/',
    '/activist-dashboard.html',
    '/activist-screener.html',
    '/food-service.html',
    '/saas.html',
    '/risk-assessment.html',
    '/news/',
    '/team.html',
    '/privacy.html',
  ];

  for (const path of pages) {
    test(`${path} が 200 を返す`, async ({ page }) => {
      const res = await page.goto(path);
      expect(res.status()).toBe(200);
    });
  }
});

// --- Activist Screener: フィルター/ソート ---
test.describe('Activist Screener 操作', () => {
  test('ページが表示されJSエラーがない', async ({ page }) => {
    const errors = [];
    page.on('pageerror', err => errors.push(err.message));
    await page.goto('/activist-screener.html');
    await page.waitForTimeout(2000);
    // Firebase auth エラーは無視 (未ログイン状態)
    const criticalErrors = errors.filter(e =>
      !e.includes('auth') && !e.includes('firebase') && !e.includes('Firebase')
    );
    expect(criticalErrors).toHaveLength(0);
  });
});

// --- Activist Dashboard: タブ/フィルター ---
test.describe('Activist Dashboard 操作', () => {
  test('ページが表示されJSエラーがない', async ({ page }) => {
    const errors = [];
    page.on('pageerror', err => errors.push(err.message));
    await page.goto('/activist-dashboard.html');
    await page.waitForTimeout(2000);
    const criticalErrors = errors.filter(e =>
      !e.includes('auth') && !e.includes('firebase') && !e.includes('Firebase')
    );
    expect(criticalErrors).toHaveLength(0);
  });
});

// --- News ページ ---
test.describe('News ページ', () => {
  test('ページが表示されJSエラーがない', async ({ page }) => {
    const errors = [];
    page.on('pageerror', err => errors.push(err.message));
    await page.goto('/news/');
    await page.waitForTimeout(2000);
    const criticalErrors = errors.filter(e =>
      !e.includes('auth') && !e.includes('firebase') && !e.includes('Firebase')
    );
    expect(criticalErrors).toHaveLength(0);
  });
});

// --- inline handler ゼロ確認 ---
test.describe('inline handler 完全除去確認', () => {
  const pages = [
    '/',
    '/activist-dashboard.html',
    '/activist-screener.html',
    '/food-service.html',
    '/saas.html',
    '/risk-assessment.html',
    '/news/',
  ];

  for (const path of pages) {
    test(`${path} に onclick/onsubmit/onchange がない`, async ({ page }) => {
      await page.goto(path);
      const inlineHandlers = await page.evaluate(() => {
        const all = document.querySelectorAll('[onclick],[onsubmit],[onchange]');
        return all.length;
      });
      expect(inlineHandlers).toBe(0);
    });
  }
});
