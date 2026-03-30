// @ts-check
/**
 * 認証済みテスト: ログイン後のUI・データ表示を確認
 *
 * 前提: auth.setup.js で storageState が保存済みであること
 * 実行: npx playwright test tests/authenticated.spec.js
 */
const { test, expect } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const STORAGE_STATE_PATH = path.join(__dirname, '.auth', 'user.json');

// storageState が存在しなければ全テストスキップ
const hasAuth = fs.existsSync(STORAGE_STATE_PATH);

test.describe('認証済み: Activist Dashboard', () => {
  test.skip(!hasAuth, 'storageState 未作成 — 先に auth.setup.js を実行してください');
  if (hasAuth) test.use({ storageState: STORAGE_STATE_PATH });

  test('ログイン後にユーザー名が表示される', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    await expect(page.locator('#navUserInfo')).toBeVisible({ timeout: 15000 });
    const name = await page.locator('#navUserName').textContent();
    expect(name.trim().length).toBeGreaterThan(0);
  });

  test('プレミアムオーバーレイが非表示', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    await page.waitForTimeout(3000);
    const overlay = page.locator('.premium-overlay, .premium-wall, #premiumOverlay');
    if (await overlay.count() > 0) {
      await expect(overlay.first()).not.toBeVisible();
    }
  });

  test('ログアウトボタンが動作する', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    await expect(page.locator('#navUserInfo')).toBeVisible({ timeout: 15000 });

    await page.locator('[data-auth-action="handleLogout"]').first().click();
    await expect(page.locator('#navAuthBtns')).toBeVisible({ timeout: 10000 });
  });
});

test.describe('認証済み: Food Service', () => {
  test.skip(!hasAuth, 'storageState 未作成');
  if (hasAuth) test.use({ storageState: STORAGE_STATE_PATH });

  test('ダッシュボードコンテンツが表示される', async ({ page }) => {
    await page.goto('/food-service.html');
    await page.waitForTimeout(3000);
    const content = page.locator('#dashboardContent, .dashboard-content');
    if (await content.count() > 0) {
      const filter = await content.first().evaluate(el => getComputedStyle(el).filter);
      expect(filter).not.toContain('blur');
    }
  });
});

test.describe('認証済み: SaaS', () => {
  test.skip(!hasAuth, 'storageState 未作成');
  if (hasAuth) test.use({ storageState: STORAGE_STATE_PATH });

  test('ダッシュボードコンテンツが表示される', async ({ page }) => {
    await page.goto('/saas.html');
    await page.waitForTimeout(3000);
    const content = page.locator('#dashboardContent, .dashboard-content');
    if (await content.count() > 0) {
      const filter = await content.first().evaluate(el => getComputedStyle(el).filter);
      expect(filter).not.toContain('blur');
    }
  });
});

test.describe('認証済み: マイページ', () => {
  test.skip(!hasAuth, 'storageState 未作成');
  if (hasAuth) test.use({ storageState: STORAGE_STATE_PATH });

  test('マイページリンクが開ける', async ({ page }) => {
    await page.goto('/activist-dashboard.html');
    await expect(page.locator('#navUserInfo')).toBeVisible({ timeout: 15000 });

    await page.locator('[data-auth-action="openMyPage"]').first().click();
    await expect(page.locator('#myPageModal, #myPage, .mypage-overlay')).toBeVisible({ timeout: 5000 });
  });
});
