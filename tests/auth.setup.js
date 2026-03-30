// @ts-check
/**
 * 認証セットアップ: Firebase Auth でログインし、storageState を保存
 *
 * 使い方:
 *   1. .env に TEST_EMAIL / TEST_PASSWORD を設定
 *   2. npx playwright test tests/auth.setup.js  (単独実行で認証状態を保存)
 *   3. 他のテストは storageState を読み込んでログイン済み状態で実行
 *
 * ログイン済みテスト実行:
 *   npx playwright test tests/authenticated.spec.js
 */
const { test, expect } = require('@playwright/test');
const path = require('path');

const STORAGE_STATE_PATH = path.join(__dirname, '.auth', 'user.json');

test('Firebase Auth でログインし storageState を保存', async ({ page }) => {
  const email = process.env.TEST_EMAIL;
  const password = process.env.TEST_PASSWORD;

  if (!email || !password) {
    console.log('TEST_EMAIL / TEST_PASSWORD が未設定のためスキップ');
    console.log('設定例: TEST_EMAIL=test@example.com TEST_PASSWORD=xxx npx playwright test tests/auth.setup.js');
    test.skip();
    return;
  }

  // Dashboard ページへ移動（認証モーダルあり）
  await page.goto('/activist-dashboard.html');

  // ログインモーダルを開く
  await page.locator('[data-auth-action="openModal:login"]').first().click();
  await expect(page.locator('#loginModal')).toBeVisible();

  // メールアドレス・パスワードを入力
  await page.locator('#loginEmail').fill(email);
  await page.locator('#loginPassword').fill(password);

  // 送信
  await page.locator('#loginModal form [type="submit"]').click();

  // ログイン完了を待つ（ナビにユーザー名が表示される）
  await expect(page.locator('#navUserInfo')).toBeVisible({ timeout: 15000 });

  // storageState を保存
  const context = page.context();
  await context.storageState({ path: STORAGE_STATE_PATH });

  console.log(`storageState saved to: ${STORAGE_STATE_PATH}`);
});
