// @ts-check
const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './tests',
  timeout: 30000,
  retries: 0,
  use: {
    baseURL: 'https://aoyama-nogizaka.com',
    headless: true,
    screenshot: 'only-on-failure',
    trace: 'on-first-retry',
  },
  projects: [
    // 1. 認証セットアップ（TEST_EMAIL/TEST_PASSWORD 設定時のみ有効）
    {
      name: 'auth-setup',
      testMatch: 'auth.setup.js',
    },
    // 2. スモークテスト（未ログイン）
    {
      name: 'smoke',
      testMatch: 'smoke.spec.js',
      use: { browserName: 'chromium' },
    },
    // 3. 認証済みテスト（auth-setup 完了後に実行）
    {
      name: 'authenticated',
      testMatch: 'authenticated.spec.js',
      dependencies: ['auth-setup'],
      use: { browserName: 'chromium' },
    },
  ],
});
