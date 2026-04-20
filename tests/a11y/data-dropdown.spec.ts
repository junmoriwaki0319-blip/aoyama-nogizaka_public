import { test, expect, Page } from '@playwright/test';

// 「データ」ドロップダウン (.nav-dropdown > .nav-dropdown-toggle + .nav-dropdown-menu) の
// キーボード / ARIA アクセシビリティを検証。
// 2026-04-20 時点で実装は CSS :hover のみ / JS なし / aria-expanded なし のため、
// 本テストは現状ほぼ全ケースで失敗することを想定（FIXME マーカーで受け入れ）。

const TOGGLE = '.nav-dropdown .nav-dropdown-toggle';
const MENU = '.nav-dropdown .nav-dropdown-menu';
const MENU_FIRST_LINK = '.nav-dropdown .nav-dropdown-menu li a:first-of-type';

async function gotoHome(page: Page) {
  await page.goto('/');
  await page.waitForLoadState('domcontentloaded');
  // CSS hover を使わずに keyboard で開くことを期待
}

test.describe('データ dropdown a11y', () => {

  test('FIXME: a11y — aria-expanded が click で true/false に切り替わる', async ({ page }) => {
    await gotoHome(page);
    const toggle = page.locator(TOGGLE).first();
    // 初期状態: aria-expanded が存在し false であること
    await expect(toggle).toHaveAttribute('aria-expanded', 'false');

    await toggle.click();
    await expect(toggle).toHaveAttribute('aria-expanded', 'true');

    await toggle.click();
    await expect(toggle).toHaveAttribute('aria-expanded', 'false');
  });

  test('FIXME: a11y — Tab でメニュー項目を順に巡回できる', async ({ page }) => {
    await gotoHome(page);
    const toggle = page.locator(TOGGLE).first();

    await toggle.focus();
    await expect(toggle).toBeFocused();

    // 開く: Enter or Space で開けるはず
    await page.keyboard.press('Enter');
    await expect(page.locator(MENU).first()).toBeVisible();

    // Tab で最初のメニュー項目へ
    await page.keyboard.press('Tab');
    await expect(page.locator(MENU_FIRST_LINK).first()).toBeFocused();
  });

  test('FIXME: a11y — Shift+Tab でトグルへ戻る', async ({ page }) => {
    await gotoHome(page);
    const toggle = page.locator(TOGGLE).first();

    await toggle.focus();
    await page.keyboard.press('Enter');
    await page.keyboard.press('Tab');
    await expect(page.locator(MENU_FIRST_LINK).first()).toBeFocused();

    await page.keyboard.press('Shift+Tab');
    await expect(toggle).toBeFocused();
  });

  test('FIXME: a11y — Esc でメニューが閉じてトグルへフォーカス復帰', async ({ page }) => {
    await gotoHome(page);
    const toggle = page.locator(TOGGLE).first();

    await toggle.focus();
    await page.keyboard.press('Enter');
    await expect(page.locator(MENU).first()).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.locator(MENU).first()).toBeHidden();
    await expect(toggle).toBeFocused();
  });

  test('FIXME: a11y — 外部クリックでメニューが閉じる', async ({ page }) => {
    await gotoHome(page);
    const toggle = page.locator(TOGGLE).first();

    await toggle.click();
    await expect(page.locator(MENU).first()).toBeVisible();

    // nav の外側 (ページ本文) をクリック
    await page.locator('main, body').first().click({ position: { x: 20, y: 400 } });
    await expect(page.locator(MENU).first()).toBeHidden();
  });
});
