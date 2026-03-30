/**
 * モバイルナビゲーションハンドラー
 *
 * 用途: メニューの開閉機能をJavaScriptで管理
 * 対象: ハンバーガーボタン、ナビゲーションリンク
 *
 * 配置先: js/nav-handler.js
 * 読み込み: <script src="js/nav-handler.js"></script> (</body>直前)
 */

(function() {
  'use strict';

  var SELECTORS = {
    navToggle: '[data-nav-toggle]',
    navLink: '[data-nav-link]',
    navMenu: '.mobile-menu',
    activeClass: 'open'
  };

  var menuState = { isOpen: false };

  function toggleMenu(menu) {
    if (!menu) return;
    menuState.isOpen = !menuState.isOpen;
    menu.classList.toggle(SELECTORS.activeClass);
    var toggleBtn = document.querySelector(SELECTORS.navToggle);
    if (toggleBtn) {
      toggleBtn.setAttribute('aria-expanded', String(menuState.isOpen));
    }
  }

  function closeMenu(menu) {
    if (!menu || !menuState.isOpen) return;
    menuState.isOpen = false;
    menu.classList.remove(SELECTORS.activeClass);
    var toggleBtn = document.querySelector(SELECTORS.navToggle);
    if (toggleBtn) {
      toggleBtn.setAttribute('aria-expanded', 'false');
    }
  }

  function init() {
    var menu = document.querySelector(SELECTORS.navMenu);
    var toggleBtn = document.querySelector(SELECTORS.navToggle);
    var navLinks = document.querySelectorAll(SELECTORS.navLink);

    if (!menu || !toggleBtn) return;

    toggleBtn.addEventListener('click', function(e) {
      toggleMenu(menu);
      e.stopPropagation();
    }, false);

    navLinks.forEach(function(link) {
      link.addEventListener('click', function() {
        closeMenu(menu);
      }, false);
    });

    toggleBtn.setAttribute('aria-expanded', 'false');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, false);
  } else {
    init();
  }
})();
