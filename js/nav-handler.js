/**
 * モバイルナビゲーションハンドラー
 *
 * 用途: メニューの開閉機能、データセクションのアコーディオンをJSで管理
 * 対象: ハンバーガーボタン、ナビゲーションリンク、データドロップダウン
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

  /* ─── データセクション アコーディオン ─── */

  function injectAccordionStyles() {
    var style = document.createElement('style');
    style.textContent = [
      '.mob-data-toggle{color:rgba(255,255,255,0.8);font-size:1rem;letter-spacing:0.06em;padding:18px 0;border-bottom:1px solid rgba(255,255,255,0.1);cursor:pointer;display:flex;justify-content:space-between;align-items:center}',
      '.mob-arrow{font-size:10px;color:rgba(200,185,154,0.7);transition:transform 0.3s}',
      '.mob-data-toggle.active .mob-arrow{transform:rotate(180deg)}',
      '.mob-data-group{display:none;flex-direction:column;background:rgba(0,0,0,0.12);border-radius:4px;margin-bottom:4px;overflow:hidden}',
      '.mob-data-group.open{display:flex}',
      '.mob-data-cat{padding:14px 16px 6px;font-size:0.68rem;color:rgba(200,185,154,0.8);letter-spacing:0.1em;font-weight:600}',
      '.mob-data-cat:not(:first-child){border-top:1px solid rgba(255,255,255,0.06);padding-top:14px}',
      '.mobile-menu .mob-data-link{text-decoration:none;color:rgba(255,255,255,0.6);font-size:0.82rem;padding:9px 16px 9px 24px;border-bottom:none;transition:color 0.2s;letter-spacing:0.03em}',
      '.mobile-menu .mob-data-link:last-child{padding-bottom:14px}',
      '.mobile-menu .mob-data-link:hover,.mobile-menu .mob-data-link.active{color:rgba(200,185,154,0.9)}'
    ].join('\n');
    document.head.appendChild(style);
  }

  function initDataAccordion(menu) {
    var children = Array.from(menu.children);
    var dataItems = [];

    children.forEach(function(el) {
      var style = el.getAttribute('style') || '';
      var isCategory = style.indexOf('gold-light') > -1 && el.tagName === 'DIV';
      var isSubLink = style.indexOf('padding-left') > -1 && el.tagName === 'A';
      if (isCategory || isSubLink) {
        dataItems.push(el);
      }
    });

    if (dataItems.length === 0) return;

    injectAccordionStyles();

    // Create toggle
    var toggle = document.createElement('div');
    toggle.className = 'mob-data-toggle';
    toggle.setAttribute('role', 'button');
    toggle.setAttribute('aria-expanded', 'false');
    toggle.innerHTML = '\u30C7\u30FC\u30BF <span class="mob-arrow">\u25BC</span>';

    // Create group container
    var group = document.createElement('div');
    group.className = 'mob-data-group';

    // Insert toggle before first data item
    menu.insertBefore(toggle, dataItems[0]);
    // Insert group after toggle
    menu.insertBefore(group, toggle.nextSibling);

    // Move data items into group and reclassify
    dataItems.forEach(function(item) {
      if (item.tagName === 'DIV') {
        item.className = 'mob-data-cat';
      } else {
        item.classList.add('mob-data-link');
      }
      item.removeAttribute('style');
      group.appendChild(item);
    });

    // Toggle click
    toggle.addEventListener('click', function(e) {
      var isOpen = group.classList.contains('open');
      toggle.classList.toggle('active');
      group.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(!isOpen));
      e.stopPropagation();
    });
  }

  /* ─── 初期化 ─── */

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

    // データセクションをアコーディオン化
    initDataAccordion(menu);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, false);
  } else {
    init();
  }
})();
