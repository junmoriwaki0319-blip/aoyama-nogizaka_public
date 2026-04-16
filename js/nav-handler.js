/**
 * モバイルナビゲーションハンドラー v2
 *
 * 機能:
 *   - ハンバーガーメニュー開閉 + X アニメーション
 *   - ドロワーメニュー (transform スライド)
 *   - データサブメニューのアコーディオン展開
 *   - メニュー外タップで閉じる
 *   - オーバーレイ背景
 *
 * 配置先: js/nav-handler.js
 * 読み込み: <script src="/js/nav-handler.js"></script> (</body>直前)
 */

(function() {
  'use strict';

  var SEL = {
    toggle:    '[data-nav-toggle]',
    link:      '[data-nav-link]',
    menu:      '.mobile-menu',
    overlay:   '.mobile-overlay',
    accordion: '[data-accordion-toggle]',
    panel:     '[data-accordion-panel]',
    activeMenu:    'open',
    activeHamburger: 'is-active',
    activeAccordion: 'is-open',
    activeOverlay:   'active'
  };

  var state = { isOpen: false };

  /* --- オーバーレイ生成 --- */
  function ensureOverlay() {
    var overlay = document.querySelector(SEL.overlay);
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.className = 'mobile-overlay';
      var menu = document.querySelector(SEL.menu);
      if (menu && menu.parentNode) {
        menu.parentNode.insertBefore(overlay, menu);
      }
    }
    return overlay;
  }

  /* --- メニュー開閉 --- */
  function openMenu(menu, btn, overlay) {
    state.isOpen = true;
    menu.classList.add(SEL.activeMenu);
    btn.classList.add(SEL.activeHamburger);
    btn.setAttribute('aria-expanded', 'true');
    btn.setAttribute('aria-label', 'メニューを閉じる');
    if (overlay) overlay.classList.add(SEL.activeOverlay);
    document.body.style.overflow = 'hidden';
  }

  function closeMenu(menu, btn, overlay) {
    if (!state.isOpen) return;
    state.isOpen = false;
    menu.classList.remove(SEL.activeMenu);
    btn.classList.remove(SEL.activeHamburger);
    btn.setAttribute('aria-expanded', 'false');
    btn.setAttribute('aria-label', 'メニューを開く');
    if (overlay) overlay.classList.remove(SEL.activeOverlay);
    document.body.style.overflow = '';
  }

  function toggleMenu(menu, btn, overlay) {
    if (state.isOpen) {
      closeMenu(menu, btn, overlay);
    } else {
      openMenu(menu, btn, overlay);
    }
  }

  /* --- アコーディオン --- */
  function toggleAccordion(accordionEl) {
    var isOpen = accordionEl.classList.contains(SEL.activeAccordion);
    accordionEl.classList.toggle(SEL.activeAccordion);
    var toggleBtn = accordionEl.querySelector(SEL.accordion);
    if (toggleBtn) {
      toggleBtn.setAttribute('aria-expanded', String(!isOpen));
    }
  }

  /* --- 初期化 --- */
  function init() {
    var menu = document.querySelector(SEL.menu);
    var btn  = document.querySelector(SEL.toggle);
    if (!menu || !btn) return;

    var overlay = ensureOverlay();

    /* ハンバーガーボタン */
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      toggleMenu(menu, btn, overlay);
    }, false);

    /* ナビリンククリックで閉じる */
    var links = menu.querySelectorAll(SEL.link);
    for (var i = 0; i < links.length; i++) {
      links[i].addEventListener('click', function() {
        closeMenu(menu, btn, overlay);
      }, false);
    }

    /* オーバーレイタップで閉じる */
    if (overlay) {
      overlay.addEventListener('click', function() {
        closeMenu(menu, btn, overlay);
      }, false);
    }

    /* メニュー外タップで閉じる */
    document.addEventListener('click', function(e) {
      if (!state.isOpen) return;
      if (menu.contains(e.target) || btn.contains(e.target)) return;
      closeMenu(menu, btn, overlay);
    }, false);

    /* Escキーで閉じる */
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && state.isOpen) {
        closeMenu(menu, btn, overlay);
        btn.focus();
      }
    }, false);

    /* アコーディオン */
    var accordionToggles = menu.querySelectorAll(SEL.accordion);
    for (var j = 0; j < accordionToggles.length; j++) {
      accordionToggles[j].addEventListener('click', function(e) {
        e.preventDefault();
        var accordion = this.closest('.mobile-accordion');
        if (accordion) toggleAccordion(accordion);
      }, false);
    }

    btn.setAttribute('aria-expanded', 'false');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, false);
  } else {
    init();
  }
})();
