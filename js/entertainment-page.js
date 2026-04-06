/**
 * エンターテインメント・ゲームセクター ダッシュボード
 * Tab navigation + Category filter
 */
(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    initTabs();
    initCategoryFilter();
    initNavScroll();
  });

  /* === Tab navigation === */
  function initTabs() {
    var navItems = document.querySelectorAll('.nav-item[data-tab]');
    if (!navItems.length) return;

    navItems.forEach(function (item) {
      item.addEventListener('click', function () {
        var tab = this.getAttribute('data-tab');
        navItems.forEach(function (n) { n.classList.remove('active'); });
        this.classList.add('active');
        document.querySelectorAll('.section').forEach(function (s) { s.classList.remove('active'); });
        var target = document.getElementById('sec-' + tab);
        if (target) target.classList.add('active');
      });
    });
  }

  /* === Nav scroll buttons === */
  function initNavScroll() {
    var navInner = document.getElementById('mainNav');
    var leftBtn = document.getElementById('navScrollLeft');
    var rightBtn = document.getElementById('navScrollRight');
    if (!navInner || !leftBtn || !rightBtn) return;

    function updateButtons() {
      leftBtn.classList.toggle('hidden', navInner.scrollLeft <= 0);
      rightBtn.classList.toggle('hidden', navInner.scrollLeft + navInner.clientWidth >= navInner.scrollWidth - 1);
    }
    leftBtn.addEventListener('click', function () { navInner.scrollBy({ left: -200, behavior: 'smooth' }); });
    rightBtn.addEventListener('click', function () { navInner.scrollBy({ left: 200, behavior: 'smooth' }); });
    navInner.addEventListener('scroll', updateButtons);
    window.addEventListener('resize', updateButtons);
    updateButtons();
  }

  /* === Category filter === */
  function initCategoryFilter() {
    var filterButtons = document.querySelectorAll('.filter-btn');
    var tableRows = document.querySelectorAll('#companyTable tbody tr');

    if (!filterButtons.length || !tableRows.length) return;

    filterButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var filter = this.getAttribute('data-filter');

        filterButtons.forEach(function (b) { b.classList.remove('active'); });
        this.classList.add('active');

        var visibleCount = 0;
        tableRows.forEach(function (row) {
          var category = row.getAttribute('data-category');
          if (filter === 'all' || category === filter) {
            row.classList.remove('hidden');
            visibleCount++;
          } else {
            row.classList.add('hidden');
          }
        });

        if (filter === 'all') {
          this.textContent = '全て (' + tableRows.length + ')';
        }
      });
    });

    var allBtn = document.querySelector('[data-filter="all"]');
    if (allBtn) {
      allBtn.textContent = '全て (' + tableRows.length + ')';
    }
  }
})();
