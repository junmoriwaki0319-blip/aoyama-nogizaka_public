/**
 * エンターテインメント・ゲームセクター ダッシュボード
 * Firestoreからデータを動的取得 + Tab navigation + Category filter
 */
(function () {
  'use strict';

  let companies = [];

  const CATEGORY_LABELS = {
    'game-publisher': 'パブリッシャー',
    'mobile-game': 'モバイル',
    'anime-ip': 'アニメ・IP',
    'vtuber-meta': 'VTuber',
    'esports-peripheral': 'eスポーツ',
    'dev-tools': '開発ツール',
  };
  const BADGE_CLASS = {
    'game-publisher': 'badge-publisher',
    'mobile-game': 'badge-mobile',
    'anime-ip': 'badge-anime',
    'vtuber-meta': 'badge-vtuber',
    'esports-peripheral': 'badge-esports',
    'dev-tools': 'badge-tools',
  };

  function populateTable() {
    var tbody = document.querySelector('#companyTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    companies.forEach(function (c) {
      var tr = document.createElement('tr');
      tr.setAttribute('data-category', c.category || '');
      var cap = c.marketCap != null ? c.marketCap.toLocaleString() + '億円' : '—';
      var label = CATEGORY_LABELS[c.category] || c.category;
      var badge = BADGE_CLASS[c.category] || 'badge-publisher';
      tr.innerHTML =
        '<td>' + c.ticker + '</td>' +
        '<td class="company-name">' + c.name + '</td>' +
        '<td><span class="badge ' + badge + '">' + label + '</span></td>' +
        '<td class="cap">' + cap + '</td>' +
        '<td>' + (c.operatingMargin != null ? c.operatingMargin.toFixed(1) + '%' : '—') + '</td>' +
        '<td>' + (c.roe != null ? c.roe.toFixed(1) + '%' : '—') + '</td>';
      tbody.appendChild(tr);
    });
  }

  function updateFilterCounts() {
    var rows = document.querySelectorAll('#companyTable tbody tr');
    var allBtn = document.querySelector('[data-filter="all"]');
    if (allBtn) allBtn.textContent = '全て (' + rows.length + ')';
    // Update per-category counts
    document.querySelectorAll('.filter-btn[data-filter]').forEach(function (btn) {
      var f = btn.getAttribute('data-filter');
      if (f === 'all') return;
      var count = 0;
      rows.forEach(function (r) { if (r.getAttribute('data-category') === f) count++; });
      var label = CATEGORY_LABELS[f] || f;
      btn.textContent = label + ' (' + count + ')';
    });
  }

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
    if (!filterButtons.length) return;

    filterButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var filter = this.getAttribute('data-filter');
        filterButtons.forEach(function (b) { b.classList.remove('active'); });
        this.classList.add('active');

        var tableRows = document.querySelectorAll('#companyTable tbody tr');
        tableRows.forEach(function (row) {
          var category = row.getAttribute('data-category');
          if (filter === 'all' || category === filter) {
            row.classList.remove('hidden');
          } else {
            row.classList.add('hidden');
          }
        });
      });
    });
  }

  // Firestoreからプレミアムデータを取得してダッシュボード初期化
  window.loadPremiumData = async function () {
    if (!window.firebaseDb) return;
    try {
      var { doc, getDoc } = await import('https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js');
      var db = window.firebaseDb;
      var compSnap = await getDoc(doc(db, 'premiumContent', 'entertainment-companies'));
      if (compSnap.exists()) {
        companies = compSnap.data().companies || [];
      }
    } catch (e) {
      console.error('Premium data load failed:', e);
      return;
    }

    // Update company count
    var countEl = document.getElementById('companyCount');
    if (countEl) countEl.textContent = companies.length;

    populateTable();
    updateFilterCounts();
    initCategoryFilter();
  };

  document.addEventListener('DOMContentLoaded', function () {
    initTabs();
    initNavScroll();
  });
})();
