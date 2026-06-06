/**
 * showcase-tabs.js — トップページ「Data & Insights」ショーケースのカテゴリ絞り込み
 * .showcase-tab[data-filter] をクリックで .data-card[data-cat] を表示/非表示。
 */
(function () {
  'use strict';
  function init() {
    var tabs = document.querySelectorAll('.showcase-tab');
    var cards = document.querySelectorAll('.data-grid .data-card');
    if (!tabs.length || !cards.length) return;

    function apply(filter) {
      for (var i = 0; i < cards.length; i++) {
        var cat = cards[i].getAttribute('data-cat');
        var show = (filter === 'all' || cat === filter);
        cards[i].classList.toggle('is-hidden', !show);
      }
    }

    for (var t = 0; t < tabs.length; t++) {
      tabs[t].addEventListener('click', function () {
        for (var j = 0; j < tabs.length; j++) tabs[j].classList.remove('active');
        this.classList.add('active');
        apply(this.getAttribute('data-filter'));
      }, false);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, false);
  } else {
    init();
  }
})();
