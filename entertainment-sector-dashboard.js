/**
 * エンターテインメント・ゲームセクター ダッシュボード
 * カテゴリフィルター機能
 */
(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    initCategoryFilter();
  });

  /**
   * カテゴリフィルターの初期化
   */
  function initCategoryFilter() {
    var filterButtons = document.querySelectorAll('.filter-btn');
    var tableRows = document.querySelectorAll('#companyTable tbody tr');

    if (!filterButtons.length || !tableRows.length) {
      return;
    }

    filterButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var filter = this.getAttribute('data-filter');

        // アクティブ状態の更新
        filterButtons.forEach(function (b) {
          b.classList.remove('active');
        });
        this.classList.add('active');

        // 行の表示/非表示
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

        // ボタンテキストに件数反映（全体ボタンのみ）
        if (filter === 'all') {
          this.textContent = '全て (' + tableRows.length + ')';
        }
      });
    });

    // 初期表示時に全体カウントを設定
    var allBtn = document.querySelector('[data-filter="all"]');
    if (allBtn) {
      allBtn.textContent = '全て (' + tableRows.length + ')';
    }
  }
})();
