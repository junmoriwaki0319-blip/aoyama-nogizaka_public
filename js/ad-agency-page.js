(function() {
  'use strict';

  // 企業データはFirestoreから動的取得（認証後にloadPremiumData()で取得）
  let companies = [];

  function populateTable() {
    const tbody = document.getElementById('companiesTableBody');
    tbody.innerHTML = '';
    companies.forEach((company, index) => {
      const row = document.createElement('tr');
      row.innerHTML = `
                    <td>${index + 1}</td>
                    <td><strong>${company.name}</strong></td>
                    <td>${company.ticker}</td>
                    <td>${company.market || '-'}</td>
                    <td>${company.marketCap != null ? company.marketCap.toLocaleString() : '-'}</td>
                    <td>${company.operatingMargin != null ? company.operatingMargin.toFixed(1) + '%' : '-'}</td>
                    <td>${company.roe != null ? company.roe.toFixed(1) + '%' : '-'}</td>
                    <td><span style="padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: 600; background-color: ${company.tier === 'Tier1' ? 'var(--navy)' : company.tier === 'Tier2' ? 'var(--info)' : 'var(--text-tertiary)'}; color: white;">${company.tier}</span></td>
                `;
      tbody.appendChild(row);
    });
  }

  function createMarketCapChart() {
    const ctx = document.getElementById('marketCapChart');
    if (!ctx) return;

    const topCompanies = companies
      .filter(c => c.marketCap != null && c.marketCap > 1500)
      .sort((a, b) => b.marketCap - a.marketCap)
      .slice(0, 30);

    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: topCompanies.map(c => c.name),
        datasets: [{
          label: '時価総額 (億円)',
          data: topCompanies.map(c => c.marketCap),
          backgroundColor: 'var(--navy)',
          borderColor: 'var(--gold)',
          borderWidth: 2,
          borderRadius: 4,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: {
              color: 'var(--text-primary)',
              font: { size: 12 }
            }
          },
          tooltip: {
            backgroundColor: 'var(--bg-white)',
            titleColor: 'var(--navy)',
            bodyColor: 'var(--text-primary)',
            borderColor: 'var(--navy)',
            borderWidth: 1,
            padding: 12,
            displayColors: false,
            callbacks: {
              label: function(ctx) {
                return '時価総額: ¥' + ctx.parsed.x.toLocaleString() + '億';
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { color: 'var(--text-tertiary)' },
            grid: { color: 'var(--border-light)' }
          },
          y: {
            ticks: { color: 'var(--text-primary)', font: { size: 11 } },
            grid: { display: false }
          }
        }
      }
    });
  }

  function setupNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
      item.addEventListener('click', function() {
        navItems.forEach(i => i.classList.remove('active'));
        this.classList.add('active');
      });
    });
  }

  // Firestoreからプレミアムデータを取得してダッシュボード初期化
  window.loadPremiumData = async function() {
    if (!window.firebaseDb) return;
    try {
      const { doc, getDoc } = await import('https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js');
      const db = window.firebaseDb;
      const compSnap = await getDoc(doc(db, 'premiumContent', 'ad-agency-companies'));
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

    // Compute and update KPI values
    var totalMcap = 0, sumMargin = 0, sumRoe = 0, cntMargin = 0, cntRoe = 0;
    companies.forEach(function(c) {
      if (c.marketCap != null) totalMcap += c.marketCap;
      if (c.operatingMargin != null) { sumMargin += c.operatingMargin; cntMargin++; }
      if (c.roe != null) { sumRoe += c.roe; cntRoe++; }
    });
    var kpiMcap = document.getElementById('kpiTotalMcap');
    var kpiMargin = document.getElementById('kpiAvgMargin');
    var kpiRoe = document.getElementById('kpiAvgRoe');
    if (kpiMcap) kpiMcap.textContent = (totalMcap / 10000).toFixed(1);
    if (kpiMargin && cntMargin) kpiMargin.textContent = (sumMargin / cntMargin).toFixed(1);
    if (kpiRoe && cntRoe) kpiRoe.textContent = (sumRoe / cntRoe).toFixed(1);

    populateTable();
    createMarketCapChart();
    setupNavigation();
  };

  // Init nav on load (for non-authenticated state too)
  document.addEventListener('DOMContentLoaded', function() {
    setupNavigation();
  });
})();
