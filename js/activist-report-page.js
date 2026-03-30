/* ─── Mobile Nav: handled by nav-handler.js ─── */

/* ─── Progress Bar ─── */
window.addEventListener('scroll', () => {
  const pct = window.scrollY / (document.body.scrollHeight - window.innerHeight) * 100;
  document.getElementById('progress').style.width = pct + '%';
});

/* ─── Color Helper ─── */
const C = {
  navy:  '#1a2d4f', gold:  '#9b8b6e', cream: '#f8f7f5',
  white: '#ffffff', mid:   '#6b7280',
  navyA: (a) => `rgba(26,45,79,${a})`,
  goldA: (a) => `rgba(155,139,110,${a})`,
};
Chart.defaults.font.family = "'Noto Sans JP', sans-serif";
Chart.defaults.color = '#6b7280';

/* ─── Chart 1: Proposals Trend ─── */
new Chart(document.getElementById('proposalChart'), {
  type: 'bar',
  data: {
    labels: ['2020年', '2021年', '2022年', '2023年', '2024年', '2025年'],
    datasets: [
      {
        label: '株主提案のあった会社数',
        data: [45, 55, 65, 80, 113, 114],
        backgroundColor: C.navyA(.75),
        borderColor: C.navy, borderWidth: 1,
      },
      {
        label: 'アクティビスト等機関投資家（提案社数）',
        data: [18, 28, 40, 61, 59, null],
        backgroundColor: C.goldA(.7),
        borderColor: C.gold, borderWidth: 1,
      },
      {
        label: 'アクティビスト等（議案数）',
        data: [null, null, null, null, null, 146],
        backgroundColor: 'rgba(192,57,43,.7)',
        borderColor: '#c0392b', borderWidth: 1,
      }
    ]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'top' }, tooltip: { mode: 'index' } },
    scales: {
      y: { beginAtZero: true, title: { display: true, text: '社数' } }
    }
  }
});

/* ─── Chart 2: Activist Sensitivity Score by Sector ─── */
new Chart(document.getElementById('scoreChart'), {
  type: 'bar',
  data: {
    labels: ['化学・素材', '電力・ガス', '小売・流通', '金融・銀行', 'IT(グロース)', '機械・製造', '総合商社', '不動産'],
    datasets: [{
      label: 'アクティビスト感応度スコア（100点満点）',
      data: [78, 72, 65, 60, 58, 55, 48, 70],
      backgroundColor: [
        '#e74c3c','#e67e22','#f39c12','#3498db',
        '#9b59b6','#95a5a6','#1abc9c','#e74c3c'
      ].map(c => c + 'bb'),
      borderColor: [
        '#e74c3c','#e67e22','#f39c12','#3498db',
        '#9b59b6','#95a5a6','#1abc9c','#e74c3c'
      ],
      borderWidth: 1.5,
    }]
  },
  options: {
    responsive: true, maintainAspectRatio: false, indexAxis: 'y',
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` リスクスコア: ${ctx.parsed.x}点 / 100点`
        }
      }
    },
    scales: {
      x: { beginAtZero: true, max: 100, title: { display: true, text: 'スコア（高いほどリスク大）' }, ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.1)' } },
      y: { ticks: { color: 'rgba(255,255,255,.8)' }, grid: { display: false } }
    }
  }
});

/* ─── Chart 3: Lifecycle ─── */
new Chart(document.getElementById('lifecycleChart'), {
  type: 'line',
  data: {
    labels: ['-12M', '-9M', '-6M', '-3M', '介入公表', '+3M', '+6M', '+9M', '+12M', '+24M', '+36M'],
    datasets: [
      {
        label: '要求実現グループ（対TOPIX超過リターン）',
        data: [2, 3, 4, 6, 18, 20, 21, 22, 24, 32, 41],
        borderColor: '#2ecc71', backgroundColor: 'rgba(46,204,113,.1)',
        tension: .4, fill: true, pointRadius: 4,
      },
      {
        label: '要求未実現グループ（対TOPIX超過リターン）',
        data: [1, 2, 3, 5, 15, 12, 10, 8, 6, 0, -8],
        borderColor: '#e74c3c', backgroundColor: 'rgba(231,76,60,.1)',
        tension: .4, fill: true, pointRadius: 4,
      },
      {
        label: '全体平均',
        data: [1, 2, 3, 5, 17, 16, 15, 15, 15, 16, 17],
        borderColor: C.gold, borderDash: [5,5],
        tension: .4, fill: false, pointRadius: 3,
      }
    ]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top', labels: { color: 'rgba(255,255,255,.8)', boxWidth: 14 } },
      tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: +${ctx.parsed.y}%` } }
    },
    scales: {
      x: { ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.08)' } },
      y: { title: { display: true, text: '対TOPIX超過リターン（%）', color: 'rgba(255,255,255,.6)' }, ticks: { color: 'rgba(255,255,255,.6)', callback: v => v + '%' }, grid: { color: 'rgba(255,255,255,.08)' } }
    }
  }
});

/* ─── Chart 4: Sector Bubble ─── */
new Chart(document.getElementById('sectorChart'), {
  type: 'bubble',
  data: {
    datasets: [
      { label: '化学・素材',       data: [{ x: 82, y: 78, r: 22 }], backgroundColor: 'rgba(231,76,60,.65)' },
      { label: '金融・銀行',       data: [{ x: 60, y: 55, r: 16 }], backgroundColor: 'rgba(52,152,219,.65)' },
      { label: 'インフラ・電力',   data: [{ x: 72, y: 68, r: 18 }], backgroundColor: 'rgba(46,204,113,.65)' },
      { label: '小売・流通',       data: [{ x: 65, y: 72, r: 15 }], backgroundColor: 'rgba(243,156,18,.65)' },
      { label: 'IT・テック(グロース)', data: [{ x: 58, y: 85, r: 12 }], backgroundColor: 'rgba(155,89,182,.65)' },
      { label: '総合商社',         data: [{ x: 48, y: 60, r: 14 }], backgroundColor: 'rgba(26,188,156,.65)' },
      { label: '不動産',           data: [{ x: 70, y: 65, r: 13 }], backgroundColor: 'rgba(230,126,34,.65)' },
      { label: '機械・製造',       data: [{ x: 55, y: 58, r: 17 }], backgroundColor: 'rgba(149,165,166,.65)' },
    ]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => [
            ctx.dataset.label,
            `介入リスク: ${ctx.parsed.x}点`,
            `改革ポテンシャル: ${ctx.parsed.y}点`,
            `キャンペーン件数（バブル比）`
          ]
        }
      }
    },
    scales: {
      x: { min: 30, max: 100, title: { display: true, text: 'アクティビスト介入リスク →', color: 'rgba(255,255,255,.6)' }, ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.08)' } },
      y: { min: 30, max: 100, title: { display: true, text: '介入後の企業価値向上ポテンシャル →', color: 'rgba(255,255,255,.6)' }, ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.08)' } }
    }
  }
});

/* ─── Chart 5: Scatter – Matrix Quadrant ─── */
new Chart(document.getElementById('scatterChart'), {
  type: 'scatter',
  data: {
    datasets: [
      {
        label: '模範型',
        data: [{ x: 85, y: 35 }, { x: 90, y: 28 }, { x: 78, y: 40 }],
        backgroundColor: 'rgba(41,128,185,.8)', pointRadius: 10,
      },
      {
        label: '理想型',
        data: [{ x: 30, y: 55 }, { x: 40, y: 48 }, { x: 25, y: 62 }],
        backgroundColor: 'rgba(39,174,96,.8)', pointRadius: 10,
      },
      {
        label: '対症療法型',
        data: [{ x: 72, y: 72 }, { x: 80, y: 68 }, { x: 65, y: 75 }],
        backgroundColor: 'rgba(230,126,34,.8)', pointRadius: 10,
      },
      {
        label: '危険型',
        data: [{ x: 20, y: 88 }, { x: 30, y: 82 }, { x: 15, y: 92 }],
        backgroundColor: 'rgba(192,57,43,.8)', pointRadius: 10,
      },
    ]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top', labels: { color: 'rgba(255,255,255,.8)', boxWidth: 12 } },
      tooltip: {
        callbacks: {
          label: ctx => `${ctx.dataset.label}｜IR質: ${ctx.parsed.x}点 / リスク: ${ctx.parsed.y}点`
        }
      }
    },
    scales: {
      x: { min: 0, max: 100, title: { display: true, text: '← 対話品質（IRの実質性）', color: 'rgba(255,255,255,.6)' }, ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.08)' } },
      y: { min: 0, max: 100, title: { display: true, text: 'アクティビスト介入リスクスコア →', color: 'rgba(255,255,255,.6)' }, ticks: { color: 'rgba(255,255,255,.6)' }, grid: { color: 'rgba(255,255,255,.08)' } }
    }
  }
});

const GAS_URL = 'https://script.google.com/macros/s/AKfycbwQby968Ts6rSFyzTZZFfblfFDYOogRPWYa7kYl5eHdXaemLIilu6FnOemvX5ygIBwt/exec';

async function submitConsultForm(e) {
  e.preventDefault();
  const form = document.getElementById('consultForm');
  const btn = document.getElementById('consultSubmitBtn');
  const status = document.getElementById('consultStatus');
  const fd = new FormData(form);
  const data = { _formType: 'consultation' };
  fd.forEach((v, k) => { data[k] = v; });
  btn.disabled = true; btn.textContent = '送信中...';
  status.style.display = 'none';
  try {
    await fetch(GAS_URL, { method: 'POST', body: JSON.stringify(data), mode: 'no-cors' });
    status.style.display = 'block'; status.style.background = '#d1fae5'; status.style.color = '#065f46';
    status.textContent = 'ご相談を受け付けました。確認メールをお送りしましたのでご確認ください。';
    form.reset();
  } catch (err) {
    status.style.display = 'block'; status.style.background = '#fee2e2'; status.style.color = '#991b1b';
    status.textContent = '送信に失敗しました。時間をおいて再度お試しください。';
  } finally { btn.disabled = false; btn.textContent = '送信する'; }
  return false;
}

/* ── Event Delegation: form submit ── */
document.addEventListener('submit', function(e) {
  var form = e.target.closest('[data-form-action]');
  if (form && form.dataset.formAction === 'submitConsultForm') {
    submitConsultForm(e);
  }
});
