// ===== GLOBAL DATA =====
let reportData = { reports: [], activists: {} };
let filteredReports = [];
let dataLoaded = false;
const PREMIUM_API = 'https://aoyama-nogizakapublic.vercel.app/api/premium-reports';
const DATA_URL = '/data/reports.json';
// ===== INIT =====
document.addEventListener('DOMContentLoaded', async () => {
  await loadData();
  renderAll();
});
async function loadData() {
  if (dataLoaded) return;
  // まず認証付きAPIを試行、失敗したら直接ファイル取得にフォールバック
  let loaded = false;
  const user = window.currentUser;
  if (user) {
    try {
      const idToken = await user.getIdToken();
      const resp = await fetch(PREMIUM_API + '?type=reports', {
        headers: { 'Authorization': 'Bearer ' + idToken }
      });
      if (resp.ok) {
        reportData = await resp.json();
        loaded = true;
      }
    } catch (e) { /* フォールバックへ */ }
  }
  // フォールバック: 直接ファイル取得
  if (!loaded) {
    try {
      const cacheBuster = '?t=' + Date.now();
      const resp = await fetch(DATA_URL + cacheBuster, { cache: 'no-store' });
      if (resp.ok) {
        reportData = await resp.json();
      }
    } catch (e) {
      console.log('Data not yet available, using empty dataset');
    }
  }
  // reports が空なら空配列を保証
  if (!reportData.reports) reportData.reports = [];
  if (!reportData.activists) reportData.activists = {};
  // known_activists.json を読み込み
  try {
    let kaResp;
    if (user) {
      try {
        const idToken = await user.getIdToken();
        kaResp = await fetch(PREMIUM_API + '?type=activists', {
          headers: { 'Authorization': 'Bearer ' + idToken }
        });
      } catch (e) { /* フォールバックへ */ }
    }
    if (!kaResp || !kaResp.ok) {
      kaResp = await fetch('/scripts/known_activists.json?t=' + Date.now(), { cache: 'no-store' });
    }
    if (kaResp.ok) {
      const kaData = await kaResp.json();
      const groups = kaData.groups || {};
      const activistList = kaData.activists || [];
      // グループに属する個別IDをグループIDに名寄せ（Python側で既にマージ済みだが、0件エントリ用）
      const idToGroup = {};
      activistList.forEach(ka => {
        if (ka.group_id && groups[ka.group_id]) {
          idToGroup[ka.id] = ka.group_id;
        }
      });
      // グループエントリを追加（0件でも表示）
      Object.entries(groups).forEach(([gid, g]) => {
        if (!reportData.activists[gid]) {
          reportData.activists[gid] = {
            id: gid,
            name: g.name || gid,
            type: g.type || 'activist',
            representative: g.representative || '',
            headquarters: '',
            description: g.description || '',
            focus_sectors: [],
            holdings: [],
            report_count: 0,
            latest_date: '',
            member_ids: activistList.filter(a => a.group_id === gid).map(a => a.id)
          };
        }
      });
      // グループに属さない個別投資家を追加
      activistList.forEach(ka => {
        if (!idToGroup[ka.id] && !reportData.activists[ka.id]) {
          reportData.activists[ka.id] = {
            id: ka.id,
            name: ka.name,
            type: ka.type || 'activist',
            representative: ka.representative || '',
            headquarters: ka.headquarters || '',
            description: ka.description || '',
            focus_sectors: ka.focus_sectors || [],
            holdings: [],
            report_count: 0,
            latest_date: ''
          };
        }
      });
      // グループに属する個別エントリを除外（二重カウント防止）
      Object.keys(idToGroup).forEach(id => {
        if (reportData.activists[id]) {
          delete reportData.activists[id];
        }
      });
    }
  } catch (e) {
    console.log('known_activists.json not available');
  }
  // 全レポートの比率変動（デルタ）を事前計算
  const sortedForDelta = [...reportData.reports].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const prevRatioMap = {};
  sortedForDelta.forEach(r => {
    const key = (r.filer_name || '') + '|' + (r.sec_code || '');
    const cur = parseFloat(r.holding_ratio);
    if (prevRatioMap[key] !== undefined && !isNaN(cur)) {
      r._delta = cur - prevRatioMap[key];
    } else {
      r._delta = null;
    }
    if (!isNaN(cur)) prevRatioMap[key] = cur;
  });
  filteredReports = [...reportData.reports];
  dataLoaded = true;
}
function renderAll() {
  updateStats();
  renderReportTable();
  renderActivistCards();
  renderActivistRanking();
  renderTimeline();
  updateLastUpdated();
}
// ===== STATS =====
function updateStats() {
  const now = new Date();
  const thisMonth = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  const monthReports = reportData.reports.filter(r => r.date && r.date.startsWith(thisMonth));
  const activistIds = new Set();
  reportData.reports.forEach(r => { if (r.activist_id) activistIds.add(r.activist_id); });
  document.getElementById('statTotal').textContent = reportData.total_reports || reportData.reports.length;
  document.getElementById('statActivist').textContent = (reportData.activist_reports || 0) + reportData.reports.filter(r => r.is_notable).length;
  document.getElementById('statThisMonth').textContent = monthReports.length;
  document.getElementById('statTracked').textContent = Object.keys(reportData.activists).length;
}
function updateLastUpdated() {
  const el = document.getElementById('lastUpdated');
  if (reportData.last_updated) {
    const d = new Date(reportData.last_updated);
    el.textContent = d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  } else {
    el.textContent = 'データ未取得（EDINET APIキー設定後に自動更新）';
  }
}
// ===== REPORT TABLE =====
function renderReportTable() {
  const tbody = document.getElementById('reportTableBody');
  const reports = filteredReports;
  if (!reports.length) {
    tbody.innerHTML = '<tr><td colspan="8" class="no-data">条件に合う報告がありません。EDINET APIキー設定後にデータが表示されます。</td></tr>';
    return;
  }
  tbody.innerHTML = reports.slice(0, 50).map(r => {
    const isNew = isRecentReport(r.date, 3);
    const typeClass = r.report_type === '新規報告' ? 'type-new' : 'type-change';
    const ratioNum = r.holding_ratio != null ? parseFloat(r.holding_ratio) : NaN;
    const ratio = !isNaN(ratioNum) ? ratioNum.toFixed(1) + '%' : '—';
    const barWidth = !isNaN(ratioNum) ? Math.min(ratioNum * 2, 100) : 0;
    const deltaHtml = r._delta != null ? (() => {
      const d = r._delta;
      const sign = d > 0 ? '+' : '';
      const color = d > 0 ? '#c0392b' : d < 0 ? '#2980b9' : 'var(--text-light)';
      return `<span style="font-size:10px;color:${color};margin-left:4px;">(${sign}${d.toFixed(1)})</span>`;
    })() : '';
    const edinetUrl = r.doc_id ? `https://disclosure2.edinet-fsa.go.jp/WZEK0040.aspx?${encodeURIComponent(r.doc_id)}` : '#';
    return `<tr data-action="searchByCode" data-arg="${escapeJs(r.sec_code || '')}" style="cursor:pointer">
      <td class="date-cell col-date">${escapeHtml(r.date || '—')}${isNew ? '<span class="badge-new">NEW</span>' : ''}</td>
      <td class="col-filer" title="${escapeHtml(r.filer_name || '')}"><span class="activist-name">${escapeHtml(r.filer_name || '—')}</span>${r.activist_type === 'individual_investor' ? ' <span class="activist-type" style="background:#f0e6ff;color:#7c3aed;">個人</span>' : r.is_activist ? ' <span class="activist-type type-activist">ACT</span>' : r.is_notable ? ' <span class="activist-type type-notable">注目</span>' : ''}</td>
      <td class="col-type"><span class="activist-type ${typeClass}">${escapeHtml(r.report_type || '—')}</span></td>
      <td class="col-code">${r.sec_code ? '<span class="sec-code">' + escapeHtml(r.sec_code) + '</span>' : '—'}</td>
      <td class="col-target" title="${escapeHtml(r.target_company || '')}">${escapeHtml(r.target_company || '—')}</td>
      <td class="col-ratio">${ratio}${deltaHtml}${barWidth ? '<div class="holding-bar"><div class="holding-bar-fill" style="width:' + barWidth + '%"></div></div>' : ''}</td>
      <td class="col-purpose" style="font-size:12px;">${escapeHtml(r.purpose || '—')}</td>
      <td class="col-edinet"><a href="${edinetUrl}" target="_blank" rel="noopener" class="edinet-link" data-action="stopPropagation">原文</a></td>
    </tr>`;
  }).join('');
}
// ===== ACTIVIST CARDS =====
function renderActivistCards() {
  const container = document.getElementById('activistCards');
  const activists = reportData.activists;
  if (!Object.keys(activists).length) {
    container.innerHTML = '<div class="no-data">EDINET APIキー設定後にアクティビスト情報が表示されます。</div>';
    return;
  }
  let html = '';
  let count = 0;
  for (const [id, data] of Object.entries(activists)) {
    count++;
    const isPremium = count > 3;
    // 保有比率の高い上位3銘柄を表示（比率ありを優先）
    const sortedHoldings = [...(data.holdings || [])].sort((a, b) => (b.holding_ratio || 0) - (a.holding_ratio || 0));
    const topHoldings = sortedHoldings.slice(0, 3);
    const holdingsHtml = topHoldings.map(h => {
      const genericNames = ['変更報告書', '大量保有報告書', '訂正報告書', '変更報告書（特例対象株券等）', '大量保有報告書（特例対象株券等）', '訂正報告書（', ''];
      const name = h.target_company && !genericNames.some(g => h.target_company === g || h.target_company.startsWith('訂正報告書'))
        ? h.target_company : (h.sec_code ? h.sec_code : '—');
      const codeLabel = h.sec_code ? '<span class="sec-code">' + h.sec_code + '</span> ' : '';
      const hRatio = h.holding_ratio != null ? parseFloat(h.holding_ratio) : NaN;
      const ratio = !isNaN(hRatio) ? hRatio.toFixed(1) + '%' : '—';
      const typeTag = h.report_type === '新規報告' ? '<span class="badge-new" style="margin-left:6px;">新規</span>' : '';
      return `<div class="ac-stock-item" style="padding:6px 0; border-bottom:1px solid var(--light-gray);">
        <span style="flex:1;">${codeLabel}${escapeHtml(name)}${typeTag}</span>
        <span class="ac-stock-pct" style="min-width:50px; text-align:right;">${ratio}</span>
      </div>`;
    }).join('');
    const moreCount = sortedHoldings.length - 3;
    const moreHtml = moreCount > 0 ? `<div style="font-size:11px; color:var(--text-light); margin-top:6px;">他 ${moreCount} 銘柄</div>` : '';
    const cardHtml = `
      <div class="activist-card">
        <div class="ac-header">
          <div>
            <div class="ac-name" style="cursor:pointer;text-decoration:underline;text-decoration-color:var(--mid-gray);text-underline-offset:2px;" data-action="openInvestorDetail" data-arg="${escapeJs(id)}">${escapeHtml(data.name)}</div>
            <div class="ac-meta">
              ${data.representative ? '代表: ' + escapeHtml(data.representative) + ' ｜ ' : ''}
              ${data.headquarters ? '本拠: ' + escapeHtml(data.headquarters) + ' ｜ ' : ''}
              報告件数: ${data.report_count || 0}件
            </div>
          </div>
          <span class="activist-type ${data.type === 'activist' ? 'type-activist' : data.type === 'individual_investor' ? 'type-individual' : data.type === 'notable_holder' ? 'type-notable' : 'type-fund'}">${data.type === 'activist' ? 'アクティビスト' : data.type === 'individual_investor' ? '個人注目投資家' : data.type === 'notable_holder' ? '注目投資家' : 'ファンド'}</span>
        </div>
        <div class="ac-grid">
          <div>
            <div class="ac-section-title">保有銘柄（大量保有報告ベース）</div>
            <div class="ac-stock-list">${holdingsHtml || '<div style="font-size:13px; color:var(--text-light);">データなし</div>'}${moreHtml}</div>
          </div>
          <div>
            <div class="ac-section-title">活動傾向</div>
            <div style="font-size:13px; color:var(--text-mid); line-height:1.8;">${escapeHtml(data.description || '—')}</div>
            ${data.focus_sectors && data.focus_sectors.length ?
              '<div style="margin-top:8px; display:flex; gap:4px; flex-wrap:wrap;">' +
              data.focus_sectors.map(s => '<span class="timeline-tag">' + escapeHtml(s) + '</span>').join('') +
              '</div>' : ''}
          </div>
        </div>
      </div>`;
    if (isPremium && count === 4) {
      html += `<div class="premium-content">
        <div class="premium-blur">${cardHtml}</div>
        <div class="premium-overlay"><div class="premium-badge"><p>この先は会員限定コンテンツです</p><button data-auth-action="openModal:register">無料会員登録して続きを見る</button></div></div>
      </div>`;
      break;
    } else {
      html += cardHtml;
    }
  }
  container.innerHTML = html;
}
// ===== ACTIVIST RANKING =====
let rankingTypeFilter = 'all';
function setRankingFilter(type) {
  rankingTypeFilter = type;
  document.querySelectorAll('.ranking-filter-btn').forEach(btn => {
    if (btn.dataset.filter === type) {
      btn.style.background = 'var(--navy)';
      btn.style.color = '#fff';
    } else {
      btn.style.background = '#fff';
      btn.style.color = 'var(--text-mid)';
    }
  });
  filterRanking();
}
function filterRanking() {
  const query = (document.getElementById('rankingSearch').value || '').toLowerCase();
  const container = document.getElementById('activistRankingList');
  const activists = reportData.activists;
  if (!Object.keys(activists).length) {
    container.innerHTML = '<div style="font-size:12px; color:var(--text-light);">データ取得後に表示</div>';
    return;
  }
  let sorted = Object.values(activists).sort((a, b) => (b.report_count || 0) - (a.report_count || 0));
  // タイプフィルター
  if (rankingTypeFilter !== 'all') {
    sorted = sorted.filter(a => a.type === rankingTypeFilter);
  }
  // テキスト検索
  if (query) {
    sorted = sorted.filter(a => (a.name || '').toLowerCase().includes(query) || (a.id || '').toLowerCase().includes(query));
  }
  sorted = sorted.slice(0, 50);
  if (!sorted.length) {
    container.innerHTML = '<div style="font-size:12px; color:var(--text-light);">該当する投資家がありません</div>';
    return;
  }
  container.innerHTML = sorted.map((a, idx) => {
    const typeTag = a.type === 'activist' ? '<span class="activist-type type-activist" style="font-size:11px;padding:1px 4px;margin-left:4px;">ACT</span>'
      : a.type === 'individual_investor' ? '<span class="activist-type" style="font-size:11px;padding:1px 4px;margin-left:4px;background:#f0e6ff;color:#7c3aed;border-radius:4px;">個人</span>'
      : a.type === 'notable_holder' ? '<span class="activist-type type-notable" style="font-size:11px;padding:1px 4px;margin-left:4px;">注目</span>'
      : '';
    const rank = idx + 1;
    const rankStyle = rank <= 3 ? 'font-weight:bold;color:var(--gold);' : 'color:var(--text-light);';
    const isGroup = a.member_ids && a.member_ids.length;
    const groupIcon = isGroup ? '<span title="グループ名寄せ" style="font-size:11px;color:var(--text-light);margin-left:2px;">&#9679;</span>' : '';
    return `<div class="ac-stock-item" style="cursor:pointer;display:flex;align-items:center;" data-action="openInvestorDetail" data-arg="${escapeJs(a.id)}">
      <span style="min-width:22px;${rankStyle}">${rank}</span>
      <span style="flex:1;">${escapeHtml(a.name)}${groupIcon}${typeTag}</span>
      <span style="color:${(a.report_count || 0) >= 5 ? 'var(--red)' : 'var(--text-mid)'}; margin-right:6px;">${a.report_count || 0}件</span>
      <span title="報告一覧を絞り込み" style="font-size:10px;color:var(--text-light);cursor:pointer;" data-action="filterByActivist" data-arg="${escapeJs(a.id)}">&#9654;</span>
    </div>`;
  }).join('');
}
function renderActivistRanking() {
  filterRanking();
}
// ===== TIMELINE =====
function renderTimeline() {
  const container = document.getElementById('timelineContainer');
  const activistReports = reportData.reports.filter(r => r.is_activist || r.is_notable).slice(0, 10);
  if (!activistReports.length) {
    container.innerHTML = '<div class="no-data">注目投資家の報告がまだありません。</div>';
    return;
  }
  container.innerHTML = activistReports.map(r =>
    `<div class="timeline-item">
      <div class="timeline-date">${escapeHtml(r.date || '—')}</div>
      <div class="timeline-content">
        <div class="timeline-title">${escapeHtml(r.filer_name || '')} ${r.sec_code ? '(' + escapeHtml(r.sec_code) + ')' : ''}</div>
        <div class="timeline-detail">${escapeHtml(r.target_company || '')}${r.holding_ratio != null && !isNaN(parseFloat(r.holding_ratio)) ? ' — 保有比率: ' + parseFloat(r.holding_ratio).toFixed(1) + '%' : ''}${r.purpose ? ' — 目的: ' + escapeHtml(r.purpose) : ''}</div>
        <div class="timeline-tags">
          <span class="timeline-tag">${r.report_type || '報告'}</span>
          ${r.is_activist ? '<span class="timeline-tag">アクティビスト</span>' : ''}
        </div>
      </div>
    </div>`
  ).join('');
  // Show login wall after timeline
  document.getElementById('loginWall').style.display = 'block';
}
// ===== FILTERS =====
function applyFilters() {
  const keyword = document.getElementById('searchKeyword').value.trim().toLowerCase();
  const investorFilter = document.getElementById('investorFilter').value;
  const period = parseInt(document.getElementById('periodFilter').value);
  const sortBy = document.getElementById('sortSelect').value;
  const reportTypes = [];
  document.querySelectorAll('input[name="reportType"]:checked').forEach(cb => reportTypes.push(cb.value));
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - period);
  const cutoffStr = cutoffDate.toISOString().slice(0, 10);
  filteredReports = reportData.reports.filter(r => {
    if (keyword) {
      const text = [r.filer_name, r.sec_code, r.target_company, r.purpose].join(' ').toLowerCase();
      if (!text.includes(keyword)) return false;
    }
    if (investorFilter === 'activist' && !r.is_activist) return false;
    if (investorFilter === 'notable' && !r.is_notable) return false;
    if (investorFilter === 'both' && !r.is_activist && !r.is_notable) return false;
    if (r.date && r.date < cutoffStr) return false;
    if (reportTypes.length) {
      const typeMap = { 'new': '新規報告', 'change': '変更報告' };
      const allowed = reportTypes.map(t => typeMap[t]);
      if (!allowed.includes(r.report_type)) return false;
    }
    return true;
  });
  switch (sortBy) {
    case 'date-asc':
      filteredReports.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
      break;
    case 'ratio-desc':
      filteredReports.sort((a, b) => (parseFloat(b.holding_ratio) || 0) - (parseFloat(a.holding_ratio) || 0));
      break;
    case 'ratio-asc':
      filteredReports.sort((a, b) => {
        const ra = a.holding_ratio != null ? parseFloat(a.holding_ratio) : Infinity;
        const rb = b.holding_ratio != null ? parseFloat(b.holding_ratio) : Infinity;
        return ra - rb;
      });
      break;
    case 'delta-desc':
      filteredReports.sort((a, b) => {
        const da = a._delta != null ? a._delta : -Infinity;
        const db = b._delta != null ? b._delta : -Infinity;
        return db - da;
      });
      break;
    case 'delta-asc':
      filteredReports.sort((a, b) => {
        const da = a._delta != null ? a._delta : Infinity;
        const db = b._delta != null ? b._delta : Infinity;
        return da - db;
      });
      break;
    default: // date-desc
      filteredReports.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }
  renderReportTable();
}
function resetFilters() {
  document.getElementById('searchKeyword').value = '';
  document.getElementById('investorFilter').value = 'all';
  document.getElementById('periodFilter').value = '90';
  document.querySelectorAll('input[name="reportType"]').forEach(cb => cb.checked = true);
  filteredReports = [...reportData.reports];
  renderReportTable();
}
function searchByCode(code) {
  if (!code) return;
  document.getElementById('searchKeyword').value = code;
  applyFilters();
  switchTab('list', document.querySelector('.tab-btn'));
}
function filterByActivist(activistId) {
  let activist = reportData.activists[activistId];
  if (!activist) {
    // IDで見つからない場合、名前で検索
    const found = Object.entries(reportData.activists || {}).find(([, a]) => a.name === activistId);
    if (found) { activistId = found[0]; activist = found[1]; }
  }
  if (!activist) return;
  // グループの場合はメンバーの報告を全て表示
  if (activist.member_ids && activist.member_ids.length) {
    const memberIds = new Set(activist.member_ids);
    const keyword = document.getElementById('searchKeyword');
    keyword.value = '';
    // メンバーIDでフィルター
    const period = parseInt(document.getElementById('periodFilter').value);
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - period);
    const cutoffStr = cutoffDate.toISOString().slice(0, 10);
    filteredReports = reportData.reports.filter(r => {
      if (!memberIds.has(r.activist_id)) return false;
      if (r.date && r.date < cutoffStr) return false;
      return true;
    });
    renderReportTable();
  } else {
    document.getElementById('searchKeyword').value = activist.name;
    applyFilters();
  }
  switchTab('list', document.querySelector('.tab-btn'));
}
// ===== CHARTS =====
let chartsInitialized = false;
function initCharts() {
  if (chartsInitialized) return;
  chartsInitialized = true;
  // 月別集計
  const monthCounts = {};
  reportData.reports.forEach(r => {
    if (!r.date) return;
    const ym = r.date.slice(0, 7);
    if (!monthCounts[ym]) monthCounts[ym] = { total: 0, activist: 0 };
    monthCounts[ym].total++;
    if (r.is_activist) monthCounts[ym].activist++;
  });
  const months = Object.keys(monthCounts).sort().slice(-6);
  const totalData = months.map(m => monthCounts[m].total);
  const activistData = months.map(m => monthCounts[m].activist);
  new Chart(document.getElementById('trendChart'), {
    type: 'line',
    data: {
      labels: months,
      datasets: [
        { label: '全報告', data: totalData, borderColor: '#1a2d4f', backgroundColor: 'rgba(26,45,79,0.08)', fill: true, tension: 0.4 },
        { label: 'アクティビスト関連', data: activistData, borderColor: '#9b8b6e', backgroundColor: 'rgba(155,139,110,0.08)', fill: true, tension: 0.4 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { family: "'Noto Sans JP', sans-serif", size: 12 } } } },
      scales: { y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.04)' } }, x: { grid: { display: false } } }
    }
  });
}
// ===== TAB SWITCH =====
function switchTab(tab, btn) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  if (btn) btn.classList.add('active');
  if (tab === 'chart') setTimeout(initCharts, 100);
}
/* Auth modal/handler functions moved to auth.js */
async function updateAuthUI(user) {
  const authBtns = document.getElementById('navAuthBtns');
  const userInfo = document.getElementById('navUserInfo');
  const userName = document.getElementById('navUserName');
  const mobileAuthBtns = document.getElementById('mobileAuthBtns');
  const mobileUserInfo = document.getElementById('mobileUserInfo');
  const mobileUserName = document.getElementById('mobileUserName');
  const premiumEls = document.querySelectorAll('.premium-overlay');
  const premiumBlurs = document.querySelectorAll('.premium-blur');
  const loginWall = document.getElementById('loginWall');
  if (user) {
    // プロフィール未完了チェック
    try {
      const profile = await window.firebaseGetProfile(user.uid);
      if (profile && (!profile.affiliation || !profile.jobTitle || !profile.company)) {
        openModal('completeProfile');
      }
    } catch (e) { /* ignore */ }
    const displayName = (user.displayName || user.email) + ' 様';
    authBtns.style.display = 'none';
    userInfo.style.display = 'flex';
    userName.textContent = displayName;
    if (mobileAuthBtns) mobileAuthBtns.style.display = 'none';
    if (mobileUserInfo) { mobileUserInfo.style.display = 'flex'; mobileUserName.textContent = displayName; }
    premiumEls.forEach(el => el.style.display = 'none');
    premiumBlurs.forEach(el => el.classList.remove('premium-blur'));
    if (loginWall) loginWall.style.display = 'none';
    // 認証後にデータを読み込み
    if (!dataLoaded) {
      await loadData();
      renderAll();
    }
  } else {
    authBtns.style.display = 'flex';
    userInfo.style.display = 'none';
    userName.textContent = '';
    if (mobileAuthBtns) mobileAuthBtns.style.display = 'flex';
    if (mobileUserInfo) { mobileUserInfo.style.display = 'none'; mobileUserName.textContent = ''; }
    premiumEls.forEach(el => el.style.display = 'flex');
    premiumBlurs.forEach(el => el.classList.add('premium-blur'));
    if (loginWall) loginWall.style.display = 'block';
  }
}
// ===== NAV =====
// toggleMenu functionality moved to nav-handler.js
// ===== UTILS =====
function escapeHtml(str) { const d = document.createElement('div'); d.textContent = str; return d.innerHTML; }
function escapeJs(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"'); }
function isRecentReport(dateStr, days) {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  const now = new Date();
  return (now - d) < days * 86400000;
}
// ===== CSV EXPORT (会員限定) =====
function exportCSV() {
  if (!window.currentUser) {
    openModal('login');
    return;
  }
  const rows = [['日付', '報告者', '証券コード', '対象企業', '報告種別', '保有比率(%)', '保有目的', 'アクティビスト']];
  filteredReports.forEach(r => {
    rows.push([
      r.date || '',
      r.filer_name || '',
      r.sec_code || '',
      r.target_company || '',
      r.report_type || '',
      r.holding_ratio != null && !isNaN(parseFloat(r.holding_ratio)) ? parseFloat(r.holding_ratio).toFixed(2) : '',
      r.purpose || '',
      r.is_activist ? 'Yes' : 'No'
    ]);
  });
  const bom = '\uFEFF';
  const csv = bom + rows.map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'activist_reports_' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}
// ===== PDF EXPORT (会員限定) =====
function exportPDF() {
  if (!window.currentUser) {
    openModal('login');
    return;
  }
  // ブラウザ印刷（PDF保存）で日本語を完全サポート
  const data = filteredReports.slice(0, 300);
  const dateStr = new Date().toLocaleString('ja-JP');
  let html = `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<title>大量保有報告書レポート</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Noto Sans JP',sans-serif; font-size:10px; color:#1a2d4f; padding:15mm; }
  h1 { font-size:16px; margin-bottom:4px; }
  .meta { font-size:11px; color:#666; margin-bottom:10px; }
  table { width:100%; border-collapse:collapse; font-size:11px; }
  th { background:#1a2d4f; color:#fff; padding:4px 6px; text-align:left; font-weight:700; white-space:nowrap; }
  td { padding:3px 6px; border-bottom:1px solid #e0ddd8; }
  tr:nth-child(even) { background:#f8f7f5; }
  .ratio-cell { text-align:right; }
  .act-badge { display:inline-block; background:#c7a04b; color:#fff; font-size:7px; padding:1px 4px; border-radius:3px; }
  .notable-badge { display:inline-block; background:#6b8e6b; color:#fff; font-size:7px; padding:1px 4px; border-radius:3px; }
  @media print {
    body { padding:10mm; }
    @page { size:A4 landscape; margin:10mm; }
  }
</style></head><body>
<h1>大量保有報告書 レポート</h1>
<div class="meta">出力日時: ${escapeHtml(dateStr)} ｜ 件数: ${data.length} 件 ｜ データソース: EDINET（金融庁）</div>
<table>
<thead><tr>
  <th>報告日</th><th>報告者</th><th>種別</th><th>証券コード</th><th>対象企業</th>
  <th>保有比率</th><th>目的</th><th>区分</th>
</tr></thead><tbody>`;
  data.forEach(r => {
    const ratioNum = r.holding_ratio != null ? parseFloat(r.holding_ratio) : NaN;
    const ratio = !isNaN(ratioNum) ? ratioNum.toFixed(2) + '%' : '—';
    const badge = r.is_activist ? '<span class="act-badge">ACT</span>'
      : r.is_notable ? '<span class="notable-badge">注目</span>' : '—';
    html += `<tr>
      <td>${escapeHtml(r.date || '')}</td>
      <td>${escapeHtml(r.filer_name || '')}</td>
      <td>${escapeHtml(r.report_type || '')}</td>
      <td>${escapeHtml(r.sec_code || '')}</td>
      <td>${escapeHtml(r.target_company || '')}</td>
      <td class="ratio-cell">${ratio}</td>
      <td>${escapeHtml(r.purpose || '')}</td>
      <td>${badge}</td>
    </tr>`;
  });
  html += '</tbody></table></body></html>';
  const printWin = window.open('', '_blank');
  printWin.document.write(html);
  printWin.document.close();
  // フォント読み込み待ちの後に印刷ダイアログを表示
  printWin.onload = () => setTimeout(() => printWin.print(), 500);
}
// ===== WATCHLIST =====
let watchlist = JSON.parse(localStorage.getItem('activist_watchlist') || '[]');
function saveWatchlist() {
  localStorage.setItem('activist_watchlist', JSON.stringify(watchlist));
}
function addToWatchlist(name) {
  if (!name) {
    name = document.getElementById('watchlistInput').value.trim();
  }
  if (!name) return;
  // 重複チェック
  if (watchlist.some(w => w.name === name)) {
    renderWatchlist();
    document.getElementById('watchlistInput').value = '';
    return;
  }
  // known_activistsから情報を取得
  const activists = reportData.activists || {};
  let info = null;
  for (const [id, a] of Object.entries(activists)) {
    if (a.name === name || (a.name && a.name.includes(name)) || name.includes(a.name || '')) {
      info = { id, ...a };
      break;
    }
  }
  watchlist.push({
    name: info ? info.name : name,
    id: info ? info.id : '',
    type: info ? (info.type || '') : '',
    addedAt: new Date().toISOString()
  });
  saveWatchlist();
  renderWatchlist();
  renderWatchlistReports();
  document.getElementById('watchlistInput').value = '';
  // プリセットボタンの状態更新
  renderWatchlistPresets();
}
function removeFromWatchlist(name) {
  watchlist = watchlist.filter(w => w.name !== name);
  saveWatchlist();
  renderWatchlist();
  renderWatchlistReports();
  renderWatchlistPresets();
}
function renderWatchlist() {
  const container = document.getElementById('watchlistItems');
  const countEl = document.getElementById('watchlistCount');
  countEl.textContent = watchlist.length + ' 件';
  if (!watchlist.length) {
    container.innerHTML = '<div class="watchlist-empty">ウォッチリストは空です。上のボタンまたは入力欄からアクティビストを追加してください。</div>';
    return;
  }
  container.innerHTML = watchlist.map(w => {
    const typeLabel = w.type === 'activist' ? 'アクティビスト' : w.type === 'fund' ? 'ファンド' : '';
    return `<div class="watchlist-item">
      <div class="watchlist-item-info">
        <div class="watchlist-item-name">${escapeHtml(w.name)}</div>
        <div class="watchlist-item-meta">${typeLabel ? typeLabel + ' · ' : ''}追加日: ${w.addedAt ? w.addedAt.slice(0, 10) : '—'}</div>
      </div>
      <div class="watchlist-item-actions">
        <button class="watchlist-btn-remove" data-action="filterByActivist" data-arg="${escapeJs(w.name)}">報告を表示</button>
        <button class="watchlist-btn-remove" data-action="removeFromWatchlist" data-arg="${escapeJs(w.name)}">削除</button>
      </div>
    </div>`;
  }).join('');
}
function renderWatchlistPresets() {
  const container = document.getElementById('watchlistPresets');
  const activists = reportData.activists || {};
  const names = Object.values(activists).map(a => a.name).filter(Boolean).slice(0, 15);
  container.innerHTML = names.map(name => {
    const isAdded = watchlist.some(w => w.name === name);
    return `<button class="watchlist-preset-btn${isAdded ? ' added' : ''}" data-action="${isAdded ? 'removeFromWatchlist' : 'addToWatchlist'}" data-arg="${escapeJs(name)}">${isAdded ? '✓ ' : ''}${escapeHtml(name)}</button>`;
  }).join('');
}
function renderWatchlistReports() {
  const container = document.getElementById('watchlistReports');
  if (!watchlist.length) {
    container.innerHTML = '';
    return;
  }
  const watchNames = watchlist.map(w => w.name);
  const matched = reportData.reports.filter(r => {
    return watchNames.some(name => {
      const filer = r.filer_name || '';
      return filer.includes(name) || name.includes(filer);
    });
  }).slice(0, 20);
  if (!matched.length) {
    container.innerHTML = '<div style="padding:16px; color:var(--text-light); font-size:12px;">ウォッチリスト対象の報告はまだありません。</div>';
    return;
  }
  let html = `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
    <span style="font-size:12px; color:var(--text-light);">直近 ${matched.length} 件</span>
    <button class="export-btn" data-action="exportWatchlistCSV" style="font-size:11px; padding:4px 10px;">CSV出力<span class="member-badge">会員限定</span></button>
  </div>`;
  html += '<table style="width:100%; border-collapse:collapse; font-size:12px;">';
  html += '<tr style="background:var(--light-gray);"><th style="padding:6px 8px; text-align:left;">日付</th><th style="padding:6px 8px; text-align:left;">報告者</th><th style="padding:6px 8px; text-align:left;">対象</th><th style="padding:6px 8px; text-align:right;">保有比率</th></tr>';
  matched.forEach(r => {
    html += `<tr style="border-bottom:1px solid var(--light-gray);">
      <td style="padding:6px 8px;">${escapeHtml(r.date || '—')}</td>
      <td style="padding:6px 8px;">${escapeHtml(r.filer_name || '')}</td>
      <td style="padding:6px 8px;">${escapeHtml((r.target_company || '').slice(0, 30))}</td>
      <td style="padding:6px 8px; text-align:right;">${r.holding_ratio != null && !isNaN(parseFloat(r.holding_ratio)) ? parseFloat(r.holding_ratio).toFixed(1) + '%' : '—'}</td>
    </tr>`;
  });
  html += '</table>';
  container.innerHTML = html;
}
function exportWatchlistCSV() {
  if (!window.currentUser) {
    openModal('login');
    return;
  }
  const watchNames = watchlist.map(w => w.name);
  const matched = reportData.reports.filter(r => {
    return watchNames.some(name => {
      const filer = r.filer_name || '';
      return filer.includes(name) || name.includes(filer);
    });
  });
  const rows = [['日付', '報告者', '証券コード', '対象企業', '報告種別', '保有比率(%)', '保有目的']];
  matched.forEach(r => {
    rows.push([
      r.date || '', r.filer_name || '', r.sec_code || '', r.target_company || '',
      r.report_type || '', r.holding_ratio != null && !isNaN(parseFloat(r.holding_ratio)) ? parseFloat(r.holding_ratio).toFixed(2) : '', r.purpose || ''
    ]);
  });
  const bom = '\uFEFF';
  const csv = bom + rows.map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'watchlist_reports_' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}
// ウォッチリストタブ切替時に初期化
const origSwitchTab = switchTab;
switchTab = function(tab, btn) {
  origSwitchTab(tab, btn);
  if (tab === 'watchlist') {
    renderWatchlist();
    renderWatchlistPresets();
    renderWatchlistReports();
  }
};
// ===== INVESTOR DETAIL MODAL =====
function openInvestorDetail(activistId) {
  const activist = reportData.activists[activistId];
  if (!activist) return;
  document.getElementById('idmName').textContent = activist.name;
  const metaParts = [];
  if (activist.representative) metaParts.push('代表: ' + activist.representative);
  if (activist.headquarters) metaParts.push('本拠: ' + activist.headquarters);
  const typeLabel = activist.type === 'activist' ? 'アクティビスト' : activist.type === 'individual_investor' ? '個人注目投資家' : activist.type === 'notable_holder' ? '注目投資家' : 'ファンド';
  metaParts.push(typeLabel);
  if (activist.member_ids && activist.member_ids.length) metaParts.push('グループ（' + activist.member_ids.length + '主体）');
  document.getElementById('idmMeta').textContent = metaParts.join(' ｜ ');
  // この投資家の全報告を取得（グループ対応）
  let investorReports;
  if (activist.member_ids && activist.member_ids.length) {
    const memberIds = new Set(activist.member_ids);
    investorReports = reportData.reports.filter(r => memberIds.has(r.activist_id));
  } else {
    const name = activist.name;
    investorReports = reportData.reports.filter(r => r.filer_name === name || r.activist_id === activistId);
  }
  investorReports.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  // サマリー統計
  const uniqueSecCodes = new Set(investorReports.filter(r => r.sec_code).map(r => r.sec_code));
  const ratios = investorReports.map(r => parseFloat(r.holding_ratio)).filter(v => !isNaN(v));
  const maxRatio = ratios.length ? Math.max(...ratios) : null;
  const latestDate = investorReports.length ? investorReports[0].date : '—';
  const newCount = investorReports.filter(r => r.report_type === '新規報告').length;
  document.getElementById('idmSummary').innerHTML = `
    <div class="idm-stat"><div class="idm-stat-label">報告件数</div><div class="idm-stat-value">${investorReports.length}</div></div>
    <div class="idm-stat"><div class="idm-stat-label">対象銘柄数</div><div class="idm-stat-value">${uniqueSecCodes.size}</div></div>
    <div class="idm-stat"><div class="idm-stat-label">最高保有比率</div><div class="idm-stat-value">${maxRatio != null ? maxRatio.toFixed(1) + '%' : '—'}</div></div>
    <div class="idm-stat"><div class="idm-stat-label">新規報告</div><div class="idm-stat-value">${newCount}</div></div>
    <div class="idm-stat"><div class="idm-stat-label">最新報告日</div><div class="idm-stat-value" style="font-size:14px;">${latestDate}</div></div>
  `;
  // デルタ計算用に古い順でソート
  const chronological = [...investorReports].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const prevMap = {};
  chronological.forEach(r => {
    const key = (r.filer_name || '') + '|' + (r.sec_code || '');
    const cur = parseFloat(r.holding_ratio);
    if (prevMap[key] !== undefined && !isNaN(cur)) {
      r._idm_delta = cur - prevMap[key];
    } else {
      r._idm_delta = null;
    }
    if (!isNaN(cur)) prevMap[key] = cur;
  });
  // 時系列テーブル（新→旧）
  const tbody = document.getElementById('idmTableBody');
  if (!investorReports.length) {
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--text-light);">報告データがありません</td></tr>';
  } else {
    tbody.innerHTML = investorReports.map(r => {
      const ratioVal = r.holding_ratio != null ? parseFloat(r.holding_ratio) : NaN;
      const ratioStr = !isNaN(ratioVal) ? ratioVal.toFixed(1) + '%' : '—';
      let deltaStr = '—';
      if (r._idm_delta != null) {
        const d = r._idm_delta;
        const sign = d > 0 ? '+' : '';
        const cls = d > 0 ? 'idm-delta-up' : d < 0 ? 'idm-delta-down' : '';
        deltaStr = '<span class="' + cls + '">' + sign + d.toFixed(1) + '%</span>';
      }
      const typeClass = r.report_type === '新規報告' ? 'type-new' : 'type-change';
      return '<tr>' +
        '<td>' + escapeHtml(r.date || '—') + '</td>' +
        '<td><span class="activist-type ' + typeClass + '" style="font-size:10px;padding:1px 6px;">' + escapeHtml(r.report_type || '—') + '</span></td>' +
        '<td>' + (r.sec_code ? '<span class="sec-code">' + escapeHtml(r.sec_code) + '</span>' : '—') + '</td>' +
        '<td title="' + escapeHtml(r.target_company || '') + '">' + escapeHtml((r.target_company || '—').slice(0, 25)) + '</td>' +
        '<td style="text-align:right;">' + ratioStr + '</td>' +
        '<td style="text-align:right;">' + deltaStr + '</td>' +
        '<td style="font-size:11px;">' + escapeHtml((r.purpose || '—').slice(0, 20)) + '</td>' +
        '</tr>';
    }).join('');
  }
  // 銘柄別サマリー
  const holdingMap = {};
  chronological.forEach(r => {
    if (!r.sec_code) return;
    if (!holdingMap[r.sec_code]) {
      holdingMap[r.sec_code] = { sec_code: r.sec_code, target: r.target_company || '—', reports: [], latestRatio: null, firstRatio: null };
    }
    const h = holdingMap[r.sec_code];
    h.reports.push(r);
    const rv = parseFloat(r.holding_ratio);
    if (!isNaN(rv)) {
      if (h.firstRatio === null) h.firstRatio = rv;
      h.latestRatio = rv;
    }
    if (r.target_company && r.target_company !== '—') h.target = r.target_company;
  });
  const holdingSummary = document.getElementById('idmHoldingSummary');
  const holdings = Object.values(holdingMap).sort((a, b) => (b.latestRatio || 0) - (a.latestRatio || 0));
  if (!holdings.length) {
    holdingSummary.innerHTML = '<div style="font-size:12px;color:var(--text-light);padding:12px;">銘柄データがありません</div>';
  } else {
    holdingSummary.innerHTML = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px;">' +
      holdings.map(h => {
        const totalDelta = (h.latestRatio != null && h.firstRatio != null) ? h.latestRatio - h.firstRatio : null;
        let deltaLabel = '';
        if (totalDelta != null) {
          const sign = totalDelta > 0 ? '+' : '';
          const cls = totalDelta > 0 ? 'idm-delta-up' : totalDelta < 0 ? 'idm-delta-down' : '';
          deltaLabel = '<span class="' + cls + '" style="font-size:11px;margin-left:6px;">(' + sign + totalDelta.toFixed(1) + '%)</span>';
        }
        return '<div style="background:var(--off-white);border-radius:6px;padding:12px 14px;">' +
          '<div style="display:flex;justify-content:space-between;align-items:center;">' +
          '<span style="font-size:12px;font-weight:500;color:var(--navy);">' +
          '<span class="sec-code">' + escapeHtml(h.sec_code) + '</span> ' + escapeHtml(h.target.slice(0, 20)) +
          '</span></div>' +
          '<div style="margin-top:6px;font-size:13px;">' +
          '最新比率: <strong>' + (h.latestRatio != null ? h.latestRatio.toFixed(1) + '%' : '—') + '</strong>' + deltaLabel +
          '</div>' +
          '<div style="font-size:11px;color:var(--text-light);margin-top:2px;">報告 ' + h.reports.length + '件</div>' +
          '</div>';
      }).join('') + '</div>';
  }
  document.getElementById('investorDetailOverlay').classList.add('show');
  document.body.style.overflow = 'hidden';
}
function closeInvestorDetail() {
  document.getElementById('investorDetailOverlay').classList.remove('show');
  document.body.style.overflow = '';
}
/* ── Event Delegation for Page Actions ── */
(function() {
  document.addEventListener('click', function(e) {
    var el = e.target.closest('[data-action]');
    if (!el) return;
    var parts = el.dataset.action.split(':');
    var fn = parts[0];
    var arg = el.dataset.arg || parts[1] || null;
    switch(fn) {
      case 'applyFilters': applyFilters(); break;
      case 'resetFilters': resetFilters(); break;
      case 'setRankingFilter': setRankingFilter(arg); break;
      case 'switchTab': switchTab(arg, el); break;
      case 'exportCSV': exportCSV(); break;
      case 'exportPDF': exportPDF(); break;
      case 'exportWatchlistCSV': exportWatchlistCSV(); break;
      case 'addToWatchlist':
        if (arg) addToWatchlist(arg);
        else addToWatchlist();
        break;
      case 'removeFromWatchlist': removeFromWatchlist(arg); break;
      case 'closeInvestorDetail': closeInvestorDetail(); break;
      case 'searchByCode': searchByCode(arg); break;
      case 'openInvestorDetail': openInvestorDetail(arg); break;
      case 'filterByActivist': e.stopPropagation(); filterByActivist(arg); break;
      case 'stopPropagation': e.stopPropagation(); break;
    }
    if (el.tagName === 'A' && !el.getAttribute('href')?.startsWith('http')) e.preventDefault();
  });
  // Overlay click-to-close
  document.addEventListener('click', function(e) {
    if (e.target.id === 'investorDetailOverlay') closeInvestorDetail();
  });
  // Change delegation
  document.addEventListener('change', function(e) {
    var el = e.target.closest('[data-change]');
    if (!el) return;
    if (el.dataset.change === 'applyFilters') applyFilters();
  });
  // Input delegation
  document.addEventListener('input', function(e) {
    var el = e.target.closest('[data-input]');
    if (!el) return;
    if (el.dataset.input === 'filterRanking') filterRanking();
  });
})();
