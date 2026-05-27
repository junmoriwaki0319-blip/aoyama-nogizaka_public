/**
 * activist-campaigns-page.js
 * アクティビスト・キャンペーン一覧ページのスクリーニング・表示ロジック
 */

const DATA_URL = '/data/activist-campaigns.json';
const PER_PAGE = 20;

let campaignData = null;
let filteredCampaigns = [];
let currentPage = 1;
let selectedActivist = null;

// ─── 初期化 ───
document.addEventListener('DOMContentLoaded', init);

async function init() {
  let loadFailed = false;
  try {
    const resp = await fetch(DATA_URL + '?t=' + Date.now(), { cache: 'no-store' });
    if (!resp.ok) throw new Error('Failed to load');
    campaignData = await resp.json();
  } catch (e) {
    console.error('[campaigns] Load error:', e);
    campaignData = { campaigns: [], activists: [], stats: {} };
    loadFailed = true;
  }

  renderStats();
  renderActivistList();
  applyFilters();
  bindEvents();

  if (loadFailed) {
    document.getElementById('campaignList').innerHTML =
      '<div class="empty-state"><div class="empty-state-icon">&#x26A0;&#xFE0F;</div>' +
      '<div class="empty-state-text">データの取得に失敗しました。ネットワーク接続を確認の上、ページをリロードしてください。</div></div>';
  }
}

// ─── Stats ───
function renderStats() {
  const s = campaignData.stats || {};
  document.getElementById('statTotal').textContent = (s.total_campaigns || 0).toLocaleString();
  document.getElementById('statCurated').textContent = s.curated_campaigns || 0;
  document.getElementById('statActivists').textContent = s.unique_activists || 0;

  const proposalCount = (campaignData.campaigns || []).filter(
    c => c.campaign_type === 'shareholder_proposal' || (c.purposes && c.purposes.includes('株主提案'))
  ).length;
  document.getElementById('statProposals').textContent = proposalCount.toLocaleString();
}

// ─── Activist Sidebar ───
function renderActivistList() {
  const list = document.getElementById('activistList');
  const activists = campaignData.activists || [];

  // 「すべて」を先頭に
  let html = '<li class="activist-item active" data-activist="">' +
    '<span class="activist-item-name">すべての投資家</span>' +
    '<span class="activist-item-count">' + (campaignData.campaigns || []).length + '</span></li>';

  for (const a of activists) {
    html += '<li class="activist-item" data-activist="' + a.id + '">' +
      '<span class="activist-item-name">' + escHtml(a.name) + '</span>' +
      '<span class="activist-item-count">' + a.campaign_count + '</span></li>';
  }

  list.innerHTML = html;
}

// ─── Filters ───
function applyFilters() {
  const typeFilter = document.getElementById('filterType').value;
  const search = document.getElementById('filterSearch').value.trim().toLowerCase();
  const curatedOnly = document.getElementById('filterCuratedOnly').checked;

  let results = campaignData.campaigns || [];

  // Activist filter
  if (selectedActivist) {
    results = results.filter(c => c.activist_id === selectedActivist);
  }

  // Type filter
  if (typeFilter) {
    results = results.filter(c => {
      if (typeFilter === 'shareholder_proposal') {
        return c.campaign_type === 'shareholder_proposal' || (c.purposes && c.purposes.includes('株主提案'));
      }
      if (typeFilter === 'engagement') {
        return c.campaign_type === 'engagement' || (c.purposes && c.purposes.includes('経営関与'));
      }
      return c.campaign_type === typeFilter;
    });
  }

  // Search
  if (search) {
    results = results.filter(c => {
      const haystack = [
        c.target_company, c.sec_code, c.activist_name,
        c.tagline, c.category
      ].filter(Boolean).join(' ').toLowerCase();
      return haystack.includes(search);
    });
  }

  // Curated only
  if (curatedOnly) {
    results = results.filter(c => c.is_curated);
  }

  // Sort: curated優先 → date_start 降順 → activist名
  results = results.slice().sort((a, b) => {
    if (!!b.is_curated - !!a.is_curated !== 0) return !!b.is_curated - !!a.is_curated;
    const da = a.date_start || '';
    const db = b.date_start || '';
    if (db !== da) return db.localeCompare(da);
    return (a.activist_name || '').localeCompare(b.activist_name || '');
  });

  filteredCampaigns = results;
  currentPage = 1;

  document.getElementById('resultCount').textContent = results.length + '件';
  renderCampaigns();
  renderPagination();
}

// ─── Render Campaigns ───
function renderCampaigns() {
  const container = document.getElementById('campaignList');
  const start = (currentPage - 1) * PER_PAGE;
  const page = filteredCampaigns.slice(start, start + PER_PAGE);

  if (!page.length) {
    container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">&#x1F50D;</div><div class="empty-state-text">条件に一致するキャンペーンがありません</div></div>';
    return;
  }

  container.innerHTML = page.map((c, idx) => renderCampaignCard(c, start + idx)).join('');
}

function renderCampaignCard(c, idx) {
  const typeLabel = getCampaignTypeLabel(c.campaign_type);
  const typeBadgeClass = getCampaignTypeBadgeClass(c.campaign_type);

  let html = '<article class="campaign-card' + (c.is_curated ? ' curated' : '') + '">';

  // Header
  html += '<div class="campaign-card-header">';
  html += '<div class="campaign-card-title">' + escHtml(c.target_company || '—') + ' (' + escHtml(c.sec_code || '') + ')' +
    ' &times; ' + escHtml(c.activist_name || '—') + '</div>';
  html += '<div class="campaign-card-meta">';
  if (c.is_curated) html += '<span class="badge badge-curated">CURATED</span>';
  html += '<span class="badge ' + typeBadgeClass + '">' + typeLabel + '</span>';
  if (c.date_start) html += '<span style="font-size:12px;color:var(--text-light);">' + escHtml(c.date_start) + (c.date_end && c.date_end !== c.date_start ? ' 〜 ' + escHtml(c.date_end) : '') + '</span>';
  html += '</div></div>';

  // Tagline
  if (c.tagline) {
    html += '<div class="campaign-tagline">' + escHtml(c.tagline) + '</div>';
  }

  // Body
  html += '<div class="campaign-card-body">';
  if (c.holding_ratio || c.holding_ratio_max) {
    html += '<div class="campaign-detail"><span class="campaign-detail-label">最大保有比率</span><span class="campaign-detail-value">' + (c.holding_ratio || c.holding_ratio_max || 0).toFixed(1) + '%</span></div>';
  }
  if (c.category) {
    html += '<div class="campaign-detail"><span class="campaign-detail-label">セクター</span><span class="campaign-detail-value">' + escHtml(c.category) + '</span></div>';
  }
  if (c.report_count) {
    html += '<div class="campaign-detail"><span class="campaign-detail-label">EDINET報告数</span><span class="campaign-detail-value">' + c.report_count + '件</span></div>';
  }
  if (c.purposes && c.purposes.length) {
    html += '<div class="campaign-detail"><span class="campaign-detail-label">保有目的</span><span class="campaign-detail-value">' + c.purposes.map(escHtml).join(', ') + '</span></div>';
  }
  if (c.group_name) {
    html += '<div class="campaign-detail"><span class="campaign-detail-label">グループ</span><span class="campaign-detail-value">' + escHtml(c.group_name) + '</span></div>';
  }
  html += '</div>';

  // Materials (5+ items → fold into <details>)
  if (c.materials && c.materials.length) {
    const FOLD_THRESHOLD = 5;
    const renderMaterial = (m) => {
      const iconLabel = getMaterialIcon(m.type);
      let s = '<a href="' + escHtml(m.url) + '" target="_blank" rel="noopener" class="material-link">' +
        '<span aria-hidden="true">' + iconLabel + '</span> ' + escHtml(m.label) + '</a>';
      // Wayback archive link
      s += '<a href="https://web.archive.org/web/' + escHtml(m.url) + '" target="_blank" rel="noopener" ' +
        'class="material-link material-link-archive" aria-label="Wayback Machine アーカイブを開く" ' +
        'title="Wayback Machine アーカイブ"><span aria-hidden="true">&#x1F4BE;</span> Archive</a>';
      // Local backup link (separate green-tinted class)
      if (m.backup_url) {
        s += '<a href="' + escHtml(m.backup_url) + '" target="_blank" rel="noopener" ' +
          'class="material-link material-link-local" aria-label="ローカルバックアップを開く" ' +
          'title="ローカルバックアップ"><span aria-hidden="true">&#x1F5C2;&#xFE0F;</span> Local</a>';
      }
      return s;
    };
    if (c.materials.length <= FOLD_THRESHOLD) {
      html += '<div class="campaign-materials">' + c.materials.map(renderMaterial).join('') + '</div>';
    } else {
      const visible = c.materials.slice(0, 3);
      const hidden = c.materials.slice(3);
      html += '<div class="campaign-materials">' + visible.map(renderMaterial).join('') + '</div>';
      html += '<details class="campaign-materials-more">' +
        '<summary>他 ' + hidden.length + ' 件の資料を表示</summary>' +
        '<div class="campaign-materials">' + hidden.map(renderMaterial).join('') + '</div>' +
        '</details>';
    }
  }

  // Filing history toggle
  if (c.filing_history && c.filing_history.length) {
    html += '<button class="campaign-filing-toggle" data-toggle="filing-' + idx + '">EDINET報告履歴 (' + c.filing_history.length + '件) ▼</button>';
    html += '<div class="filing-history" id="filing-' + idx + '">';
    html += '<table class="filing-table"><thead><tr><th>日付</th><th>種別</th><th>保有比率</th><th>目的</th><th>EDINET</th></tr></thead><tbody>';
    for (const f of c.filing_history.slice(0, 20)) {
      html += '<tr><td>' + escHtml(f.date || '—') + '</td><td>' + escHtml(f.report_type || '') + '</td>' +
        '<td style="text-align:right;">' + (f.holding_ratio != null ? f.holding_ratio.toFixed(1) + '%' : '—') + '</td>' +
        '<td>' + escHtml(f.purpose || '') + '</td>' +
        '<td><a href="' + escHtml(f.edinet_url || '') + '" target="_blank" rel="noopener">原文</a></td></tr>';
    }
    if (c.filing_history.length > 20) {
      html += '<tr><td colspan="5" style="text-align:center;color:var(--text-light);">…他 ' + (c.filing_history.length - 20) + '件</td></tr>';
    }
    html += '</tbody></table></div>';
  }

  html += '</article>';
  return html;
}

// ─── Pagination ───
function renderPagination() {
  const container = document.getElementById('pagination');
  const totalPages = Math.ceil(filteredCampaigns.length / PER_PAGE);

  if (totalPages <= 1) {
    container.innerHTML = '';
    return;
  }

  let html = '';
  const start = Math.max(1, currentPage - 3);
  const end = Math.min(totalPages, currentPage + 3);

  if (currentPage > 1) {
    html += '<button class="page-btn" data-page="' + (currentPage - 1) + '">&laquo;</button>';
  }
  for (let i = start; i <= end; i++) {
    html += '<button class="page-btn' + (i === currentPage ? ' active' : '') + '" data-page="' + i + '">' + i + '</button>';
  }
  if (currentPage < totalPages) {
    html += '<button class="page-btn" data-page="' + (currentPage + 1) + '">&raquo;</button>';
  }

  container.innerHTML = html;
}

// ─── Events ───
function bindEvents() {
  document.getElementById('filterType').addEventListener('change', applyFilters);
  document.getElementById('filterSearch').addEventListener('input', debounce(applyFilters, 300));
  document.getElementById('filterCuratedOnly').addEventListener('change', applyFilters);
  document.getElementById('filterReset').addEventListener('click', () => {
    document.getElementById('filterType').value = '';
    document.getElementById('filterSearch').value = '';
    document.getElementById('filterCuratedOnly').checked = false;
    selectedActivist = null;
    updateActivistSelection();
    applyFilters();
  });

  // Activist sidebar
  document.getElementById('activistList').addEventListener('click', (e) => {
    const item = e.target.closest('.activist-item');
    if (!item) return;
    selectedActivist = item.dataset.activist || null;
    updateActivistSelection();
    applyFilters();
  });

  // Pagination
  document.getElementById('pagination').addEventListener('click', (e) => {
    const btn = e.target.closest('.page-btn');
    if (!btn) return;
    currentPage = parseInt(btn.dataset.page);
    renderCampaigns();
    renderPagination();
    document.getElementById('main').scrollIntoView({ behavior: 'smooth' });
  });

  // Filing history toggle
  document.getElementById('campaignList').addEventListener('click', (e) => {
    const toggle = e.target.closest('.campaign-filing-toggle');
    if (!toggle) return;
    const targetId = toggle.dataset.toggle;
    const el = document.getElementById(targetId);
    if (el) {
      el.classList.toggle('open');
      toggle.textContent = el.classList.contains('open')
        ? toggle.textContent.replace('▼', '▲')
        : toggle.textContent.replace('▲', '▼');
    }
  });
}

function updateActivistSelection() {
  const items = document.querySelectorAll('#activistList .activist-item');
  items.forEach(item => {
    item.classList.toggle('active', (item.dataset.activist || '') === (selectedActivist || ''));
  });
}

// ─── Helpers ───
function getCampaignTypeLabel(type) {
  const map = {
    shareholder_proposal: '株主提案',
    engagement: '経営関与',
    presentation: 'プレゼン公開',
    website: '特設サイト',
    proxy_fight: 'プロキシーファイト',
    filing: 'EDINET報告'
  };
  return map[type] || type || '—';
}

function getCampaignTypeBadgeClass(type) {
  const map = {
    shareholder_proposal: 'badge-proposal',
    engagement: 'badge-engagement',
    presentation: 'badge-presentation',
    website: 'badge-website',
    proxy_fight: 'badge-proxy',
    filing: 'badge-filing'
  };
  return map[type] || 'badge-filing';
}

function getMaterialIcon(type) {
  const map = {
    pdf: '&#x1F4C4;',
    website: '&#x1F310;',
    article: '&#x1F4F0;',
    press: '&#x1F4E2;',
    response: '&#x1F4DD;',
    letter: '&#x2709;',
    announcement: '&#x1F4E3;'
  };
  return map[type] || '&#x1F517;';
}

function escHtml(s) {
  if (!s) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function debounce(fn, ms) {
  let timer;
  return function () {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, arguments), ms);
  };
}
