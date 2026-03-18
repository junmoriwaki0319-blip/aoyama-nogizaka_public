#!/usr/bin/env node
/**
 * reports.json の全レポートを最新の known_activists.json で再タグ付けし、
 * アクティビスト別サマリーを再構築するスクリプト。
 */
const fs = require('fs');
const path = require('path');

const BASE = path.resolve(__dirname, '..');
const KA_FILE = path.join(__dirname, 'known_activists.json');
const REPORTS_FILE = path.join(BASE, 'data', 'reports.json');

// Load known activists
const ka = JSON.parse(fs.readFileSync(KA_FILE, 'utf8'));
const activists = ka.activists;
const groups = ka.groups || {};

/** 全角英数字・記号を半角に正規化 */
function normalizeWidth(text) {
  return text.replace(/[\uFF01-\uFF5E]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
  ).replace(/\u3000/g, ' ');
}

function matchActivist(filerName, activists) {
  const filerNorm = normalizeWidth(filerName);
  for (const a of activists) {
    const names = [a.name, ...(a.aliases || [])];
    for (const name of names) {
      const nameNorm = normalizeWidth(name);
      if (filerNorm.includes(nameNorm) || nameNorm.includes(filerNorm)) {
        return a;
      }
    }
  }
  return null;
}

// Load reports.json
const data = JSON.parse(fs.readFileSync(REPORTS_FILE, 'utf8'));
let retagCount = 0;

for (const r of data.reports) {
  const filerName = r.filer_name || '';
  if (!filerName) continue;
  const matched = matchActivist(filerName, activists);
  const oldId = r.activist_id || '';
  if (matched) {
    const newId = matched.id;
    const investorType = matched.type || 'activist';
    if (newId !== oldId) retagCount++;
    r.activist_id = newId;
    r.activist_type = investorType;
    if (investorType === 'notable_holder') {
      r.is_activist = false;
      r.is_notable = true;
    } else {
      r.is_activist = true;
      r.is_notable = false;
    }
  }
}

console.log('Re-tagged:', retagCount, 'reports');

// Rebuild activist summary with group merging
const idToGroup = {};
for (const a of activists) {
  if (a.group_id && groups[a.group_id]) {
    idToGroup[a.id] = a.group_id;
  }
}

const activistHoldings = {};

for (const r of data.reports) {
  if (!r.is_activist && !r.is_notable) continue;
  const aid = r.activist_id || '';
  if (!aid) continue;
  const effectiveId = idToGroup[aid] || aid;

  if (!activistHoldings[effectiveId]) {
    if (groups[effectiveId]) {
      const g = groups[effectiveId];
      activistHoldings[effectiveId] = {
        id: effectiveId, name: g.name || '', type: g.type || 'activist',
        representative: g.representative || '', headquarters: '',
        description: g.description || '', focus_sectors: [],
        holdings: [], report_count: 0, latest_date: '',
        member_ids: activists.filter(a => a.group_id === effectiveId).map(a => a.id)
      };
    } else {
      const base = activists.find(a => a.id === aid) || {};
      activistHoldings[effectiveId] = {
        id: effectiveId, name: base.name || r.filer_name || '',
        type: base.type || 'fund', representative: base.representative || '',
        headquarters: base.headquarters || '', description: base.description || '',
        focus_sectors: base.focus_sectors || [],
        holdings: [], report_count: 0, latest_date: ''
      };
    }
  }

  const entry = activistHoldings[effectiveId];
  entry.report_count++;
  if ((r.date || '') > (entry.latest_date || '')) entry.latest_date = r.date;

  const secCode = r.sec_code || '';
  const existing = entry.holdings.find(h => h.sec_code === secCode);
  if (existing) {
    if ((r.date || '') >= (existing.date || '')) {
      existing.date = r.date;
      existing.holding_ratio = r.holding_ratio;
      existing.purpose = r.purpose || '';
      existing.report_type = r.report_type || '';
    }
  } else {
    entry.holdings.push({
      sec_code: secCode, filer_name: r.filer_name || '',
      target_company: r.target_company || '', date: r.date,
      holding_ratio: r.holding_ratio, purpose: r.purpose || '',
      report_type: r.report_type || ''
    });
  }
}

// Sort by report_count descending
const sorted = Object.entries(activistHoldings)
  .sort((a, b) => b[1].report_count - a[1].report_count);
data.activists = Object.fromEntries(sorted);
data.activist_reports = data.reports.filter(r => r.is_activist).length;
data.last_updated = new Date().toISOString();

fs.writeFileSync(REPORTS_FILE, JSON.stringify(data, null, 2), 'utf8');

// Show oasis specifically
const oasis = data.activists['oasis'];
console.log('Oasis reports:', oasis ? oasis.report_count : 0);
if (oasis && oasis.holdings) {
  oasis.holdings.forEach(h => {
    const ratio = h.holding_ratio != null ? h.holding_ratio + '%' : '—';
    console.log('  ', h.target_company || h.sec_code, ratio, h.date);
  });
}
console.log('Total activist reports:', data.activist_reports);
console.log('Total tracked investors:', sorted.length);
console.log('Done. Written to', REPORTS_FILE);
