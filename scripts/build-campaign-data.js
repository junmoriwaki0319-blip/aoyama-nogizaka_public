#!/usr/bin/env node
/**
 * build-campaign-data.js
 *
 * reports.json + known_activists.json からアクティビスト・キャンペーンを自動検出し、
 * 手動キュレーション事例とマージして /data/activist-campaigns.json を生成する。
 *
 * オプション:
 *   --archive   Wayback Machine に各URLの保存リクエストを送信
 *
 * Usage:
 *   node scripts/build-campaign-data.js
 *   node scripts/build-campaign-data.js --archive
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const REPORTS_PATH = path.join(ROOT, 'data', 'reports.json');
const ACTIVISTS_PATH = path.join(ROOT, 'scripts', 'known_activists.json');
const OUTPUT_PATH = path.join(ROOT, 'data', 'activist-campaigns.json');

const doArchive = process.argv.includes('--archive');

// ─── 手動キュレーション事例（HTMLページの6件 + 追加可能） ───
const CURATED_CAMPAIGNS = [
  {
    id: 'sqex_3d_2025',
    activist_id: '3d_investment',
    activist_name: '3D Investment Partners',
    target_company: 'スクウェア・エニックスHD',
    sec_code: '9684',
    category: 'エンタメ・ゲーム',
    date_start: '2025-12',
    tagline: '「あの感動をもう一度」— 110ページ超の経営改革プレゼンテーション',
    campaign_type: 'presentation',
    materials: [
      { type: 'pdf', label: 'プレゼン資料 (PDF)', url: 'https://www.3dipartners.com/wp-content/uploads/square-enix-presentation-material-jp-202512.pdf' },
      { type: 'article', label: '報道記事', url: 'https://gamestalk.net/square-enix-3d-investment-shareholder-proposal/' },
      { type: 'response', label: 'スクエニ側 大量保有対応', url: 'https://www.hd.square-enix.com/jpn/ir/pdf/20250616_01.pdf' }
    ],
    holding_ratio: 14.36,
    is_curated: true
  },
  {
    id: 'gungho_sc_2025',
    activist_id: 'strategic_capital',
    activist_name: 'ストラテジックキャピタル',
    target_company: 'ガンホー・オンライン',
    sec_code: '3765',
    category: 'エンタメ・ゲーム',
    date_start: '2025-01',
    tagline: '「ガンホー再起の処方箋」— 特設サイト開設型キャンペーン',
    campaign_type: 'website',
    materials: [
      { type: 'website', label: '特設サイト「再起の処方箋」', url: 'https://stracap.jp/3765-GUNGHO/' },
      { type: 'pdf', label: 'SC反論書 (PDF)', url: 'https://stracap.jp/3765-GUNGHO/Objection.pdf' },
      { type: 'press', label: 'プレスリリース', url: 'https://prtimes.jp/main/html/rd/p/000000088.000052343.html' }
    ],
    holding_ratio: 5.4,
    is_curated: true
  },
  {
    id: 'sapporo_3d_2023',
    activist_id: '3d_investment',
    activist_name: '3D Investment Partners',
    target_company: 'サッポロHD',
    sec_code: '2501',
    category: '消費財',
    date_start: '2023-01',
    date_end: '2025-03',
    tagline: '不動産スピンオフ提案 — 特設サイト「compoundsapporo.com」で長期キャンペーン',
    campaign_type: 'website',
    materials: [
      { type: 'website', label: 'プレゼン資料ページ', url: 'https://www.compoundsapporo.com/presentation' },
      { type: 'press', label: 'プレスリリース一覧', url: 'https://www.compoundsapporo.com/press-release' }
    ],
    is_curated: true
  },
  {
    id: 'fujisoft_3d_2024',
    activist_id: '3d_investment',
    activist_name: '3D Investment Partners',
    target_company: '富士ソフト',
    sec_code: '9749',
    category: 'IT',
    date_start: '2024-02',
    tagline: '「企業価値最大化に向けて」— 非公開化提案を含む包括プレゼン',
    campaign_type: 'presentation',
    materials: [
      { type: 'pdf', label: 'プレゼン資料 (PDF)', url: 'https://www.3dipartners.com/wp-content/uploads/fujisoft-presentation-on-shareholderproposal-jp-202402.pdf' },
      { type: 'website', label: '特設サイト', url: 'https://www.compoundfujisoft.com/proposal' }
    ],
    is_curated: true
  },
  {
    id: 'fujimedia_dalton_2025',
    activist_id: 'dalton',
    activist_name: 'ダルトン・インベストメンツ',
    target_company: 'フジ・メディアHD',
    sec_code: '4676',
    category: 'メディア',
    date_start: '2025-04',
    date_end: '2025-06',
    tagline: '取締役12名「総入れ替え」提案 — 北尾吉孝氏を含む候補者リスト',
    campaign_type: 'proxy_fight',
    materials: [
      { type: 'announcement', label: '株主提案公表', url: 'https://www.daltoninvestments.co.jp/en/news/20250416' },
      { type: 'letter', label: '株主向けレター', url: 'https://www.daltoninvestments.co.jp/en/news/20250601' }
    ],
    is_curated: true
  },
  {
    id: 'nssol_3d_2025',
    activist_id: '3d_investment',
    activist_name: '3D Investment Partners',
    target_company: 'NSソリューションズ',
    sec_code: '2327',
    category: 'IT',
    date_start: '2025-03',
    tagline: '企業価値最大化を阻む要因 — プレゼンテーション公開',
    campaign_type: 'presentation',
    materials: [
      { type: 'press', label: 'BusinessWire発表', url: 'https://www.businesswire.com/news/home/20250328623776/en/3D-Investment-Partners-Releases-Investor-Presentation-Highlighting-Issues-Preventing-NS-Solutions-from-Maximizing-Corporate-Value' }
    ],
    is_curated: true
  }
];

// ─── メイン処理 ───
async function main() {
  console.log('[build-campaign] Loading data...');

  const reportsData = JSON.parse(fs.readFileSync(REPORTS_PATH, 'utf8'));
  const activistsData = JSON.parse(fs.readFileSync(ACTIVISTS_PATH, 'utf8'));
  const reports = reportsData.reports || [];

  // アクティビスト名→IDマッピングを構築
  const activistNameMap = new Map();
  const activistById = new Map();
  for (const a of activistsData.activists) {
    activistById.set(a.id, a);
    activistNameMap.set(a.name, a.id);
    for (const alias of (a.aliases || [])) {
      activistNameMap.set(alias, a.id);
    }
  }

  // グループ情報
  const groups = activistsData.groups || {};

  // ─── 自動検出: purpose="株主提案" または "経営関与" でアクティビスト紐づき ───
  console.log('[build-campaign] Detecting campaigns from reports...');

  // キャンペーン = (activist_id, sec_code) の組み合わせ
  const campaignMap = new Map(); // key: `${activist_id}__${sec_code}`

  for (const r of reports) {
    const isProposal = r.purpose === '株主提案' || r.purpose === '経営関与';
    const isActivist = r.is_activist || r.is_notable;

    if (!isProposal && !isActivist) continue;

    // activist_id を特定
    let aid = r.activist_id || null;
    if (!aid && r.filer_name) {
      aid = activistNameMap.get(r.filer_name) || null;
    }
    if (!aid) continue;

    const key = `${aid}__${r.sec_code}`;

    if (!campaignMap.has(key)) {
      campaignMap.set(key, {
        activist_id: aid,
        sec_code: r.sec_code,
        target_company: r.target_company,
        reports: [],
        purposes: new Set(),
        max_ratio: 0
      });
    }

    const c = campaignMap.get(key);
    c.reports.push({
      doc_id: r.doc_id,
      date: r.date,
      report_type: r.report_type,
      holding_ratio: r.holding_ratio,
      purpose: r.purpose,
      purpose_detail: r.purpose_detail || '',
      edinet_url: r.edinet_url
    });
    if (r.purpose) c.purposes.add(r.purpose);
    if (r.holding_ratio > c.max_ratio) c.max_ratio = r.holding_ratio;
  }

  console.log(`[build-campaign] Found ${campaignMap.size} activist-company pairs`);

  // ─── 自動検出キャンペーンをフォーマット ───
  const autoDetected = [];

  for (const [key, c] of campaignMap) {
    const activist = activistById.get(c.activist_id);
    if (!activist) continue;

    // 日付でソート
    c.reports.sort((a, b) => (a.date || '').localeCompare(b.date || ''));

    const firstDate = c.reports[0]?.date || '';
    const lastDate = c.reports[c.reports.length - 1]?.date || '';

    // キャンペーンタイプ推定
    let campaignType = 'filing'; // デフォルト: EDINET報告のみ
    const purposeArr = [...c.purposes];
    if (purposeArr.includes('株主提案')) campaignType = 'shareholder_proposal';
    else if (purposeArr.includes('経営関与')) campaignType = 'engagement';

    // グループ解決
    let groupName = null;
    if (activist.group_id && groups[activist.group_id]) {
      groupName = groups[activist.group_id].name;
    }

    autoDetected.push({
      id: `auto_${c.activist_id}_${c.sec_code}`,
      activist_id: c.activist_id,
      activist_name: activist.name,
      activist_type: activist.type,
      group_id: activist.group_id || null,
      group_name: groupName,
      target_company: c.target_company,
      sec_code: c.sec_code,
      category: null,
      date_start: firstDate.slice(0, 7) || null,
      date_end: lastDate.slice(0, 7) || null,
      campaign_type: campaignType,
      holding_ratio_max: c.max_ratio,
      purposes: purposeArr,
      report_count: c.reports.length,
      filing_history: c.reports,
      materials: [],
      is_curated: false
    });
  }

  // ─── キュレーション事例とマージ ───
  // キュレーション事例のキーを生成
  const curatedKeys = new Set();
  for (const cc of CURATED_CAMPAIGNS) {
    curatedKeys.add(`${cc.activist_id}__${cc.sec_code}`);
  }

  // 自動検出とキュレーション事例をマージ
  const merged = [];

  // まずキュレーション事例を追加（自動検出データで補完）
  for (const cc of CURATED_CAMPAIGNS) {
    const autoKey = `${cc.activist_id}__${cc.sec_code}`;
    const autoData = campaignMap.get(autoKey);

    const enriched = { ...cc };
    if (autoData) {
      autoData.reports.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
      enriched.report_count = autoData.reports.length;
      enriched.filing_history = autoData.reports;
      enriched.holding_ratio_max = autoData.max_ratio;
      enriched.purposes = [...autoData.purposes];
      if (!enriched.holding_ratio && autoData.max_ratio) {
        enriched.holding_ratio = autoData.max_ratio;
      }
    } else {
      enriched.report_count = 0;
      enriched.filing_history = [];
      enriched.purposes = ['株主提案'];
    }

    const activist = activistById.get(cc.activist_id);
    if (activist) {
      enriched.activist_type = activist.type;
      enriched.group_id = activist.group_id || null;
      enriched.group_name = activist.group_id && groups[activist.group_id]
        ? groups[activist.group_id].name : null;
    }

    merged.push(enriched);
  }

  // 自動検出のうちキュレーション済みでないものを追加
  for (const ad of autoDetected) {
    const key = `${ad.activist_id}__${ad.sec_code}`;
    if (!curatedKeys.has(key)) {
      merged.push(ad);
    }
  }

  // ソート: キュレーション済み優先、日付新しい順
  merged.sort((a, b) => {
    if (a.is_curated !== b.is_curated) return a.is_curated ? -1 : 1;
    return (b.date_start || '').localeCompare(a.date_start || '');
  });

  // ─── アクティビスト別サマリ ───
  const activistSummary = {};
  for (const c of merged) {
    if (!activistSummary[c.activist_id]) {
      activistSummary[c.activist_id] = {
        id: c.activist_id,
        name: c.activist_name,
        type: c.activist_type || 'activist',
        group_id: c.group_id || null,
        group_name: c.group_name || null,
        campaign_count: 0,
        curated_count: 0,
        target_companies: []
      };
    }
    const s = activistSummary[c.activist_id];
    s.campaign_count++;
    if (c.is_curated) s.curated_count++;
    s.target_companies.push({
      sec_code: c.sec_code,
      name: c.target_company,
      campaign_type: c.campaign_type
    });
  }

  // ─── Wayback Machine アーカイブ ───
  const archiveResults = [];
  if (doArchive) {
    console.log('[build-campaign] Submitting URLs to Wayback Machine...');
    const urlsToArchive = [];

    for (const c of merged) {
      for (const m of (c.materials || [])) {
        if (m.url) urlsToArchive.push(m.url);
      }
    }

    const uniqueUrls = [...new Set(urlsToArchive)];
    console.log(`[build-campaign] Archiving ${uniqueUrls.length} unique URLs...`);

    for (const url of uniqueUrls) {
      try {
        const resp = await fetch(`https://web.archive.org/save/${url}`, {
          method: 'GET',
          headers: { 'User-Agent': 'aoyama-nogizaka-campaign-archiver/1.0' }
        });
        const archived = resp.ok;
        const archiveUrl = archived ? `https://web.archive.org/web/${url}` : null;
        archiveResults.push({ url, archived, archive_url: archiveUrl });
        if (archived) {
          console.log(`  ✓ Archived: ${url}`);
        } else {
          console.log(`  ✗ Failed (${resp.status}): ${url}`);
        }
        // Rate limit: 1 request per 5 seconds
        await new Promise(r => setTimeout(r, 5000));
      } catch (e) {
        console.log(`  ✗ Error: ${url} — ${e.message}`);
        archiveResults.push({ url, archived: false, error: e.message });
      }
    }
  }

  // ─── 出力 ───
  const output = {
    _generated: new Date().toISOString(),
    _description: 'アクティビスト・キャンペーンデータ（自動検出 + キュレーション）',
    stats: {
      total_campaigns: merged.length,
      curated_campaigns: merged.filter(c => c.is_curated).length,
      auto_detected: merged.filter(c => !c.is_curated).length,
      unique_activists: Object.keys(activistSummary).length,
      archive_urls: archiveResults.length
    },
    activists: Object.values(activistSummary).sort((a, b) => b.campaign_count - a.campaign_count),
    campaigns: merged,
    archive_log: archiveResults.length > 0 ? archiveResults : undefined
  };

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(output, null, 2), 'utf8');
  console.log(`[build-campaign] Output: ${OUTPUT_PATH}`);
  console.log(`[build-campaign] Stats: ${output.stats.total_campaigns} campaigns (${output.stats.curated_campaigns} curated, ${output.stats.auto_detected} auto-detected), ${output.stats.unique_activists} activists`);
}

main().catch(e => { console.error(e); process.exit(1); });
