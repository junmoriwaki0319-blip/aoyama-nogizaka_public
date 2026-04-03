#!/usr/bin/env node
/**
 * 全期間の欠損データを再パースするバッチスクリプト
 *
 * 特徴:
 * - 500件ごとに中間保存（中断しても再開可能）
 * - 既にパース済み（ratio+purpose両方あり）のレコードはスキップ
 * - EDINET API レート制限対応（1秒間隔）
 * - 古いドキュメント（EDINET保持期限切れ）はスキップして続行
 *
 * 使い方:
 *   EDINET_API_KEY=xxx node scripts/reparse_all.js
 *   EDINET_API_KEY=xxx node scripts/reparse_all.js --from 2025-01 --to 2025-06
 *   EDINET_API_KEY=xxx node scripts/reparse_all.js --dry-run
 */
const https = require('https');
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');

const DATA_FILE = path.join(__dirname, '..', 'data', 'reports.json');
const PROGRESS_FILE = path.join(__dirname, '..', 'data', '.reparse_progress.json');
const API_KEY = process.env.EDINET_API_KEY || '';
const SAVE_INTERVAL = 500; // 何件ごとに中間保存するか

if (!API_KEY) {
  console.error('EDINET_API_KEY environment variable required');
  process.exit(1);
}

// ─── CLI引数パース ───
const args = process.argv.slice(2);
let fromMonth = null, toMonth = null, dryRun = false;
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--from' && args[i + 1]) fromMonth = args[++i];
  else if (args[i] === '--to' && args[i + 1]) toMonth = args[++i];
  else if (args[i] === '--dry-run') dryRun = true;
}

// ─── HTTP / ZIP ユーティリティ ───

function downloadDoc(docID) {
  return new Promise((resolve, reject) => {
    const url = `https://api.edinet-fsa.go.jp/api/v2/documents/${docID}?type=1&Subscription-Key=${API_KEY}`;
    const req = https.get(url, { timeout: 60000 }, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        https.get(res.headers.location, { timeout: 60000 }, res2 => {
          const chunks = [];
          res2.on('data', c => chunks.push(c));
          res2.on('end', () => resolve({ status: res2.statusCode, buf: Buffer.concat(chunks) }));
          res2.on('error', reject);
        }).on('error', reject);
        return;
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, buf: Buffer.concat(chunks) }));
      res.on('error', reject);
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}

function readZipEntries(buf) {
  const entries = [];
  let offset = 0;
  while (offset < buf.length - 4) {
    if (buf.readUInt32LE(offset) !== 0x04034b50) break;
    const method = buf.readUInt16LE(offset + 8);
    const compSize = buf.readUInt32LE(offset + 18);
    const uncompSize = buf.readUInt32LE(offset + 22);
    const nameLen = buf.readUInt16LE(offset + 26);
    const extraLen = buf.readUInt16LE(offset + 28);
    const name = buf.slice(offset + 30, offset + 30 + nameLen).toString('utf8');
    const dataStart = offset + 30 + nameLen + extraLen;
    const rawData = buf.slice(dataStart, dataStart + compSize);
    entries.push({ name, method, rawData, compSize, uncompSize });
    offset = dataStart + compSize;
  }
  return entries;
}

function extractEntry(entry) {
  if (entry.method === 0) return entry.rawData;
  if (entry.method === 8) return zlib.inflateRawSync(entry.rawData);
  throw new Error('Unsupported compression: ' + entry.method);
}

// ─── XBRL パーサー ───

function classifyPurpose(text) {
  if (text.includes('純投資')) return '純投資';
  if (text.includes('政策')) return '政策投資';
  if (text.includes('株主提案') || text.includes('提案')) return '株主提案';
  if (text.includes('経営') || text.includes('支配') || text.includes('関与')) return '経営関与';
  if (text.includes('重要提案行為')) return '重要提案';
  return 'その他';
}

function parseXbrlZip(buf) {
  const entries = readZipEntries(buf);
  const result = {};

  let allFiles = entries.filter(e => /\.(xbrl|htm|html)$/i.test(e.name));

  // XBRL本文を優先ソート
  allFiles.sort((a, b) => {
    function score(fname) {
      let s = 0;
      if (/XBRL\//i.test(fname)) s -= 100;
      if (/honbun/.test(fname)) s -= 50;
      if (fname.endsWith('.xbrl')) s -= 30;
      if (/header/.test(fname)) s += 10;
      return s;
    }
    return score(a.name) - score(b.name);
  });

  for (const entry of allFiles) {
    let text;
    try { text = extractEntry(entry).toString('utf8'); } catch (e) { continue; }

    // target_company
    if (!result.target_company) {
      const patterns = [
        /name="[^"]*(?:[Ii]ssuer[Nn]ame|NameOfIssuer|IssuerNameJp)[^"]*"[^>]*>([^<]+)/,
        /発行者の名称[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})/,
        /発行者の名称.*?<[^>]*>\s*([^\n<]{2,40}?)\s*</,
        /株券等の発行者[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})/,
      ];
      for (const p of patterns) {
        const m = text.match(p);
        if (m) {
          const n = m[1].trim();
          if (n && n.length >= 2 && !n.includes('報告書') && !n.includes('提出者') && !n.includes('代表取締役')) {
            result.target_company = n;
            break;
          }
        }
      }
    }

    // sec_code
    if (!result.sec_code) {
      const patterns = [
        /(?:証券コード|銘柄コード)[^\dA-Za-z]{0,10}([\dA-Za-z]{4,5})/,
        /name="[^"]*(?:[Ss]ecurity[Cc]ode|SecuritiesCode)[^"]*"[^>]*>([\dA-Za-z]{4,5})/,
      ];
      for (const p of patterns) {
        const m = text.match(p);
        if (m) { result.sec_code = m[1].trim(); break; }
      }
    }

    // holding_ratio
    if (result.holding_ratio === undefined) {
      const patterns = [
        /name="[^"]*HoldingRatioOfShareCertificatesEtc[^"]*"[^>]*>([\d]+[\.．][\d]+)/,
        /<[^>]*:HoldingRatioOfShareCertificatesEtc[^>]*>([\d]+[\.．][\d]+)<\//,
        /name="[^"]*(?:HoldingRatio|OwnershipRatio)[^"]*"[^>]*>([\d]+[\.．][\d]+)/,
        /(?:保有割合|所有割合)[^\d]{0,30}?([\d]+[\.．][\d]+)\s*[%％]/,
        /([\d]+[\.．][\d]+)\s*[%％]\s*(?:（.*?保有割合|を保有)/,
      ];
      for (const p of patterns) {
        const m = text.match(p);
        if (m) {
          let val = parseFloat(m[1].replace('．', '.'));
          if (val < 1.0 && val > 0) val = Math.round(val * 10000) / 100;
          result.holding_ratio = val;
          break;
        }
      }
    }

    // purpose
    if (!result.purpose) {
      const patterns = [
        /name="[^"]*PurposeOfHolding[^"]*"[^>]*>([\s\S]*?)<\/(?:ix:nonNumeric|jplvh)/,
        /<[^>]*:PurposeOfHolding[^>]*>([\s\S]*?)<\//,
        /保有目的[^\n]{0,5}[：:]\s*([^\n<]{2,60})/,
        /(?:当該株券等の発行者の事業活動を|純投資|投資及び状況に応じて|政策投資|経営参加|株主提案)[^\n<]{0,80}/,
      ];
      for (const p of patterns) {
        const m = text.match(p);
        if (m) {
          let raw = m[1] ? m[1] : m[0];
          raw = raw.replace(/<[^>]*>/g, '').trim();
          if (raw && raw.length >= 2) {
            result.purpose = classifyPurpose(raw);
            result.purpose_detail = raw.slice(0, 100);
            break;
          }
        }
      }
    }

    if (result.target_company && result.sec_code && result.holding_ratio !== undefined && result.purpose) break;
  }

  return result;
}

function extractSecCode(raw) {
  if (!raw) return '';
  const code = String(raw).trim();
  return code.length === 5 ? code.slice(0, 4) : code;
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─── 進捗管理 ───

function loadProgress() {
  try {
    return JSON.parse(fs.readFileSync(PROGRESS_FILE, 'utf8'));
  } catch (e) {
    return { completed: {} }; // doc_id -> true
  }
}

function saveProgress(progress) {
  fs.writeFileSync(PROGRESS_FILE, JSON.stringify(progress), 'utf8');
}

// ─── メイン処理 ───

async function main() {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  const reports = data.reports;
  const progress = loadProgress();

  // 対象レポートを抽出（ratio or purpose が欠損 & 未処理）
  let toReparse = reports.filter(r => {
    if (r.holding_ratio != null && r.purpose) return false; // 既に完全
    if (progress.completed[r.doc_id]) return false; // 前回処理済み
    return true;
  });

  // 月範囲フィルタ
  if (fromMonth) toReparse = toReparse.filter(r => r.date >= fromMonth);
  if (toMonth) toReparse = toReparse.filter(r => r.date < toMonth + '-32');

  // 日付順（新→古）でソート：新しいデータの方がEDINETに残っている可能性が高い
  toReparse.sort((a, b) => b.date.localeCompare(a.date));

  console.log(`=== 全期間再パース ===`);
  console.log(`全レポート: ${reports.length} 件`);
  console.log(`対象（欠損あり）: ${toReparse.length} 件`);
  if (fromMonth || toMonth) console.log(`期間フィルタ: ${fromMonth || '(なし)'} ～ ${toMonth || '(なし)'}`);
  if (dryRun) { console.log('(dry-run: 実行せず終了)'); return; }
  console.log(`推定所要時間: ${Math.round(toReparse.length * 1.5 / 60)} 分`);
  console.log('');

  let fixed = 0, ratioFixed = 0, purposeFixed = 0;
  let skipped = 0, errors = 0, notFound = 0;

  for (let i = 0; i < toReparse.length; i++) {
    const r = toReparse[i];
    const pct = Math.round((i + 1) / toReparse.length * 100);
    process.stdout.write(`\r  [${pct}%] ${i + 1}/${toReparse.length} ${r.doc_id} (fixed:${fixed} err:${errors} 404:${notFound})`);

    try {
      await sleep(1000);
      const { status, buf } = await downloadDoc(r.doc_id);

      if (status === 404 || status === 400) {
        // EDINET保持期限切れ or 無効なドキュメント
        notFound++;
        progress.completed[r.doc_id] = 'not_found';
        if ((i + 1) % SAVE_INTERVAL === 0) { saveProgress(progress); }
        continue;
      }

      if (status !== 200 || buf.length < 100 || buf.readUInt32LE(0) !== 0x04034b50) {
        errors++;
        continue;
      }

      const xbrlData = parseXbrlZip(buf);
      let changed = false;

      if (xbrlData.holding_ratio !== undefined && r.holding_ratio == null) {
        r.holding_ratio = xbrlData.holding_ratio;
        ratioFixed++;
        changed = true;
      }
      if (xbrlData.purpose && !r.purpose) {
        r.purpose = xbrlData.purpose;
        r.purpose_detail = xbrlData.purpose_detail;
        purposeFixed++;
        changed = true;
      }
      if (xbrlData.target_company && !r.target_company) {
        r.target_company = xbrlData.target_company;
        changed = true;
      }
      if (xbrlData.sec_code && !r.sec_code) {
        r.sec_code = extractSecCode(xbrlData.sec_code);
        changed = true;
      }

      if (changed) fixed++;
      progress.completed[r.doc_id] = changed ? 'fixed' : 'no_change';

    } catch (e) {
      errors++;
      // タイムアウト等は再試行可能なのでprogressに記録しない
    }

    // 中間保存
    if ((i + 1) % SAVE_INTERVAL === 0) {
      process.stdout.write(' [saving...]');
      const now = new Date();
      const jstOffset = 9 * 60 * 60 * 1000;
      const jstDate = new Date(now.getTime() + jstOffset);
      data.last_updated = jstDate.toISOString().replace('Z', '+09:00');
      fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
      saveProgress(progress);
    }
  }

  // 最終保存
  console.log('\n\n=== 結果 ===');
  console.log(`  処理: ${toReparse.length} 件`);
  console.log(`  修正: ${fixed} 件`);
  console.log(`  比率修正: ${ratioFixed} 件`);
  console.log(`  目的修正: ${purposeFixed} 件`);
  console.log(`  404/期限切れ: ${notFound} 件`);
  console.log(`  エラー: ${errors} 件`);

  const now = new Date();
  const jstOffset = 9 * 60 * 60 * 1000;
  const jstDate = new Date(now.getTime() + jstOffset);
  data.last_updated = jstDate.toISOString().replace('Z', '+09:00');
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
  saveProgress(progress);

  console.log(`\nSaved to ${DATA_FILE}`);

  // 修正後の統計
  const remaining = reports.filter(r => r.holding_ratio == null || !r.purpose);
  console.log(`\n残りの欠損: ${remaining.length} 件 (${Math.round(remaining.length/reports.length*100)}%)`);
}

main().catch(e => { console.error(e); process.exit(1); });
