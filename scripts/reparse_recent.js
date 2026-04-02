#!/usr/bin/env node
/**
 * 直近1週間分の欠損データを再パースするスクリプト
 * 修正後のPythonパーサーロジックをNode.jsで再現し、reports.jsonを更新する
 */
const https = require('https');
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');

const DATA_FILE = path.join(__dirname, '..', 'data', 'reports.json');
const API_KEY = process.env.EDINET_API_KEY || '';

if (!API_KEY) {
  console.error('EDINET_API_KEY environment variable required');
  process.exit(1);
}

function downloadDoc(docID) {
  return new Promise((resolve, reject) => {
    const url = `https://api.edinet-fsa.go.jp/api/v2/documents/${docID}?type=1&Subscription-Key=${API_KEY}`;
    https.get(url, { timeout: 60000 }, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        https.get(res.headers.location, { timeout: 60000 }, res2 => {
          const chunks = [];
          res2.on('data', c => chunks.push(c));
          res2.on('end', () => resolve(Buffer.concat(chunks)));
          res2.on('error', reject);
        }).on('error', reject);
        return;
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
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

  // Sort: XBRL/ dir + honbun first, header last
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

async function main() {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  const reports = data.reports;

  // Find reports from last 7 days that are missing ratio or purpose
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 8);
  const cutoffStr = cutoff.toISOString().slice(0, 10);

  const toReparse = reports.filter(r =>
    r.date >= cutoffStr && (r.holding_ratio == null || !r.purpose)
  );

  console.log(`Re-parsing ${toReparse.length} reports from last 7 days...`);

  let fixed = 0;
  let ratioFixed = 0;
  let purposeFixed = 0;

  for (let i = 0; i < toReparse.length; i++) {
    const r = toReparse[i];
    process.stdout.write(`\r  ${i + 1}/${toReparse.length} ${r.doc_id}`);

    try {
      await sleep(1000); // Rate limit
      const buf = await downloadDoc(r.doc_id);

      // Check if it's a valid ZIP
      if (buf.length < 100 || buf.readUInt32LE(0) !== 0x04034b50) {
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
    } catch (e) {
      // Skip errors silently
    }
  }

  console.log(`\n\nResults:`);
  console.log(`  Total re-parsed: ${toReparse.length}`);
  console.log(`  Records fixed: ${fixed}`);
  console.log(`  Ratios fixed: ${ratioFixed}`);
  console.log(`  Purposes fixed: ${purposeFixed}`);

  // Save
  const now = new Date();
  const jstOffset = 9 * 60 * 60 * 1000;
  const jstDate = new Date(now.getTime() + jstOffset);
  data.last_updated = jstDate.toISOString().replace('Z', '+09:00');

  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
  console.log(`\nSaved to ${DATA_FILE}`);
}

main().catch(e => { console.error(e); process.exit(1); });
