#!/usr/bin/env node
/**
 * reports.json の欠損データを EDINET API から補完するスクリプト。
 * 1. EDINETコードリスト(CSV)を取得 → issuer_name(Eコード) → {name, sec_code}
 * 2. sec_code → target_company のクロスリファレンス補完
 * 3. sec_codeなし or holding_ratioなしの報告に対してXBRL ZIPを再取得・解析
 */
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

const API_KEY = process.env.EDINET_API_KEY || '';
const EDINET_API_BASE = 'https://api.edinet-fsa.go.jp/api/v2';
const BASE = path.resolve(__dirname, '..');
const REPORTS_FILE = path.join(BASE, 'data', 'reports.json');

// Max XBRL downloads per run (rate limit: 1 req/sec)
const MAX_XBRL_DOWNLOADS = 500;

function normalizeWidth(text) {
  return text.replace(/[\uFF01-\uFF5E]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
  ).replace(/\u3000/g, ' ');
}

function extractSecCode(raw) {
  if (!raw) return '';
  const code = String(raw).trim();
  return code.length === 5 ? code.slice(0, 4) : code;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function fetchBuffer(url) {
  return new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? https : http;
    mod.get(url, { timeout: 60000 }, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        return fetchBuffer(res.headers.location).then(resolve, reject);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ── EDINET Code List ──
async function loadEdinetCodeMap() {
  console.log('EDINETコードリスト取得中...');
  const url = 'https://disclosure2dl.edinet-fsa.go.jp/searchdocument/codelist/Edinetcode.zip';
  const buf = await fetchBuffer(url);

  // The response is a ZIP containing a CSV - parse manually with built-in zlib

  // ZIP file parsing (minimal implementation)
  const entries = parseZipEntries(buf);
  const csvEntry = entries.find(e => e.name.endsWith('.csv'));
  if (!csvEntry) {
    console.log('  CSV not found in ZIP, trying as raw CSV...');
    return parseEdinetCsv(buf.toString('utf-8'));
  }

  // CSV is Shift-JIS encoded
  const { TextDecoder } = require('util');
  const csvText = new TextDecoder('shift-jis').decode(csvEntry.data);
  return parseEdinetCsv(csvText);
}

function parseZipEntries(buf) {
  const entries = [];
  let offset = 0;
  while (offset < buf.length - 4) {
    const sig = buf.readUInt32LE(offset);
    if (sig !== 0x04034b50) break; // Not a local file header

    const compressionMethod = buf.readUInt16LE(offset + 8);
    const compressedSize = buf.readUInt32LE(offset + 18);
    const uncompressedSize = buf.readUInt32LE(offset + 22);
    const nameLen = buf.readUInt16LE(offset + 26);
    const extraLen = buf.readUInt16LE(offset + 28);
    const name = buf.slice(offset + 30, offset + 30 + nameLen).toString('utf-8');
    const dataStart = offset + 30 + nameLen + extraLen;
    const rawData = buf.slice(dataStart, dataStart + compressedSize);

    let data;
    if (compressionMethod === 0) {
      data = rawData;
    } else if (compressionMethod === 8) {
      try {
        data = require('zlib').inflateRawSync(rawData);
      } catch (e) {
        data = rawData;
      }
    } else {
      data = rawData;
    }

    entries.push({ name, data });
    offset = dataStart + compressedSize;
  }
  return entries;
}

function parseEdinetCsv(csvText) {
  const lines = csvText.split('\n');
  const map = {};
  // Skip header (first line is download date, second is column headers)
  for (let i = 2; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    // CSV format: EDINETコード,提出者種別,上場区分,連結の有無,資本金,決算日,提出者名,提出者名（英字）,提出者名（カナ）,所在地,提出者業種,証券コード,提出者法人番号
    const cols = parseCsvLine(line);
    const edinetCode = (cols[0] || '').trim();
    const companyName = (cols[6] || '').trim();
    const secCode = extractSecCode((cols[11] || '').trim());

    if (edinetCode && edinetCode.startsWith('E')) {
      map[edinetCode] = { name: companyName, sec_code: secCode };
    }
  }
  console.log(`  EDINETコードリスト: ${Object.keys(map).length} 件`);
  return map;
}

function parseCsvLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') {
        current += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        result.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  result.push(current);
  return result;
}

// ── XBRL Download & Parse ──
async function downloadAndParseXbrl(docId) {
  const url = `${EDINET_API_BASE}/documents/${docId}?type=1&Subscription-Key=${API_KEY}`;
  let buf;
  try {
    buf = await fetchBuffer(url);
  } catch (e) {
    return {};
  }

  // Check if it's a ZIP
  if (buf.length < 4 || buf.readUInt32LE(0) !== 0x04034b50) {
    return {};
  }

  const entries = parseZipEntries(buf);
  const result = {};

  for (const entry of entries) {
    if (!entry.name.endsWith('.xbrl') && !entry.name.endsWith('.htm') && !entry.name.endsWith('.html')) continue;

    let text;
    try {
      text = entry.data.toString('utf-8');
    } catch (e) {
      continue;
    }

    // Target company name
    const issuerPatterns = [
      /name="[^"]*(?:[Ii]ssuer[Nn]ame|NameOfIssuer|IssuerNameJp)[^"]*"[^>]*>([^<]+)/,
      /発行者の名称[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})/,
      /発行者の名称.*?<[^>]*>\s*([^\n<]{2,40}?)\s*</,
      /株券等の発行者[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})/
    ];
    for (const pat of issuerPatterns) {
      const m = text.match(pat);
      if (m) {
        const name = m[1].trim();
        if (name && name.length >= 2 && !name.includes('報告書') && !name.includes('提出者') && !name.includes('代表取締役')) {
          result.target_company = name;
          break;
        }
      }
    }

    // Securities code
    const codePatterns = [
      /(?:証券コード|銘柄コード)[^\dA-Za-z]{0,10}([\dA-Za-z]{4,5})/,
      /name="[^"]*(?:[Ss]ecurity[Cc]ode|SecuritiesCode)[^"]*"[^>]*>([\dA-Za-z]{4,5})/
    ];
    for (const pat of codePatterns) {
      const m = text.match(pat);
      if (m) {
        result.sec_code = m[1].trim();
        break;
      }
    }

    // Holding ratio
    const ratioPatterns = [
      /(?:保有割合|所有割合)[^\d]{0,30}?([\d]+[\.．][\d]+)\s*[%％]/,
      /name="[^"]*(?:HoldingRatio|OwnershipRatio)[^"]*"[^>]*>([\d]+[\.．][\d]+)/,
      /([\d]+[\.．][\d]+)\s*[%％]\s*(?:（.*?保有割合|を保有)/
    ];
    for (const pat of ratioPatterns) {
      const m = text.match(pat);
      if (m) {
        const ratioStr = m[1].replace('．', '.');
        const val = parseFloat(ratioStr);
        if (!isNaN(val)) {
          result.holding_ratio = val;
        }
        break;
      }
    }

    // Purpose
    const purposePatterns = [
      /保有目的[^\n]{0,5}[：:]\s*([^\n<]{2,60})/
    ];
    for (const pat of purposePatterns) {
      const m = text.match(pat);
      if (m) {
        const raw = m[1].trim();
        result.purpose = classifyPurpose(raw);
        result.purpose_detail = raw.slice(0, 100);
        break;
      }
    }

    if (Object.keys(result).length > 0) break;
  }

  return result;
}

function classifyPurpose(text) {
  if (text.includes('純投資')) return '純投資';
  if (text.includes('政策')) return '政策投資';
  if (text.includes('株主提案') || text.includes('提案')) return '株主提案';
  if (text.includes('経営') || text.includes('支配') || text.includes('関与')) return '経営関与';
  if (text.includes('重要提案行為')) return '重要提案';
  return 'その他';
}

// ── Main ──
async function main() {
  console.log('=== reports.json データ補完スクリプト ===\n');

  const data = JSON.parse(fs.readFileSync(REPORTS_FILE, 'utf8'));
  const reports = data.reports;
  console.log(`総報告数: ${reports.length}`);

  // Step 1: EDINETコードリストで補完
  let edinetMap = {};
  try {
    edinetMap = await loadEdinetCodeMap();
  } catch (e) {
    console.log(`  EDINETコードリスト取得失敗: ${e.message}`);
  }

  let patchedByCode = 0;
  if (Object.keys(edinetMap).length > 0) {
    for (const r of reports) {
      const issuerCode = (r.issuer_name || '').trim();
      if (issuerCode && issuerCode.startsWith('E')) {
        const info = edinetMap[issuerCode];
        if (info) {
          if (!r.target_company && info.name) {
            r.target_company = info.name;
            patchedByCode++;
          }
          if (!r.sec_code && info.sec_code) {
            r.sec_code = info.sec_code;
            patchedByCode++;
          }
        }
      }
    }
    console.log(`EDINETコードリスト補完: ${patchedByCode} フィールド`);
  }

  // Step 2: sec_code → target_company クロスリファレンス
  const secToName = {};
  for (const r of reports) {
    const sc = (r.sec_code || '').trim();
    const tc = (r.target_company || '').trim();
    if (sc && tc && !secToName[sc]) {
      secToName[sc] = tc;
    }
  }
  let patchedByXref = 0;
  for (const r of reports) {
    const sc = (r.sec_code || '').trim();
    if (sc && !(r.target_company || '').trim() && secToName[sc]) {
      r.target_company = secToName[sc];
      patchedByXref++;
    }
  }
  console.log(`クロスリファレンス補完: ${patchedByXref} 件`);

  // Step 3: XBRL再取得（sec_codeなし or holding_ratioなし）
  // Prioritize: recent reports first, then by importance
  const incomplete = reports.filter(r => {
    if (!r.doc_id) return false;
    // Missing sec_code entirely
    if (!r.sec_code) return true;
    // Has sec_code but missing both target_company and holding_ratio
    if (!r.target_company && r.holding_ratio == null) return true;
    return false;
  });

  // Sort by date descending (newest first)
  incomplete.sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const toProcess = incomplete.slice(0, MAX_XBRL_DOWNLOADS);
  console.log(`\nXBRL再取得対象: ${incomplete.length} 件中 ${toProcess.length} 件を処理`);

  let xbrlPatched = 0;
  let xbrlFailed = 0;
  for (let i = 0; i < toProcess.length; i++) {
    const r = toProcess[i];
    process.stdout.write(`\r  XBRL取得中: ${i + 1}/${toProcess.length} (補完: ${xbrlPatched}, 失敗: ${xbrlFailed})`);

    try {
      await sleep(1000); // Rate limit
      const xbrl = await downloadAndParseXbrl(r.doc_id);

      if (xbrl.target_company && !r.target_company) {
        // Skip if target == filer
        const norm = normalizeWidth(xbrl.target_company);
        const filerNorm = normalizeWidth(r.filer_name || '');
        if (!norm.includes(filerNorm) && !filerNorm.includes(norm)) {
          r.target_company = xbrl.target_company;
        }
      }
      if (xbrl.sec_code && !r.sec_code) {
        r.sec_code = extractSecCode(xbrl.sec_code);
      }
      if (xbrl.holding_ratio != null && r.holding_ratio == null) {
        r.holding_ratio = xbrl.holding_ratio;
      }
      if (xbrl.purpose && !r.purpose) {
        r.purpose = xbrl.purpose;
      }
      if (xbrl.purpose_detail && !r.purpose_detail) {
        r.purpose_detail = xbrl.purpose_detail;
      }

      if (Object.keys(xbrl).length > 0) xbrlPatched++;
      else xbrlFailed++;
    } catch (e) {
      xbrlFailed++;
    }
  }
  console.log(`\nXBRL補完完了: ${xbrlPatched} 件成功, ${xbrlFailed} 件失敗`);

  // Step 4: 再度クロスリファレンス（XBRL取得で新たに判明したsec_code→target_company）
  const secToName2 = {};
  for (const r of reports) {
    const sc = (r.sec_code || '').trim();
    const tc = (r.target_company || '').trim();
    if (sc && tc && !secToName2[sc]) {
      secToName2[sc] = tc;
    }
  }
  let patchedByXref2 = 0;
  for (const r of reports) {
    const sc = (r.sec_code || '').trim();
    if (sc && !(r.target_company || '').trim() && secToName2[sc]) {
      r.target_company = secToName2[sc];
      patchedByXref2++;
    }
  }
  if (patchedByXref2) console.log(`追加クロスリファレンス補完: ${patchedByXref2} 件`);

  // Stats
  let hasCodeTarget = 0, hasCodeNoTarget = 0, noCode = 0;
  for (const r of reports) {
    if (r.sec_code && r.target_company) hasCodeTarget++;
    else if (r.sec_code) hasCodeNoTarget++;
    else noCode++;
  }
  console.log(`\n--- 補完後の状態 ---`);
  console.log(`sec_code + target_company あり: ${hasCodeTarget}`);
  console.log(`sec_code あり / target なし: ${hasCodeNoTarget}`);
  console.log(`sec_code なし: ${noCode}`);

  // Save
  data.last_updated = new Date().toISOString();
  fs.writeFileSync(REPORTS_FILE, JSON.stringify(data, null, 2), 'utf8');
  console.log(`\n保存完了: ${REPORTS_FILE}`);
}

main().catch(e => { console.error(e); process.exit(1); });
