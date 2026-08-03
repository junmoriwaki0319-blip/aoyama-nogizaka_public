#!/usr/bin/env node
/**
 * 外食産業セクター指数 & TOPIX月次データを再生成し Firestore にアップロード
 *
 * v2 (2026-07): update-saas-market-index.js と同設計に全面改修。旧版の問題点を解消:
 *   - 旧: 構成15社・ウェイトをスクリプト内ハードコード → 新: Firestore premiumContent/food-companies
 *     から「更新時点の時価総額上位15社」を自動選定（時価総額加重、shares = marketCap/stockPrice 固定株数）。
 *     ハードコードと違い銘柄リストが陳腐化せず、選定基準（時価総額上位）が機械的で説明可能。
 *     上位15社で収録全社の時価総額ウェイト約8割をカバーする（ページ側の指数定義の記載と対応）
 *   - 旧: TOPIX=1306.T月次のみ（約1/10グリッチが混入、2026-02〜04の系列歪みの原因）
 *     → 新: TOPIX連動ETF 3本（1306/1305/1308）の日次月末終値を正規化し中央値
 *   - 旧: フィールド全置換PATCH（SECTOR_AVG_SSS_* が消える事故源）→ 新: updateMask付きPATCH
 *   - 旧: 15ヶ月ローリング窓（実行のたびに基準月がズレる）→ 新: 既存INDEX_MONTHSの
 *     基準月（24/06）を維持して全期間を再生成（外食docにはidx紐付け注釈が無いため安全。
 *     ページの「基準月=100」表記はINDEX_MONTHS[0]からデータ連動で描画される）
 *   - 銘柄月次の異常値は日次月末終値で裏取り（実際の急騰・急落は保持）
 *
 * 使い方:
 *   node scripts/update-food-market-index.js           # dry-run（計算結果と既存系列との差分表示）
 *   node scripts/update-food-market-index.js --apply   # Firestoreへ反映
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const PROJECT_ID = 'aoyama-nogizaka-activist';
const APPLY = process.argv.includes('--apply');

// ---------- HTTP ----------

function fetchJSON(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.get({
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36', ...headers },
      timeout: 20000,
    }, resp => {
      let data = '';
      resp.on('data', c => data += c);
      resp.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve(null); } });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout: ' + url)); });
  });
}

// ---------- Yahoo Finance ----------

// 月次終値（adjclose優先: 分割・併合対応）。同月重複は後勝ち（進行月は最新値）。
async function fetchMonthly(ticker, period1) {
  const p1 = Math.floor(period1.getTime() / 1000);
  const p2 = Math.floor(Date.now() / 1000);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1mo&period1=${p1}&period2=${p2}`;
  const data = await fetchJSON(url);
  const result = data?.chart?.result?.[0];
  if (!result || !result.timestamp) return null;
  const ts = result.timestamp;
  const closes = result.indicators?.quote?.[0]?.close || [];
  const adj = result.indicators?.adjclose?.[0]?.adjclose || [];
  const points = [];
  for (let i = 0; i < ts.length; i++) {
    const c = adj[i] != null ? adj[i] : closes[i];
    if (c == null) continue;
    const d = new Date(ts[i] * 1000);
    const label = `${String(d.getFullYear()).slice(2)}/${String(d.getMonth() + 1).padStart(2, '0')}`;
    const last = points[points.length - 1];
    if (last && last.label === label) last.close = c;
    else points.push({ label, close: c });
  }
  return points;
}

// 日次データから各月の月末終値を導出（label→月内最終取引日の終値）
async function fetchDailyMonthEnds(ticker, period1) {
  const p1 = Math.floor(period1.getTime() / 1000);
  const p2 = Math.floor(Date.now() / 1000);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1d&period1=${p1}&period2=${p2}`;
  const data = await fetchJSON(url);
  const result = data?.chart?.result?.[0];
  if (!result || !result.timestamp) return null;
  const closes = result.indicators?.quote?.[0]?.close || [];
  const byMonth = {};
  for (let i = 0; i < result.timestamp.length; i++) {
    if (closes[i] == null) continue;
    const d = new Date(result.timestamp[i] * 1000);
    byMonth[`${String(d.getFullYear()).slice(2)}/${String(d.getMonth() + 1).padStart(2, '0')}`] = closes[i];
  }
  return Object.entries(byMonth).map(([label, close]) => ({ label, close }));
}

// 日次ローリング1年リターン（TOPIX_RETURN_1Y用）
async function fetchReturn1Y(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1d&range=1y`;
  const data = await fetchJSON(url);
  const result = data?.chart?.result?.[0];
  const closes = (result?.indicators?.quote?.[0]?.close || []).filter(c => c != null);
  if (closes.length < 2) return null;
  return Math.round((closes[closes.length - 1] / closes[0] - 1) * 1000) / 10;
}

// 外れ値検証: 隣接月と±40%超乖離した点は「日次データの月末終値」で裏取りして置換
function detectSuspects(points) {
  const suspects = [];
  for (let i = 1; i < points.length; i++) {
    const prev = points[i - 1].close, cur = points[i].close;
    const next = i + 1 < points.length ? points[i + 1].close : null;
    const devPrev = Math.abs(cur / prev - 1) > 0.4;
    const devNext = next != null ? Math.abs(cur / next - 1) > 0.4 : true;
    if (devPrev && devNext) suspects.push(i);
  }
  return suspects;
}

async function verifyOutliersDaily(ticker, points, name, period1) {
  const suspects = detectSuspects(points);
  if (!suspects.length) return points;
  const daily = await fetchDailyMonthEnds(ticker, period1);
  const monthEnd = {};
  for (const p of daily || []) monthEnd[p.label] = p.close;
  for (const i of suspects) {
    const d = monthEnd[points[i].label];
    if (d == null) { console.log(`  ⚠ ${name} ${points[i].label}: 疑値 ${points[i].close.toFixed(1)} (日次で検証不可のため保持)`); continue; }
    if (Math.abs(d / points[i].close - 1) > 0.05) {
      console.log(`  ⚠ ${name} ${points[i].label}: 月次バー異常 ${points[i].close.toFixed(1)} → 日次月末終値 ${d.toFixed(1)} に置換`);
      points[i].close = d;
    } else {
      console.log(`  ✓ ${name} ${points[i].label}: 急変 ${points[i].close.toFixed(1)} は日次と一致（実際の値動きとして保持）`);
    }
  }
  return points;
}

// ---------- Firebase ----------

function getAccessToken() {
  const configPath = path.join(process.env.HOME || process.env.USERPROFILE, '.config', 'configstore', 'firebase-tools.json');
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const refreshToken = config.tokens.refresh_token;
  return new Promise((resolve, reject) => {
    const postData = `grant_type=refresh_token&refresh_token=${encodeURIComponent(refreshToken)}&client_id=563584335869-fgrhgmd47bqnekij5i8b5pr03ho849e6.apps.googleusercontent.com&client_secret=j9iVZfS8kkCEFUPaAeJV0sAi`;
    const req = https.request({
      hostname: 'oauth2.googleapis.com', path: '/token', method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const json = JSON.parse(data);
        json.access_token ? resolve(json.access_token) : reject(new Error('Token refresh failed: ' + data));
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

function fromFirestoreValue(v) {
  if (v.nullValue !== undefined) return null;
  if (v.stringValue !== undefined) return v.stringValue;
  if (v.integerValue !== undefined) return parseInt(v.integerValue);
  if (v.doubleValue !== undefined) return v.doubleValue;
  if (v.booleanValue !== undefined) return v.booleanValue;
  if (v.arrayValue !== undefined) return (v.arrayValue.values || []).map(fromFirestoreValue);
  if (v.mapValue !== undefined) {
    const out = {};
    for (const [k, val] of Object.entries(v.mapValue.fields || {})) out[k] = fromFirestoreValue(val);
    return out;
  }
  return null;
}

function toFirestoreValue(val) {
  if (val === null || val === undefined) return { nullValue: null };
  if (typeof val === 'boolean') return { booleanValue: val };
  if (typeof val === 'number') return Number.isInteger(val) ? { integerValue: String(val) } : { doubleValue: val };
  if (typeof val === 'string') return { stringValue: val };
  if (Array.isArray(val)) return { arrayValue: { values: val.map(toFirestoreValue) } };
  if (typeof val === 'object') {
    const fields = {};
    for (const [k, v] of Object.entries(val)) fields[k] = toFirestoreValue(v);
    return { mapValue: { fields } };
  }
  return { stringValue: String(val) };
}

function firestoreRead(token, docId) {
  return new Promise((resolve, reject) => {
    https.get({
      hostname: 'firestore.googleapis.com',
      path: `/v1/projects/${PROJECT_ID}/databases/(default)/documents/premiumContent/${docId}`,
      headers: { Authorization: `Bearer ${token}` },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) return reject(new Error(`read ${docId} failed (${res.statusCode})`));
        const doc = JSON.parse(data);
        const out = {};
        for (const [k, v] of Object.entries(doc.fields || {})) out[k] = fromFirestoreValue(v);
        resolve(out);
      });
    }).on('error', reject);
  });
}

// updateMask付きPATCH — 指定フィールド以外（SECTOR_AVG_SSS_* 等）には触れない
function firestorePatchMasked(token, docId, data) {
  const mask = Object.keys(data).map(k => 'updateMask.fieldPaths=' + encodeURIComponent(k)).join('&');
  const fields = {};
  for (const [k, v] of Object.entries(data)) fields[k] = toFirestoreValue(v);
  const body = JSON.stringify({ fields });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'firestore.googleapis.com',
      path: `/v1/projects/${PROJECT_ID}/databases/(default)/documents/premiumContent/${docId}?${mask}`,
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, (res) => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve();
        else reject(new Error(`PATCH ${docId} failed (${res.statusCode}): ${d.slice(0, 300)}`));
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ---------- 指数算出 ----------

function labelToDate(label) { // '24/06' → Date(2024-06-01)
  const [y, m] = label.split('/');
  return new Date(2000 + parseInt(y), parseInt(m) - 1, 1);
}

async function main() {
  console.log(`=== update-food-market-index v2 ${APPLY ? '(APPLY)' : '(dry-run)'} ===\n`);
  const token = await getAccessToken();

  // 1. 既存データ取得（基準月の特定と差分レポート用）
  const market = await firestoreRead(token, 'food-market');
  const stored = {
    months: market.INDEX_MONTHS || [],
    rest: market.RESTAURANT_INDEX_MONTHLY || [],
    topix: market.TOPIX_MONTHLY || [],
  };
  if (!stored.months.length) throw new Error('既存INDEX_MONTHSが空です（基準月を特定できないため中止）');
  const baseLabel = stored.months[0];
  console.log(`既存シリーズ: ${stored.months.length}点 (${baseLabel} ～ ${stored.months[stored.months.length - 1]})`);

  const companiesDoc = await firestoreRead(token, 'food-companies');
  const allCompanies = (companiesDoc.companies || []).filter(c => typeof c.marketCap === 'number' && c.marketCap > 0);
  if (allCompanies.length < 15) throw new Error('収録銘柄が少なすぎます');
  // 更新時点の時価総額上位15社を自動選定（弊社収録ユニバースの代表指数）
  const TOP_N = 15;
  const companies = [...allCompanies].sort((a, b) => b.marketCap - a.marketCap).slice(0, TOP_N);
  const totalMcap = allCompanies.reduce((a, c) => a + c.marketCap, 0);
  const topMcap = companies.reduce((a, c) => a + c.marketCap, 0);
  console.log(`構成銘柄: 収録${allCompanies.length}社中 時価総額上位${TOP_N}社を自動選定（全体ウェイトの${(topMcap / totalMcap * 100).toFixed(1)}%をカバー）`);
  companies.forEach((c, i) => console.log(`  ${String(i + 1).padStart(2)}. ${c.code} ${c.name} (${(c.marketCap / totalMcap * 100).toFixed(1)}%)`));
  console.log('');

  // 2. 月次データ取得（基準月の前月15日から）
  const period1 = labelToDate(baseLabel);
  period1.setDate(period1.getDate() - 15);

  // TOPIX月次: ETF 3本の日次月末終値を正規化し中央値（1306.T単独のグリッチ対策）
  console.log('=== TOPIX月次取得 (1306/1305/1308 日次月末の中央値) ===');
  const etfs = [];
  for (const sym of ['1306.T', '1305.T', '1308.T']) {
    const pts = await fetchDailyMonthEnds(sym, period1);
    if (pts && pts.find(p => p.label === baseLabel)) etfs.push({ sym, pts });
    else console.log(`  ✗ ${sym}: 取得失敗または基準月なし`);
  }
  if (etfs.length < 2) throw new Error('TOPIX ETFが2本未満しか取れません（中央値の信頼性不足のため中止）');
  const allLabels = [...new Set(etfs.flatMap(e => e.pts.map(p => p.label)))]
    .filter(l => labelToDate(l) >= labelToDate(baseLabel))
    .sort((a, b) => labelToDate(a) - labelToDate(b));
  const median = a => { const s = [...a].sort((x, y) => x - y); return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2; };
  const topixRatio = allLabels.map(label => {
    const vals = etfs.map(e => {
      const base = e.pts.find(p => p.label === baseLabel);
      const p = e.pts.find(x => x.label === label);
      return p && base ? p.close / base.close : null;
    }).filter(v => v != null);
    return vals.length ? { label, ratio: median(vals) } : null;
  }).filter(Boolean);
  for (let i = 1; i < topixRatio.length; i++) {
    if (Math.abs(topixRatio[i].ratio / topixRatio[i - 1].ratio - 1) > 0.4) {
      console.log(`  ⚠ TOPIX ${topixRatio[i].label}: 中央値でも異常(前月比${(topixRatio[i].ratio / topixRatio[i - 1].ratio).toFixed(2)}x) → 前月値で代替`);
      topixRatio[i].ratio = topixRatio[i - 1].ratio;
    }
  }
  const months = topixRatio.map(p => p.label);
  console.log(`  ETF ${etfs.length}本 / ${months.length}点 (${months[0]} ～ ${months[months.length - 1]})\n`);

  // 外食構成銘柄の月次データ
  console.log('=== 外食構成銘柄 月次取得 ===');
  const stocks = [];
  const excluded = [];
  for (let i = 0; i < companies.length; i += 5) {
    const batch = companies.slice(i, i + 5);
    const results = await Promise.allSettled(batch.map(async c => {
      const pts = await fetchMonthly(`${c.code}.T`, period1);
      return { c, pts };
    }));
    for (const r of results) {
      if (r.status !== 'fulfilled' || !r.value.pts || r.value.pts.length < 2) {
        excluded.push(r.status === 'fulfilled' ? `${r.value.c.code} ${r.value.c.name}(取得失敗)` : '?');
        continue;
      }
      const { c, pts } = r.value;
      if (typeof c.marketCap !== 'number' || typeof c.stockPrice !== 'number' || c.stockPrice <= 0) {
        excluded.push(`${c.code} ${c.name}(marketCap/stockPrice欠損)`);
        continue;
      }
      const verified = await verifyOutliersDaily(`${c.code}.T`, pts, `${c.code} ${c.name}`, period1);
      const base = verified.find(p => p.label === baseLabel);
      if (!base) {
        excluded.push(`${c.code} ${c.name}(基準月${baseLabel}データ無し)`);
        continue;
      }
      stocks.push({ code: c.code, name: c.name, shares: c.marketCap / c.stockPrice, pts: verified, base });
    }
    await new Promise(r => setTimeout(r, 300));
  }
  console.log(`  → 有効 ${stocks.length}/${companies.length}社`);
  if (excluded.length) console.log(`  除外: ${excluded.join(' / ')}`);
  console.log('');
  if (stocks.length < 10) throw new Error('有効銘柄が少なすぎます');

  // 3. 時価総額加重指数（基準月=100）— 全期間を再生成
  const out = { INDEX_MONTHS: months, TOPIX_MONTHLY: [], RESTAURANT_INDEX_MONTHLY: [] };
  for (const { label, ratio } of topixRatio) {
    out.TOPIX_MONTHLY.push(Math.round(ratio * 1000) / 10);
    let v = 0, vb = 0;
    for (const s of stocks) {
      const p = s.pts.find(x => x.label === label);
      if (p) { v += s.shares * p.close; vb += s.shares * s.base.close; }
    }
    out.RESTAURANT_INDEX_MONTHLY.push(vb > 0 ? Math.round(v / vb * 1000) / 10 : null);
  }

  // 4. TOPIX_RETURN_1Y（日次ローリング1年）
  const topix1y = await fetchReturn1Y('1306.T');
  if (topix1y != null) out.TOPIX_RETURN_1Y = topix1y;

  // 5. 既存系列との差分レポート（再生成のため置換されるが、変化量を可視化）
  console.log('=== 既存系列との差分（月: 外食 旧→新 / TOPIX 旧→新）===');
  let maxRestDiff = 0, maxTopixDiff = 0;
  for (let i = 0; i < stored.months.length; i++) {
    const j = months.indexOf(stored.months[i]);
    if (j < 0) continue;
    const rd = out.RESTAURANT_INDEX_MONTHLY[j] != null && stored.rest[i] != null ? out.RESTAURANT_INDEX_MONTHLY[j] - stored.rest[i] : null;
    const td = out.TOPIX_MONTHLY[j] != null && stored.topix[i] != null ? out.TOPIX_MONTHLY[j] - stored.topix[i] : null;
    if (rd != null) maxRestDiff = Math.max(maxRestDiff, Math.abs(rd));
    if (td != null) maxTopixDiff = Math.max(maxTopixDiff, Math.abs(td));
    if ((rd != null && Math.abs(rd) > 2) || (td != null && Math.abs(td) > 2)) {
      console.log(`  ${stored.months[i]}: 外食 ${stored.rest[i]}→${out.RESTAURANT_INDEX_MONTHLY[j]} / TOPIX ${stored.topix[i]}→${out.TOPIX_MONTHLY[j]}`);
    }
  }
  console.log(`  最大乖離: 外食 ${maxRestDiff.toFixed(1)}pt / TOPIX ${maxTopixDiff.toFixed(1)}pt\n`);

  console.log('=== 更新内容 ===');
  console.log(`  INDEX_MONTHS(${out.INDEX_MONTHS.length}): ${out.INDEX_MONTHS.join(',')}`);
  console.log(`  RESTAURANT_INDEX_MONTHLY: ${out.RESTAURANT_INDEX_MONTHLY.join(',')}`);
  console.log(`  TOPIX_MONTHLY: ${out.TOPIX_MONTHLY.join(',')}`);
  console.log(`  TOPIX_RETURN_1Y: ${out.TOPIX_RETURN_1Y}%`);

  if (!APPLY) {
    console.log('\n(dry-run のため書き込みなし。反映するには --apply を付けて実行)');
    return;
  }
  await firestorePatchMasked(token, 'food-market', out);
  console.log('\n=== premiumContent/food-market を更新しました（updateMask付きPATCH、SECTOR_AVG_SSS_*は保持） ===');
}

main().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
