#!/usr/bin/env node
/**
 * SaaSセクター指数 & TOPIX月次データを延長更新し Firestore にアップロード
 *
 * 設計方針（update-food-market-index.js の反省を反映）:
 *   - 基準月(既存INDEX_MONTHSの先頭、25/01)を維持したまま末尾に月を追加する「延長方式」。
 *     SAAS_EVENTSの注釈が series index (idx) で紐付いているため、窓のローリングは厳禁。
 *   - 構成銘柄はスクリプト内ハードコードではなく Firestore premiumContent/saas-companies の
 *     30社から動的取得（ユニバースドリフト防止）。時価総額加重（shares = marketCap/stockPrice 固定株数）。
 *   - Firestore書込は updateMask 付きPATCH（SAAS_EVENTS / QUARTERS を消さない）。
 *   - 1306.T 月次データの外れ値（隣接比±40%超）は近傍補間で補正。
 *   - TOPIX_RETURN_1Y は日次ローリング1年（stockReturn1Yと同一手法）。
 *   - 既存シリーズとの重複期間を検証し、乖離が大きい場合は既存値を保持して
 *     月次リターンのチェーン連結で延長（履歴の書き換えを回避）。
 *
 * 使い方:
 *   node scripts/update-saas-market-index.js           # dry-run（計算結果の表示のみ）
 *   node scripts/update-saas-market-index.js --apply   # Firestoreへ反映
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const PROJECT_ID = 'aoyama-nogizaka-activist';
const APPLY = process.argv.includes('--apply');
const OVERLAP_TOLERANCE = 3.0; // pt。これを超えたらチェーン延長にフォールバック

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

// 月次終値（adjclose優先: 分割・併合対応）。period1以降の各月1ポイント。
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
    // 同月が複数返った場合は後勝ち（月内の最新値を採用）
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
// （1306.T月次の既知の約1/10グリッチ対策。実際の急騰・急落は日次と一致するので保持される）
async function verifyOutliersDaily(ticker, points, name, period1) {
  const suspects = [];
  for (let i = 1; i < points.length; i++) {
    const prev = points[i - 1].close, cur = points[i].close;
    const next = i + 1 < points.length ? points[i + 1].close : null;
    const devPrev = Math.abs(cur / prev - 1) > 0.4;
    const devNext = next != null ? Math.abs(cur / next - 1) > 0.4 : true;
    if (devPrev && devNext) suspects.push(i);
  }
  if (!suspects.length) return points;
  // 日次終値を一括取得し、各月の最終値を正とする
  const p1 = Math.floor(period1.getTime() / 1000);
  const p2 = Math.floor(Date.now() / 1000);
  const data = await fetchJSON(`https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1d&period1=${p1}&period2=${p2}`);
  const result = data?.chart?.result?.[0];
  const monthEnd = {};
  if (result && result.timestamp) {
    const closes = result.indicators?.quote?.[0]?.close || [];
    for (let i = 0; i < result.timestamp.length; i++) {
      if (closes[i] == null) continue;
      const d = new Date(result.timestamp[i] * 1000);
      monthEnd[`${String(d.getFullYear()).slice(2)}/${String(d.getMonth() + 1).padStart(2, '0')}`] = closes[i];
    }
  }
  for (const i of suspects) {
    const daily = monthEnd[points[i].label];
    if (daily == null) { console.log(`  ⚠ ${name} ${points[i].label}: 疑値 ${points[i].close.toFixed(1)} (日次で検証不可のため保持)`); continue; }
    if (Math.abs(daily / points[i].close - 1) > 0.05) {
      console.log(`  ⚠ ${name} ${points[i].label}: 月次バー異常 ${points[i].close.toFixed(1)} → 日次月末終値 ${daily.toFixed(1)} に置換`);
      points[i].close = daily;
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

// updateMask付きPATCH — 指定フィールド以外（SAAS_EVENTS / QUARTERS等）には触れない
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

function labelToDate(label) { // '25/01' → Date(2025-01-01)
  const [y, m] = label.split('/');
  return new Date(2000 + parseInt(y), parseInt(m) - 1, 1);
}

async function main() {
  console.log(`=== update-saas-market-index ${APPLY ? '(APPLY)' : '(dry-run)'} ===\n`);
  const token = await getAccessToken();

  // 1. 既存データ取得
  const market = await firestoreRead(token, 'saas-market');
  const stored = {
    months: market.INDEX_MONTHS || [],
    saas: market.SAAS_INDEX_MONTHLY || [],
    topix: market.TOPIX_MONTHLY || [],
  };
  if (!stored.months.length) throw new Error('既存INDEX_MONTHSが空です（基準月を特定できないため中止）');
  const baseLabel = stored.months[0];
  console.log(`既存シリーズ: ${stored.months.length}点 (${baseLabel} ～ ${stored.months[stored.months.length - 1]})`);

  const companiesDoc = await firestoreRead(token, 'saas-companies');
  const companies = companiesDoc.companies || [];
  console.log(`構成銘柄: saas-companies から ${companies.length}社\n`);
  if (companies.length < 10) throw new Error('構成銘柄が少なすぎます');

  // 2. 月次データ取得（基準月の前月15日から）
  const period1 = labelToDate(baseLabel);
  period1.setDate(period1.getDate() - 15);

  // TOPIX月次: 1306.T単独はYahooデータに約1/10のグリッチ日が混入するため、
  // TOPIX連動ETF 3本の「日次データから導出した各月末終値」を正規化し中央値を採用。
  // シンボル固有の異常は中央値で除去され、進行中の月は最新日次終値になる。
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
  const topixRaw = allLabels.map(label => {
    const vals = etfs.map(e => {
      const base = e.pts.find(p => p.label === baseLabel);
      const p = e.pts.find(x => x.label === label);
      return p && base ? p.close / base.close : null;
    }).filter(v => v != null);
    return vals.length ? { label, close: median(vals) } : null; // close=対基準倍率
  }).filter(Boolean);
  // 中央値でも残る異常（全ETF同時グリッチ）を最終ガード: 前月比±40%超は前月値で埋めて警告
  for (let i = 1; i < topixRaw.length; i++) {
    if (Math.abs(topixRaw[i].close / topixRaw[i - 1].close - 1) > 0.4) {
      console.log(`  ⚠ TOPIX ${topixRaw[i].label}: 中央値でも異常(前月比${(topixRaw[i].close / topixRaw[i - 1].close).toFixed(2)}x) → 前月値で代替`);
      topixRaw[i].close = topixRaw[i - 1].close;
    }
  }
  const months = topixRaw.map(p => p.label);
  console.log(`  ETF ${etfs.length}本 / ${months.length}点 (${months[0]} ～ ${months[months.length - 1]})\n`);

  console.log('=== SaaS構成銘柄 月次取得 ===');
  const stocks = [];
  for (let i = 0; i < companies.length; i += 5) {
    const batch = companies.slice(i, i + 5);
    const results = await Promise.allSettled(batch.map(async c => {
      const pts = await fetchMonthly(`${c.code}.T`, period1);
      return { c, pts };
    }));
    for (const r of results) {
      if (r.status !== 'fulfilled' || !r.value.pts || r.value.pts.length < 2) {
        console.log(`  ✗ ${r.status === 'fulfilled' ? r.value.c.code + ' ' + r.value.c.name : '?'}: 取得失敗（指数から除外）`);
        continue;
      }
      const { c, pts } = r.value;
      const base = pts.find(p => p.label === baseLabel);
      if (!base) {
        console.log(`  ✗ ${c.code} ${c.name}: 基準月${baseLabel}のデータ無し（指数から除外）`);
        continue;
      }
      // 固定株数 = 現在時価総額 / 現在株価（分割はadjcloseで吸収）
      if (typeof c.marketCap !== 'number' || typeof c.stockPrice !== 'number' || c.stockPrice <= 0) {
        console.log(`  ✗ ${c.code} ${c.name}: marketCap/stockPrice欠損（指数から除外）`);
        continue;
      }
      const verified = await verifyOutliersDaily(`${c.code}.T`, pts, `${c.code} ${c.name}`, period1);
      const vbase = verified.find(p => p.label === baseLabel);
      stocks.push({ code: c.code, name: c.name, shares: c.marketCap / c.stockPrice, pts: verified, base: vbase });
    }
    await new Promise(r => setTimeout(r, 300));
  }
  console.log(`  → 有効 ${stocks.length}/${companies.length}社\n`);
  if (stocks.length < 10) throw new Error('有効銘柄が少なすぎます');

  // 3. 時価総額加重指数（基準月=100）
  const topixBase = topixRaw.find(p => p.label === baseLabel);
  if (!topixBase) throw new Error(`TOPIXに基準月${baseLabel}がありません`);
  const computed = { months, saas: [], topix: [] };
  for (const label of months) {
    const tp = topixRaw.find(p => p.label === label);
    computed.topix.push(tp ? Math.round(tp.close / topixBase.close * 1000) / 10 : null);
    let v = 0, vb = 0;
    for (const s of stocks) {
      const p = s.pts.find(x => x.label === label);
      if (p) { v += s.shares * p.close; vb += s.shares * s.base.close; }
    }
    computed.saas.push(vb > 0 ? Math.round(v / vb * 1000) / 10 : null);
  }

  // 4. 既存シリーズとの整合検証（重複期間の乖離）
  let maxDiff = 0;
  for (let i = 0; i < stored.months.length; i++) {
    const j = computed.months.indexOf(stored.months[i]);
    if (j < 0 || computed.saas[j] == null) continue;
    maxDiff = Math.max(maxDiff, Math.abs(computed.saas[j] - stored.saas[i]));
  }
  console.log(`=== 整合検証: 重複期間の最大乖離 ${maxDiff.toFixed(1)}pt (許容 ${OVERLAP_TOLERANCE}pt) ===`);

  let out;
  if (maxDiff <= OVERLAP_TOLERANCE) {
    console.log('  → 全期間を再計算値で更新（履歴込みで整合）\n');
    out = { INDEX_MONTHS: computed.months, SAAS_INDEX_MONTHLY: computed.saas, TOPIX_MONTHLY: computed.topix };
  } else {
    console.log('  → 乖離大: 既存値を保持し、月次リターンのチェーン連結で延長\n');
    const lastStored = stored.months[stored.months.length - 1];
    const li = computed.months.indexOf(lastStored);
    if (li < 0) throw new Error(`再計算シリーズに既存最終月${lastStored}がありません`);
    const outMonths = [...stored.months], outSaas = [...stored.saas], outTopix = [...stored.topix];
    for (let j = li + 1; j < computed.months.length; j++) {
      const rSaas = computed.saas[j] / computed.saas[j - 1];
      const rTopix = computed.topix[j] / computed.topix[j - 1];
      outMonths.push(computed.months[j]);
      outSaas.push(Math.round(outSaas[outSaas.length - 1] * rSaas * 10) / 10);
      outTopix.push(Math.round(outTopix[outTopix.length - 1] * rTopix * 10) / 10);
    }
    out = { INDEX_MONTHS: outMonths, SAAS_INDEX_MONTHLY: outSaas, TOPIX_MONTHLY: outTopix };
  }

  // 5. TOPIX_RETURN_1Y（日次ローリング1年）
  const topix1y = await fetchReturn1Y('1306.T');
  if (topix1y != null) out.TOPIX_RETURN_1Y = topix1y;

  console.log('=== 更新内容 ===');
  console.log(`  INDEX_MONTHS(${out.INDEX_MONTHS.length}): ${out.INDEX_MONTHS.join(',')}`);
  console.log(`  SAAS_INDEX_MONTHLY: ${out.SAAS_INDEX_MONTHLY.join(',')}`);
  console.log(`  TOPIX_MONTHLY: ${out.TOPIX_MONTHLY.join(',')}`);
  console.log(`  TOPIX_RETURN_1Y: ${out.TOPIX_RETURN_1Y}%`);

  if (!APPLY) {
    console.log('\n(dry-run のため書き込みなし。反映するには --apply を付けて実行)');
    return;
  }
  await firestorePatchMasked(token, 'saas-market', out);
  console.log('\n=== premiumContent/saas-market を更新しました（updateMask付きPATCH） ===');
}

main().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
