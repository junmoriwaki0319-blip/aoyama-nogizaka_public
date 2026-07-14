#!/usr/bin/env node
/**
 * SaaS / 外食産業セクターの財務指標を Yahoo Finance で更新するスクリプト
 *
 * - Firestore から既存データを読み込み、Yahoo Finance 派生フィールドのみ上書き
 * - ハンドキュレートフィールド（ARR・NRR・SSS・優待・brands 等）は保持
 * - 結果を data/saas-companies.json / data/food-companies.json に出力
 * - upload-premium-data.js を data/*.json から読むように修正済み前提
 *
 * 使い方: node scripts/refresh-saas-food.js
 */
const fs = require('fs');
const path = require('path');
const https = require('https');
const YahooFinance = require('yahoo-finance2').default;
const yf = new YahooFinance({ suppressNotices: ['yahooSurvey'] });

const PROJECT_ID = 'aoyama-nogizaka-activist';
const BASE = path.resolve(__dirname, '..');
const DATA_DIR = path.join(BASE, 'data');

const BATCH_SIZE = 5;
const BATCH_DELAY_MS = 300;
const sleep = ms => new Promise(r => setTimeout(r, ms));

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
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => { const json = JSON.parse(data); json.access_token ? resolve(json.access_token) : reject(new Error(data)); });
    });
    req.on('error', reject); req.write(postData); req.end();
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

function firestoreRead(token, collection, docId) {
  return new Promise((resolve, reject) => {
    const docPath = `projects/${PROJECT_ID}/databases/(default)/documents/${collection}/${docId}`;
    const req = https.request({
      hostname: 'firestore.googleapis.com', path: `/v1/${docPath}`, method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    }, (res) => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          const json = JSON.parse(data);
          const fields = {};
          for (const [k, v] of Object.entries(json.fields || {})) fields[k] = fromFirestoreValue(v);
          resolve(fields);
        } else reject(new Error(`Read failed (${res.statusCode}): ${data}`));
      });
    });
    req.on('error', reject); req.end();
  });
}

async function fetchYahoo(code, mcapUnit /* '億円' | '百万円' */) {
  const symbol = code + '.T';
  const mcapDivisor = mcapUnit === '百万円' ? 1e6 : 1e8;
  try {
    const [quote, summary] = await Promise.all([
      yf.quote(symbol),
      yf.quoteSummary(symbol, { modules: ['financialData', 'defaultKeyStatistics', 'summaryDetail'] }).catch(() => null),
    ]);
    const out = {
      stockPrice: quote.regularMarketPrice ?? null,
      marketCap: quote.marketCap ? Math.round(quote.marketCap / mcapDivisor) : null,
      per: null, pbr: null, roe: null, opMargin: null, netMargin: null,
      revenue: null, opProfit: null, netProfit: null,
      stockReturn1Y: null, dividendYield: null, psr: null,
    };
    if (summary) {
      const fd = summary.financialData || {};
      const ks = summary.defaultKeyStatistics || {};
      const sd = summary.summaryDetail || {};
      if (fd.operatingMargins != null) out.opMargin = +(fd.operatingMargins * 100).toFixed(1);
      if (fd.profitMargins != null) out.netMargin = +(fd.profitMargins * 100).toFixed(1);
      if (fd.returnOnEquity != null) out.roe = +(fd.returnOnEquity * 100).toFixed(1);
      const per = ks.trailingPE ?? ks.forwardPE ?? sd.trailingPE;
      const pbr = ks.priceToBook;
      const psr = ks.priceToSalesTrailing12Months;
      if (per != null) out.per = +per.toFixed(1);
      if (pbr != null) out.pbr = +pbr.toFixed(2);
      if (psr != null) out.psr = +psr.toFixed(1);
      if (fd.totalRevenue != null) out.revenue = Math.round(fd.totalRevenue / 1e6); // 百万円単位
      if (fd.ebitda != null) out.opProfit = Math.round(fd.ebitda / 1e6);
      if (sd.dividendYield != null) out.dividendYield = +(sd.dividendYield * 100).toFixed(2);
      if (sd.fiveYearAvgDividendYield != null && out.dividendYield == null) out.dividendYield = +sd.fiveYearAvgDividendYield.toFixed(2);
    }
    if (out.per == null && quote.trailingPE) out.per = +quote.trailingPE.toFixed(1);
    if (out.pbr == null && quote.priceToBook) out.pbr = +quote.priceToBook.toFixed(2);
    // 過去1年リターン
    try {
      const hist = await yf.chart(symbol, { period1: new Date(Date.now() - 365 * 24 * 3600 * 1000), interval: '1d' });
      if (hist && hist.quotes && hist.quotes.length > 0) {
        const first = hist.quotes.find(q => q.close != null);
        const last = [...hist.quotes].reverse().find(q => q.close != null);
        if (first && last && first.close > 0) {
          out.stockReturn1Y = +(((last.close - first.close) / first.close) * 100).toFixed(1);
        }
      }
    } catch (e) {}
    return out;
  } catch (e) {
    return null;
  }
}

async function refreshBatch(companies, mcapUnit) {
  const out = [];
  for (let i = 0; i < companies.length; i += BATCH_SIZE) {
    const batch = companies.slice(i, i + BATCH_SIZE);
    const results = await Promise.all(batch.map(async (c) => {
      const fresh = await fetchYahoo(c.code, mcapUnit);
      if (!fresh) { process.stdout.write('x'); return c; }
      process.stdout.write('.');
      // Yahoo派生フィールドのみ上書き、それ以外は保持
      const merged = { ...c };
      for (const [k, v] of Object.entries(fresh)) {
        if (v != null) merged[k] = v;
      }
      return merged;
    }));
    out.push(...results);
    if (i + BATCH_SIZE < companies.length) await sleep(BATCH_DELAY_MS);
  }
  process.stdout.write('\n');
  return out;
}

(async () => {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

  console.log('=== Firebase アクセストークンを取得中... ===');
  const token = await getAccessToken();
  console.log('  トークン取得成功\n');

  // SaaS
  console.log('=== SaaSセクター: Firestore から既存データ読込 ===');
  const saasDoc = await firestoreRead(token, 'premiumContent', 'saas-companies');
  const saasCompanies = saasDoc.companies || [];
  console.log(`  ${saasCompanies.length}社\n`);

  console.log(`=== SaaSセクター: Yahoo Finance データ取得 (${BATCH_SIZE}社並列) ===`);
  const saasRefreshed = await refreshBatch(saasCompanies, '億円');
  const saasOut = path.join(DATA_DIR, 'saas-companies.json');
  fs.writeFileSync(saasOut, JSON.stringify(saasRefreshed, null, 2), 'utf8');
  console.log(`  → ${saasRefreshed.length}社を ${saasOut} に保存\n`);

  // Food
  console.log('=== 外食産業セクター: Firestore から既存データ読込 ===');
  const foodDoc = await firestoreRead(token, 'premiumContent', 'food-companies');
  const foodCompanies = foodDoc.companies || [];
  console.log(`  ${foodCompanies.length}社\n`);

  console.log(`=== 外食産業セクター: Yahoo Finance データ取得 (${BATCH_SIZE}社並列) ===`);
  const foodRefreshed = await refreshBatch(foodCompanies, '百万円');
  const foodOut = path.join(DATA_DIR, 'food-companies.json');
  fs.writeFileSync(foodOut, JSON.stringify(foodRefreshed, null, 2), 'utf8');
  console.log(`  → ${foodRefreshed.length}社を ${foodOut} に保存\n`);

  console.log('=== 全セクター取得完了 ===');
  console.log(`  SaaS:    ${saasRefreshed.length}社`);
  console.log(`  外食:    ${foodRefreshed.length}社`);
})().catch(e => { console.error('ERROR:', e); process.exit(1); });
