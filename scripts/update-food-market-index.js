#!/usr/bin/env node
/**
 * 外食産業セクター指数 & TOPIX月次データを取得し Firestore にアップロード
 *
 * Yahoo Finance API から:
 *   - TOPIX (^TPX) の月次終値
 *   - 外食主要企業の月次終値 → 時価総額加重平均で独自指数を算出
 * 起点月を100として正規化し、Firestore premiumContent/food-market に保存
 *
 * 使い方: node scripts/update-food-market-index.js
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const PROJECT_ID = 'aoyama-nogizaka-activist';

// 外食主要企業（時価総額上位 + カバレッジ確保用）
// code は東証コード、weight は時価総額ウェイトの近似（百億円単位、定期更新推奨）
const FOOD_CONSTITUENTS = [
  { code: '7550', name: 'ゼンショー', weight: 153 },
  { code: '3563', name: 'FOOD&LIFE', weight: 112 },
  { code: '2702', name: '日本マクドナルド', weight: 104 },
  { code: '3197', name: 'すかいらーく', weight: 76 },
  { code: '3397', name: 'トリドール', weight: 38 },
  { code: '7581', name: 'サイゼリヤ', weight: 35 },
  { code: '3387', name: 'クリエイト・レストランツ', weight: 31 },
  { code: '9861', name: '吉野家', weight: 21 },
  { code: '7616', name: 'コロワイド', weight: 20 },
  { code: '9936', name: '王将フードサービス', weight: 20 },
  { code: '3097', name: '物語コーポレーション', weight: 19 },
  { code: '2695', name: 'くら寿司', weight: 15 },
  { code: '8179', name: 'ロイヤル', weight: 15 },
  { code: '7630', name: '壱番屋', weight: 14 },
  { code: '3543', name: 'コメダ', weight: 14 },
];

// ---------- Yahoo Finance ----------

function fetchYFMonthly(ticker, months = 15) {
  // months+2 で余裕を持って取得（月末ずれ対策）
  const days = (months + 2) * 31;
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1mo&range=${days}d`;
  return fetchJSON(url).then(data => {
    const result = data?.chart?.result?.[0];
    if (!result) return null;
    const timestamps = result.timestamp || [];
    const closes = result.indicators?.quote?.[0]?.close || [];
    const points = [];
    for (let i = 0; i < timestamps.length; i++) {
      if (closes[i] == null) continue;
      const d = new Date(timestamps[i] * 1000);
      const label = `${String(d.getFullYear()).slice(2)}/${String(d.getMonth() + 1).padStart(2, '0')}`;
      points.push({ label, close: closes[i] });
    }
    return points;
  });
}

// ---------- 指数算出 ----------

async function buildIndices() {
  console.log('=== TOPIX月次データ取得中... ===');
  // 1306.T = NEXT FUNDS TOPIX連動型上場投信（TOPIX ETF）
  const topixRaw = await fetchYFMonthly('1306.T', 15);
  if (!topixRaw || topixRaw.length < 2) {
    throw new Error('TOPIX データ取得失敗');
  }
  console.log(`  TOPIX: ${topixRaw.length} データポイント取得`);

  console.log('\n=== 外食主要企業の月次データ取得中... ===');
  const stockResults = await Promise.allSettled(
    FOOD_CONSTITUENTS.map(c => fetchYFMonthly(`${c.code}.T`, 15))
  );

  // 各企業の月次リターンを算出
  const validStocks = [];
  for (let i = 0; i < FOOD_CONSTITUENTS.length; i++) {
    const c = FOOD_CONSTITUENTS[i];
    if (stockResults[i].status === 'fulfilled' && stockResults[i].value && stockResults[i].value.length >= 2) {
      validStocks.push({ ...c, data: stockResults[i].value });
      console.log(`  ✓ ${c.code} ${c.name}: ${stockResults[i].value.length}ポイント`);
    } else {
      console.log(`  ✗ ${c.code} ${c.name}: 取得失敗`);
    }
  }

  if (validStocks.length < 5) {
    throw new Error(`十分な企業データが取得できません (${validStocks.length}/15)`);
  }

  // 共通の月ラベルを決定（TOPIXの月ラベルを基準）
  const topixLabels = topixRaw.map(p => p.label);

  // TOPIX指数（起点=100）
  const topixBase = topixRaw[0].close;
  const topixIndex = topixRaw.map(p => Math.round(p.close / topixBase * 1000) / 10);

  // 外食セクター指数（時価総額加重平均）
  // 各月で、各企業の月次リターン（対起点月）を時価総額ウェイトで加重平均
  const restaurantIndex = topixLabels.map(label => {
    let weightedReturn = 0;
    let totalWeight = 0;
    for (const stock of validStocks) {
      const basePoint = stock.data[0];
      const point = stock.data.find(p => p.label === label);
      if (basePoint && point) {
        const ret = point.close / basePoint.close;
        weightedReturn += ret * stock.weight;
        totalWeight += stock.weight;
      }
    }
    if (totalWeight === 0) return null;
    return Math.round(weightedReturn / totalWeight * 1000) / 10;
  });

  // 1Yリターン（直近 vs 12ヶ月前）
  const topix1y = topixRaw.length >= 13
    ? Math.round((topixRaw[topixRaw.length - 1].close / topixRaw[topixRaw.length - 13].close - 1) * 1000) / 10
    : Math.round((topixRaw[topixRaw.length - 1].close / topixRaw[0].close - 1) * 1000) / 10;

  console.log(`\n=== 算出結果 ===`);
  console.log(`  月数: ${topixLabels.length}`);
  console.log(`  期間: ${topixLabels[0]} ～ ${topixLabels[topixLabels.length - 1]}`);
  console.log(`  TOPIX 1Yリターン: ${topix1y > 0 ? '+' : ''}${topix1y}%`);
  console.log(`  TOPIX指数: [${topixIndex.join(', ')}]`);
  console.log(`  外食指数: [${restaurantIndex.join(', ')}]`);

  return {
    TOPIX_RETURN_1Y: topix1y,
    TOPIX_MONTHLY: topixIndex,
    RESTAURANT_INDEX_MONTHLY: restaurantIndex,
    INDEX_MONTHS: topixLabels,
  };
}

// ---------- Firebase ----------

function getAccessToken() {
  const configPath = path.join(process.env.HOME || process.env.USERPROFILE, '.config', 'configstore', 'firebase-tools.json');
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const refreshToken = config.tokens.refresh_token;

  return new Promise((resolve, reject) => {
    const postData = `grant_type=refresh_token&refresh_token=${encodeURIComponent(refreshToken)}&client_id=563584335869-fgrhgmd47bqnekij5i8b5pr03ho849e6.apps.googleusercontent.com&client_secret=j9iVZfS8kkCEFUPaAeJV0sAi`;
    const req = https.request({
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        const json = JSON.parse(data);
        if (json.access_token) resolve(json.access_token);
        else reject(new Error('Token refresh failed: ' + data));
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

function toFirestoreValue(val) {
  if (val === null || val === undefined) return { nullValue: null };
  if (typeof val === 'boolean') return { booleanValue: val };
  if (typeof val === 'number') {
    if (Number.isInteger(val)) return { integerValue: String(val) };
    return { doubleValue: val };
  }
  if (typeof val === 'string') return { stringValue: val };
  if (Array.isArray(val)) {
    return { arrayValue: { values: val.map(toFirestoreValue) } };
  }
  if (typeof val === 'object') {
    const fields = {};
    for (const [k, v] of Object.entries(val)) {
      fields[k] = toFirestoreValue(v);
    }
    return { mapValue: { fields } };
  }
  return { stringValue: String(val) };
}

function firestoreWrite(accessToken, collection, docId, data) {
  return new Promise((resolve, reject) => {
    const fields = {};
    for (const [key, value] of Object.entries(data)) {
      fields[key] = toFirestoreValue(value);
    }
    const body = JSON.stringify({ fields });
    const docPath = `projects/${PROJECT_ID}/databases/(default)/documents/${collection}/${docId}`;
    const req = https.request({
      hostname: 'firestore.googleapis.com',
      path: `/v1/${docPath}`,
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(JSON.parse(data));
        else reject(new Error(`Firestore write failed (${res.statusCode}): ${data}`));
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ---------- Utility ----------

function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      timeout: 15000,
    }, resp => {
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { resolve(null); }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}

// ---------- Main ----------

async function main() {
  const indices = await buildIndices();

  console.log('\n=== Firestoreにアップロード中... ===');
  const token = await getAccessToken();

  await firestoreWrite(token, 'premiumContent', 'food-market', indices);
  console.log('  premiumContent/food-market を更新しました');
  console.log('\n=== 完了 ===');
}

main().catch(e => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
