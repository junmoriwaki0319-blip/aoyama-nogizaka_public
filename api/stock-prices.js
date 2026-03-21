const https = require('https');

const ALLOWED_HOSTS = ['finance.yahoo.co.jp', 'query1.finance.yahoo.com', 'query2.finance.yahoo.com'];
const MAX_RESPONSE_SIZE = 2 * 1024 * 1024;

function isAllowedHost(url) {
  try { const h = new URL(url).hostname; return ALLOWED_HOSTS.some(a => h === a || h.endsWith('.' + a)); }
  catch { return false; }
}

/**
 * 政策保有株式の現在株価を一括取得するAPI
 * POST /api/stock-prices
 * Body: { holdings: [{ name: "ＫＤＤＩ㈱", shares: 203294600 }, ...] }
 * Returns: { results: [{ name, ticker, lastClose, avg3m, avg6m, avg12m, currency }, ...] }
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', 'https://aoyama-nogizaka.com');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  let holdings;
  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    holdings = body.holdings;
  } catch {
    return res.status(400).json({ success: false, error: 'Invalid request body' });
  }

  if (!Array.isArray(holdings) || holdings.length === 0) {
    return res.status(400).json({ success: false, error: 'holdings array required' });
  }

  const items = holdings.slice(0, 25);

  // Step 1: 全銘柄のティッカーを並列検索
  const searchResults = await Promise.allSettled(
    items.map(h => searchTicker(h.name))
  );

  const tickerMap = [];
  for (let i = 0; i < searchResults.length; i++) {
    const ticker = searchResults[i].status === 'fulfilled' ? searchResults[i].value : null;
    tickerMap.push({ name: items[i].name, shares: items[i].shares, ticker });
  }

  // Step 2: ティッカーが見つかった銘柄の株価チャートを並列取得
  const tickersFound = tickerMap.filter(t => t.ticker);
  const chartResults = await Promise.allSettled(
    tickersFound.map(t => fetchChart(t.ticker))
  );

  const priceMap = {};
  for (let i = 0; i < tickersFound.length; i++) {
    if (chartResults[i].status === 'fulfilled' && chartResults[i].value) {
      priceMap[tickersFound[i].ticker] = chartResults[i].value;
    }
  }

  const results = tickerMap.map(t => {
    const prices = t.ticker ? (priceMap[t.ticker] || null) : null;
    return {
      name: t.name,
      shares: t.shares,
      ticker: t.ticker,
      lastClose: prices?.lastClose || null,
      avg3m: prices?.avg3m || null,
      avg6m: prices?.avg6m || null,
      avg12m: prices?.avg12m || null,
      currency: prices?.currency || 'JPY',
    };
  });

  res.json({ success: true, results });
};

// ═══════════════════════════════════════════════════════════════
// 会社名クリーニング
// ═══════════════════════════════════════════════════════════════

function cleanName(name) {
  return name
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/＆/g, '&').replace(/・/g, ' ')
    .replace(/[㈱㈲]/g, '')
    .replace(/株式会社/g, '').replace(/有限会社/g, '')
    .replace(/[（）()「」\[\]]/g, '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]{2,})([A-Z][a-z])/g, '$1 $2')
    .replace(/(LTD|CO|INC|PLC|Tbk)\./gi, '$1')
    .replace(/[\s\u3000]+/g, ' ')
    .trim();
}

/** 日本語文字を含むか判定 */
function hasJapanese(str) {
  return /[ぁ-んァ-ヶー一-龠]/.test(str);
}

/** 検索しにくい社名の別名マッピング */
const NAME_ALIASES = {
  '日本電信電話': 'NTT',
  'ＮＴＴ': 'NTT',
  '東日本電信電話': 'NTT東日本',
  '西日本電信電話': 'NTT西日本',
  '三菱UFJフィナンシャル グループ': '三菱UFJ',
  '三井住友フィナンシャルグループ': '三井住友FG',
  'みずほフィナンシャルグループ': 'みずほFG',
  'SOMPOホールディングス': 'SOMPO',
  '東京海上ホールディングス': '東京海上',
};

// ═══════════════════════════════════════════════════════════════
// ティッカー検索（日本株 → Yahoo Finance Japan、海外株 → グローバルAPI）
// ═══════════════════════════════════════════════════════════════

async function searchTicker(rawName) {
  const name = cleanName(rawName);

  // 別名があればそちらでも検索
  const alias = NAME_ALIASES[name];

  // 日本語名 → Yahoo Finance Japan検索ページをスクレイプ
  if (hasJapanese(name) || alias) {
    const ticker = await searchYahooJP(alias || name);
    if (ticker) return ticker;
    // 別名で見つからなければ元の名前でも試す
    if (alias) {
      const ticker2 = await searchYahooJP(name);
      if (ticker2) return ticker2;
    }
  }

  // 英語名 → まずYahoo Finance Japanで試す
  const jpTicker = await searchYahooJP(name);
  if (jpTicker) return jpTicker;

  // グローバルYahoo Finance検索（海外株）
  return searchYahooGlobal(name);
}

/**
 * Yahoo Finance Japan検索ページから証券コードを取得
 */
async function searchYahooJP(query) {
  try {
    const html = await fetchText(
      `https://finance.yahoo.co.jp/search/?query=${encodeURIComponent(query)}`
    );
    // JSONデータ内の証券コードを抽出
    const match = html.match(/"code":"(\d{4}[A-Z0-9]?)"/);
    return match ? match[1] + '.T' : null;
  } catch {
    return null;
  }
}

/**
 * グローバルYahoo Finance検索API（海外株用）
 */
async function searchYahooGlobal(query) {
  try {
    const url = `https://query2.finance.yahoo.com/v1/finance/search?q=${encodeURIComponent(query)}&quotesCount=5&newsCount=0`;
    const data = await fetchJSON(url);
    if (!data.quotes || data.quotes.length === 0) return null;
    const equity = data.quotes.find(q => q.quoteType === 'EQUITY');
    return equity?.symbol || null;
  } catch {
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════
// 株価チャート取得
// ═══════════════════════════════════════════════════════════════

function fetchChart(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?range=1y&interval=1d`;
  return fetchJSON(url).then(data => {
    const result = data.chart?.result?.[0];
    if (!result) return null;

    const meta = result.meta || {};
    const closes = result.indicators?.quote?.[0]?.close || [];
    const valid = closes.filter(c => c != null && c > 0);
    if (valid.length === 0) return null;

    const lastClose = valid[valid.length - 1];
    const TD = 21; // 営業日/月の近似値
    const avg3m = avg(valid.slice(-TD * 3));
    const avg6m = avg(valid.slice(-TD * 6));
    const avg12m = avg(valid);

    return {
      lastClose: round2(lastClose),
      avg3m: round2(avg3m),
      avg6m: round2(avg6m),
      avg12m: round2(avg12m),
      currency: meta.currency || 'JPY',
    };
  });
}

// ═══════════════════════════════════════════════════════════════
// ユーティリティ
// ═══════════════════════════════════════════════════════════════

function avg(arr) {
  if (!arr.length) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function round2(v) {
  return v != null ? Math.round(v * 100) / 100 : null;
}

function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    if (!isAllowedHost(url)) return reject(new Error('Untrusted host'));
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      timeout: 10000,
    }, resp => {
      let data = ''; let size = 0;
      resp.on('data', chunk => { size += chunk.length; if (size > MAX_RESPONSE_SIZE) { resp.destroy(); return reject(new Error('Response too large')); } data += chunk; });
      resp.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { resolve({}); }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}

function fetchText(url) {
  return new Promise((resolve, reject) => {
    if (!isAllowedHost(url)) return reject(new Error('Untrusted host'));
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      timeout: 10000,
    }, resp => {
      let data = ''; let size = 0;
      resp.on('data', chunk => { size += chunk.length; if (size > MAX_RESPONSE_SIZE) { resp.destroy(); return reject(new Error('Response too large')); } data += chunk; });
      resp.on('end', () => resolve(data));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}
