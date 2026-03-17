const https = require('https');

/**
 * 政策保有株式の現在株価を一括取得するAPI
 * POST /api/stock-prices
 * Body: { holdings: [{ name: "ＫＤＤＩ㈱", shares: 203294600 }, ...] }
 * Returns: { results: [{ name, ticker, lastClose, avg3m, avg6m, avg12m }, ...] }
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
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

  // Step 1: Yahoo Finance検索で全銘柄のティッカーを並列取得
  const searchResults = await Promise.allSettled(
    holdings.slice(0, 25).map(h => searchTicker(cleanName(h.name)))
  );

  const tickerMap = []; // { name, shares, ticker }
  for (let i = 0; i < searchResults.length; i++) {
    const ticker = searchResults[i].status === 'fulfilled' ? searchResults[i].value : null;
    tickerMap.push({
      name: holdings[i].name,
      shares: holdings[i].shares,
      ticker,
    });
  }

  // Step 2: ティッカーが見つかった銘柄の株価チャートを並列取得
  const tickersFound = tickerMap.filter(t => t.ticker);
  const chartResults = await Promise.allSettled(
    tickersFound.map(t => fetchChart(t.ticker))
  );

  // 結果を組み立て
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

/**
 * 会社名をYahoo Finance検索用にクリーニング
 */
function cleanName(name) {
  return name
    .replace(/[㈱㈲]/g, '')
    .replace(/[株式会社|有限会社]/g, '')
    .replace(/[\s\u3000]+/g, ' ')
    .replace(/[（）()「」]/g, '')
    .trim();
}

/**
 * Yahoo Finance検索APIでティッカーシンボルを取得
 */
function searchTicker(query) {
  const url = `https://query2.finance.yahoo.com/v1/finance/search?q=${encodeURIComponent(query)}&quotesCount=5&newsCount=0&lang=ja-JP&region=JP`;
  return fetchJSON(url).then(data => {
    if (!data.quotes || data.quotes.length === 0) return null;
    // 東証（.T）を優先
    const tse = data.quotes.find(q => q.symbol && q.symbol.endsWith('.T') && q.quoteType === 'EQUITY');
    if (tse) return tse.symbol;
    // その他の取引所の株式
    const equity = data.quotes.find(q => q.quoteType === 'EQUITY');
    if (equity) return equity.symbol;
    return data.quotes[0]?.symbol || null;
  });
}

/**
 * Yahoo FinanceチャートAPIから株価データを取得（1年分）
 */
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
    const TD_PER_MONTH = 21; // 営業日/月の近似値
    const avg3m = avg(valid.slice(-TD_PER_MONTH * 3));
    const avg6m = avg(valid.slice(-TD_PER_MONTH * 6));
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

function avg(arr) {
  if (!arr.length) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function round2(v) {
  return v != null ? Math.round(v * 100) / 100 : null;
}

function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 8000,
    }, resp => {
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { resolve({}); }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}
