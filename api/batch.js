const https = require('https');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST') { return res.status(405).json({ error: 'POST only' }); }

  const codes = req.body?.codes;
  if (!Array.isArray(codes) || codes.length === 0) {
    return res.json({ success: false, error: '銘柄コードの配列が必要です' });
  }
  if (codes.length > 50) {
    return res.json({ success: false, error: '一度に50銘柄までです' });
  }

  const results = [];
  const concurrency = 3;
  for (let i = 0; i < codes.length; i += concurrency) {
    const batch = codes.slice(i, i + concurrency);
    const promises = batch.map(code => fetchStock(code));
    const batchResults = await Promise.allSettled(promises);
    results.push(...batchResults.map((r, idx) =>
      r.status === 'fulfilled' ? r.value : { code: batch[idx], error: r.reason?.message }
    ));
    if (i + concurrency < codes.length) await new Promise(r => setTimeout(r, 300));
  }
  res.json({ success: true, data: results });
};

async function fetchStock(code) {
  const [kb, yf] = await Promise.allSettled([fetchKabutan(code), fetchYFChart(code)]);
  const k = kb.status === 'fulfilled' ? kb.value : {};
  const y = yf.status === 'fulfilled' ? yf.value : {};
  return {
    code,
    companyName: k.companyName || y.companyName || '',
    price: k.price || y.price || null,
    pbr: k.pbr || null, per: k.per || null, roe: k.roe || null,
    bps: k.bps || null, eps: k.eps || null, dps: k.dps || null,
    dividendYield: k.dividendYield || null, payoutRatio: k.payoutRatio || null,
    equityRatio: k.equityRatio || null,
    marketCapOku: k.marketCapOku || null, sharesIssued: k.sharesIssued || null,
    market: k.market || null, sector: k.sector || null,
  };
}

async function fetchKabutan(code) {
  const [mainHtml, finHtml] = await Promise.all([
    fetchHtml(`https://kabutan.jp/stock/?code=${code}`),
    fetchHtml(`https://kabutan.jp/stock/finance?code=${code}`),
  ]);
  const result = {};

  const titleMatch = mainHtml.match(/<title>([^（【]+)/);
  if (titleMatch) result.companyName = titleMatch[1].trim();

  const priceMatch = mainHtml.match(/stock_price=([0-9.]+)/);
  if (priceMatch) result.price = parseFloat(priceMatch[1]);

  const indicatorBlock = mainHtml.match(/<abbr title="Price Earnings Ratio">PER<\/abbr>[\s\S]*?<\/tbody>/);
  if (indicatorBlock) {
    const rawText = indicatorBlock[0].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
    const nums = rawText.match(/([0-9,.]+)\s*倍/g);
    if (nums && nums.length >= 2) {
      result.per = parseFloat(nums[0].replace(/[倍,]/g, ''));
      result.pbr = parseFloat(nums[1].replace(/[倍,]/g, ''));
    }
    const yieldMatch = rawText.match(/([0-9.]+)\s*％/);
    if (yieldMatch) result.dividendYield = parseFloat(yieldMatch[1]);
    const mcapMatch = rawText.match(/時価総額\s*([0-9,]+)\s*兆\s*([0-9,]+)\s*億/);
    if (mcapMatch) result.marketCapOku = parseInt(mcapMatch[1].replace(/,/g, '')) * 10000 + parseInt(mcapMatch[2].replace(/,/g, ''));
    else { const m2 = rawText.match(/時価総額\s*([0-9,]+)\s*億/); if (m2) result.marketCapOku = parseInt(m2[1].replace(/,/g, '')); }
  }

  const sharesMatch = mainHtml.match(/発行済株式数[\s\S]*?([\d,]+)\s*(?:&nbsp;)*\s*株/);
  if (sharesMatch) result.sharesIssued = parseFloat(sharesMatch[1].replace(/,/g, ''));

  const sectorBlock = mainHtml.match(/>\s*業種\s*<[\s\S]*?>\s*([^\s<][^<]{1,20})\s*</);
  if (sectorBlock) result.sector = sectorBlock[1].trim();

  const marketSpan = mainHtml.match(/class="market"[^>]*>([^<]+)/);
  if (marketSpan) {
    const mt = marketSpan[1].trim();
    if (mt.includes('Ｐ') || mt.includes('P')) result.market = '東証プライム';
    else if (mt.includes('Ｓ') || mt.includes('S')) result.market = '東証スタンダード';
    else if (mt.includes('Ｇ') || mt.includes('G')) result.market = '東証グロース';
    else result.market = mt;
  }

  const kesIdx = mainHtml.indexOf('決算期');
  if (kesIdx > -1) {
    const ts = mainHtml.lastIndexOf('<table', kesIdx);
    const te = mainHtml.indexOf('</table>', kesIdx);
    if (ts > -1 && te > -1) {
      const rows = mainHtml.substring(ts, te + 10).match(/<tr[\s\S]*?<\/tr>/g);
      if (rows) {
        for (let ri = rows.length - 1; ri >= 1; ri--) {
          const cells = rows[ri].match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g);
          if (cells && cells.length >= 7) {
            const vals = cells.map(c => c.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/,/g, '').trim());
            if (vals[0].includes('前期比') || vals[0].includes('決算期')) continue;
            const epsVal = parseFloat(vals[4]);
            const dpsVal = parseFloat(vals[5]);
            if (!isNaN(epsVal) && epsVal !== 0) {
              result.eps = epsVal;
              if (!isNaN(dpsVal)) { result.dps = dpsVal; if (epsVal > 0) result.payoutRatio = Math.round(dpsVal / epsVal * 10000) / 100; }
              break;
            }
          }
        }
      }
    }
  }

  if (result.price && result.pbr && result.pbr > 0) result.bps = Math.round(result.price / result.pbr * 100) / 100;
  if (result.eps && result.bps && result.bps > 0) result.roe = Math.round(result.eps / result.bps * 10000) / 100;

  const bpsIdx = finHtml.indexOf('１株純資産');
  if (bpsIdx > -1) {
    const ts = finHtml.lastIndexOf('<table', bpsIdx);
    const te = finHtml.indexOf('</table>', bpsIdx);
    if (ts > -1 && te > -1) {
      const fRows = finHtml.substring(ts, te + 10).match(/<tr[\s\S]*?<\/tr>/g);
      if (fRows) {
        for (let ri = fRows.length - 1; ri >= 1; ri--) {
          const cells = fRows[ri].match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g);
          if (cells && cells.length >= 7) {
            const vals = cells.map(c => c.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/,/g, '').replace(/\s+/g, '').trim());
            if (vals[0].includes('前期比')) continue;
            const eqVal = parseFloat(vals[2]);
            if (!isNaN(eqVal) && eqVal > 0 && eqVal <= 100) { result.equityRatio = eqVal; break; }
          }
        }
      }
    }
  }
  return result;
}

async function fetchYFChart(code) {
  const data = await fetchJson(`https://query1.finance.yahoo.com/v8/finance/chart/${code}.T?interval=1d&range=1d`);
  const meta = data?.chart?.result?.[0]?.meta;
  if (!meta) return {};
  return { companyName: meta.longName || meta.shortName || '', price: meta.regularMarketPrice || null, previousClose: meta.chartPreviousClose || null, fiftyTwoWeekHigh: meta.fiftyTwoWeekHigh || null, fiftyTwoWeekLow: meta.fiftyTwoWeekLow || null };
}

function fetchHtml(url) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36', 'Accept': 'text/html', 'Accept-Language': 'ja' } }, (r) => {
      if (r.statusCode >= 300 && r.statusCode < 400 && r.headers.location) { fetchHtml(r.headers.location).then(resolve).catch(reject); r.resume(); return; }
      let d = ''; r.on('data', c => d += c); r.on('end', () => resolve(d));
    }).on('error', reject);
  });
}

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } }, (r) => {
      if (r.statusCode >= 300 && r.statusCode < 400 && r.headers.location) { fetchJson(r.headers.location).then(resolve).catch(reject); r.resume(); return; }
      let d = ''; r.on('data', c => d += c); r.on('end', () => { try { resolve(JSON.parse(d)); } catch { resolve(null); } });
    }).on('error', reject);
  });
}
