const https = require('https');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const code = req.query.code;
  if (!code || !/^[0-9A-Za-z]{4}$/.test(code)) {
    return res.status(400).json({ success: false, error: '4桁の証券コードを指定してください' });
  }

  try {
    const [kb, yf] = await Promise.allSettled([
      fetchKabutan(code),
      fetchYFChart(code),
    ]);
    const k = kb.status === 'fulfilled' ? kb.value : {};
    const y = yf.status === 'fulfilled' ? yf.value : {};

    const _ = (...vs) => { for (const v of vs) if (v != null) return v; return null; };
    const result = {
      companyName: k.companyName || y.companyName || '',
      price: _(k.price, y.price),
      pbr: _(k.pbr),
      per: _(k.per),
      roe: _(k.roe),
      eps: _(k.eps),
      bps: _(k.bps),
      dps: _(k.dps),
      dividendYield: _(k.dividendYield),
      payoutRatio: _(k.payoutRatio),
      equityRatio: _(k.equityRatio),
      marketCapOku: _(k.marketCapOku),
      sharesIssued: _(k.sharesIssued),
      market: k.market || null,
      sector: k.sector || null,
      fiftyTwoWeekHigh: _(y.fiftyTwoWeekHigh),
      fiftyTwoWeekLow: _(y.fiftyTwoWeekLow),
      previousClose: _(y.previousClose),
    };
    if (!result.marketCapOku && result.price && result.sharesIssued) {
      result.marketCapOku = Math.round(result.price * result.sharesIssued / 100000000);
    }
    res.json({ success: true, data: result });
  } catch (err) {
    console.error('Stock fetch error:', err.message);
    res.status(500).json({ success: false, error: 'データ取得に失敗しました' });
  }
};

// ═══════════════════════════════════════════════════════════════
// kabutan.jp parser
// ═══════════════════════════════════════════════════════════════
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
    if (mcapMatch) {
      result.marketCapOku = parseInt(mcapMatch[1].replace(/,/g, '')) * 10000 + parseInt(mcapMatch[2].replace(/,/g, ''));
    } else {
      const mcapMatch2 = rawText.match(/時価総額\s*([0-9,]+)\s*億/);
      if (mcapMatch2) result.marketCapOku = parseInt(mcapMatch2[1].replace(/,/g, ''));
    }
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

  // EPS, DPS from earnings table
  const kesIdx = mainHtml.indexOf('決算期');
  if (kesIdx > -1) {
    const tableStart = mainHtml.lastIndexOf('<table', kesIdx);
    const tableEnd = mainHtml.indexOf('</table>', kesIdx);
    if (tableStart > -1 && tableEnd > -1) {
      const tableBlock = mainHtml.substring(tableStart, tableEnd + 10);
      const rows = tableBlock.match(/<tr[\s\S]*?<\/tr>/g);
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
              if (!isNaN(dpsVal)) {
                result.dps = dpsVal;
                if (epsVal > 0) result.payoutRatio = Math.round(dpsVal / epsVal * 10000) / 100;
              }
              break;
            }
          }
        }
      }
    }
  }

  if (result.price && result.pbr && result.pbr > 0) {
    result.bps = Math.round(result.price / result.pbr * 100) / 100;
  }
  if (result.eps && result.bps && result.bps > 0) {
    result.roe = Math.round(result.eps / result.bps * 10000) / 100;
  }

  // Equity ratio from financial table
  const bpsHeaderIdx = finHtml.indexOf('１株純資産');
  if (bpsHeaderIdx > -1) {
    const tblStart = finHtml.lastIndexOf('<table', bpsHeaderIdx);
    const tblEnd = finHtml.indexOf('</table>', bpsHeaderIdx);
    if (tblStart > -1 && tblEnd > -1) {
      const tbl = finHtml.substring(tblStart, tblEnd + 10);
      const fRows = tbl.match(/<tr[\s\S]*?<\/tr>/g);
      if (fRows) {
        for (let ri = fRows.length - 1; ri >= 1; ri--) {
          const cells = fRows[ri].match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g);
          if (cells && cells.length >= 7) {
            const vals = cells.map(c => c.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/,/g, '').replace(/\s+/g, '').trim());
            if (vals[0].includes('前期比')) continue;
            const eqVal = parseFloat(vals[2]);
            if (!isNaN(eqVal) && eqVal > 0 && eqVal <= 100) {
              result.equityRatio = eqVal;
              break;
            }
          }
        }
      }
    }
  }

  return result;
}

// ═══════════════════════════════════════════════════════════════
// Yahoo Finance Chart API
// ═══════════════════════════════════════════════════════════════
async function fetchYFChart(code) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${code}.T?interval=1d&range=1d`;
  const data = await fetchJson(url);
  const meta = data?.chart?.result?.[0]?.meta;
  if (!meta) return {};
  return {
    companyName: meta.longName || meta.shortName || '',
    price: meta.regularMarketPrice || null,
    previousClose: meta.chartPreviousClose || null,
    fiftyTwoWeekHigh: meta.fiftyTwoWeekHigh || null,
    fiftyTwoWeekLow: meta.fiftyTwoWeekLow || null,
  };
}

function fetchHtml(url, depth = 0) {
  return new Promise((resolve, reject) => {
    if (depth > 5) return reject(new Error('Too many redirects'));
    const urlObj = new URL(url);
    https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ja,en;q=0.9'
      }
    }, (resp) => {
      if (resp.statusCode >= 300 && resp.statusCode < 400 && resp.headers.location) {
        fetchHtml(resp.headers.location, depth + 1).then(resolve).catch(reject);
        resp.resume();
        return;
      }
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => resolve(data));
    }).on('error', reject);
  });
}

function fetchJson(url, depth = 0) {
  return new Promise((resolve, reject) => {
    if (depth > 5) return reject(new Error('Too many redirects'));
    const urlObj = new URL(url);
    https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    }, (resp) => {
      if (resp.statusCode >= 300 && resp.statusCode < 400 && resp.headers.location) {
        fetchJson(resp.headers.location, depth + 1).then(resolve).catch(reject);
        resp.resume();
        return;
      }
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { resolve(null); }
      });
    }).on('error', reject);
  });
}
