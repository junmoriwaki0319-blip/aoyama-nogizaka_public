const https = require('https');

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const code = req.query.code;
  if (!code || !/^[0-9A-Za-z]{4}$/.test(code)) {
    return res.status(400).json({ success: false, error: '4桁の証券コードを指定してください' });
  }

  try {
    const ticker = code + '.T';
    const url = `https://finance.yahoo.co.jp/quote/${ticker}`;
    const html = await fetchHtml(url);

    // JSON埋め込みデータから抽出
    const extract = (key) => {
      const re = new RegExp(`"${key}":"([^"]*)"`, 'g');
      const matches = [...html.matchAll(re)];
      for (const m of matches) {
        if (m[1] && m[1] !== '' && m[1] !== '---') return m[1];
      }
      return null;
    };

    const pbr = extract('pbr');
    const roe = extract('roe');
    const equityRatio = extract('equityRatio');
    const eps = extract('eps');
    const bps = extract('bps');
    const dps = extract('dps');
    const sharesIssued = extract('sharesIssued');
    const shareDividendYield = extract('shareDividendYield');

    // 株価を抽出
    const priceMatches = html.match(/"price":"([\d,.]+)"/g) || [];
    let price = null;
    for (const pm of priceMatches) {
      const m = pm.match(/"price":"([\d,.]+)"/);
      if (m && m[1]) {
        const p = parseFloat(m[1].replace(/,/g, ''));
        if (!isNaN(p) && p > 0) { price = p; break; }
      }
    }

    // 企業名
    const nameMatch = html.match(/<title>([^<]*)</);
    let companyName = '';
    if (nameMatch) {
      companyName = nameMatch[1].replace(/【\d+】.*$/, '').replace(/\s*\|.*$/, '').trim();
    }

    // 発行済株式数
    const sharesIssuedNum = sharesIssued ? parseFloat(sharesIssued.replace(/,/g, '')) : null;

    // 時価総額（発行済全数での概算、自己株式控除はフロントエンド側）
    let marketCapOku = null;
    if (price && sharesIssuedNum) {
      marketCapOku = Math.round(price * sharesIssuedNum / 100000000);
    }

    // 配当性向 = DPS / EPS × 100
    let payoutRatio = null;
    const dpsNum = dps ? parseFloat(dps.replace(/,/g, '')) : null;
    const epsNum = eps ? parseFloat(eps.replace(/,/g, '')) : null;
    if (dpsNum && epsNum && epsNum > 0) {
      payoutRatio = Math.round(dpsNum / epsNum * 10000) / 100;
    }

    res.json({
      success: true,
      data: {
        companyName,
        pbr: pbr ? parseFloat(pbr) : null,
        roe: roe ? parseFloat(roe) : null,
        payoutRatio,
        equityRatio: equityRatio ? parseFloat(equityRatio) : null,
        marketCapOku,
        sharesIssued: sharesIssuedNum,
        price,
        dividendYield: shareDividendYield ? parseFloat(shareDividendYield) : null,
        dps: dpsNum,
        eps: epsNum,
        bps: bps ? parseFloat(bps.replace(/,/g, '')) : null,
      }
    });
  } catch (err) {
    console.error('Error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
};

function fetchHtml(url) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept-Language': 'ja,en;q=0.9'
      }
    };
    https.get(options, (resp) => {
      if (resp.statusCode >= 300 && resp.statusCode < 400 && resp.headers.location) {
        fetchHtml(resp.headers.location).then(resolve).catch(reject);
        resp.resume();
        return;
      }
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => resolve(data));
    }).on('error', reject);
  });
}
