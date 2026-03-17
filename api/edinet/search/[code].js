const https = require('https');

/**
 * EDINET書類検索API
 * EDINETコード(E+5桁)または証券コード(4桁数字)で有価証券報告書を検索する
 * GET /api/edinet/search/:code?apiKey=xxx
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const code = req.query.code;
  const apiKey = req.query.apiKey || process.env.EDINET_API_KEY;

  const isEdinetCode = /^E\d{5}$/.test(code);
  const isSecCode = /^[0-9A-Za-z]{4}$/.test(code) && !isEdinetCode;

  if (!code || (!isEdinetCode && !isSecCode)) {
    return res.status(400).json({ success: false, error: 'EDINETコード(E+5桁)または証券コード(4桁英数字)を指定してください' });
  }
  if (!apiKey) {
    return res.status(400).json({ success: false, error: 'EDINET APIキーが必要です' });
  }

  // 証券コードの場合、EDINET APIのsecCodeは5桁（末尾0付き）
  // 英字入りコード（166Aなど）の場合もそのまま末尾0を付与して検索
  const secCode5 = isSecCode ? code.toUpperCase() + '0' : null;

  try {
    // 過去400日分の日付を生成（直近から遡る）
    const dates = [];
    const today = new Date();
    for (let i = 0; i < 400; i++) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      const dow = d.getDay();
      if (dow === 0 || dow === 6) continue;
      dates.push(formatDate(d));
    }

    // バッチで並列検索（20件ずつ）
    const BATCH_SIZE = 20;
    let found = [];

    for (let i = 0; i < dates.length; i += BATCH_SIZE) {
      const batch = dates.slice(i, i + BATCH_SIZE);
      const promises = batch.map(date =>
        fetchDocList(date, apiKey).catch(() => [])
      );
      const results = await Promise.all(promises);

      for (const docs of results) {
        for (const doc of docs) {
          if (doc.docTypeCode !== '120') continue;
          const match = isEdinetCode
            ? doc.edinetCode === code
            : doc.secCode === secCode5;
          if (match) {
            found.push({
              docID: doc.docID,
              docDescription: doc.docDescription || '',
              periodStart: doc.periodStart,
              periodEnd: doc.periodEnd,
              submitDateTime: doc.submitDateTime,
              filerName: doc.filerName,
              edinetCode: doc.edinetCode,
              secCode: doc.secCode,
            });
          }
        }
      }

      if (found.length > 0) break;
    }

    // 提出日で降順ソート
    found.sort((a, b) => (b.submitDateTime || '').localeCompare(a.submitDateTime || ''));

    res.json({ success: true, documents: found });
  } catch (err) {
    console.error('EDINET search error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
};

function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function fetchDocList(date, apiKey) {
  const url = `https://api.edinet-fsa.go.jp/api/v2/documents.json?date=${date}&type=2&Subscription-Key=${apiKey}`;
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' },
      timeout: 8000,
    }, (resp) => {
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          resolve(parsed.results || []);
        } catch {
          resolve([]);
        }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
  });
}
