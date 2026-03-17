const https = require('https');

/**
 * EDINET書類検索API
 * 指定されたEDINETコードの最新の有価証券報告書を検索する
 * GET /api/edinet/search/:code?apiKey=xxx
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const edinetCode = req.query.code;
  const apiKey = req.query.apiKey;

  if (!edinetCode || !/^E\d{5}$/.test(edinetCode)) {
    return res.status(400).json({ success: false, error: 'EDINETコード (E+5桁) を指定してください' });
  }
  if (!apiKey) {
    return res.status(400).json({ success: false, error: 'EDINET APIキーが必要です' });
  }

  try {
    // 過去400日分の日付を生成（直近から遡る）
    const dates = [];
    const today = new Date();
    for (let i = 0; i < 400; i++) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      // 土日をスキップ
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
          if (doc.edinetCode === edinetCode && doc.docTypeCode === '120') {
            found.push({
              docID: doc.docID,
              docDescription: doc.docDescription || '',
              periodStart: doc.periodStart,
              periodEnd: doc.periodEnd,
              submitDateTime: doc.submitDateTime,
              filerName: doc.filerName,
            });
          }
        }
      }

      // 有価証券報告書が見つかったら早期終了
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
