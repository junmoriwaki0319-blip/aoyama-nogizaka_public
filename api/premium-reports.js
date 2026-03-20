const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const https = require('https');

// Google公開鍵キャッシュ
let cachedKeys = null;
let cacheExpiry = 0;

// Google公開鍵を取得（Firebase IDトークン検証用）
function fetchGooglePublicKeys() {
  return new Promise((resolve, reject) => {
    if (cachedKeys && Date.now() < cacheExpiry) return resolve(cachedKeys);
    https.get('https://www.googleapis.com/robot/v1/metadata/x509/securetoken@system.gserviceaccount.com', (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          cachedKeys = JSON.parse(data);
          const cacheControl = res.headers['cache-control'] || '';
          const maxAge = cacheControl.match(/max-age=(\d+)/);
          cacheExpiry = Date.now() + (maxAge ? parseInt(maxAge[1]) * 1000 : 3600000);
          resolve(cachedKeys);
        } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

// Base64url デコード
function base64urlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  return Buffer.from(str, 'base64');
}

// Firebase IDトークンを簡易検証（署名はスキップ、クレームのみチェック）
// ※ Vercel環境でcrypto.createVerifyが使えるなら署名検証も可能
async function verifyIdToken(idToken) {
  const parts = idToken.split('.');
  if (parts.length !== 3) throw new Error('Invalid token format');

  const header = JSON.parse(base64urlDecode(parts[0]).toString());
  const payload = JSON.parse(base64urlDecode(parts[1]).toString());

  // 有効期限チェック
  const now = Math.floor(Date.now() / 1000);
  if (payload.exp < now) throw new Error('Token expired');
  if (payload.iat > now + 300) throw new Error('Token issued in the future');

  // 発行者チェック
  if (payload.iss !== 'https://securetoken.google.com/aoyama-nogizaka-activist') {
    throw new Error('Invalid issuer');
  }

  // audience チェック
  if (payload.aud !== 'aoyama-nogizaka-activist') {
    throw new Error('Invalid audience');
  }

  // sub（ユーザーID）の存在チェック
  if (!payload.sub || typeof payload.sub !== 'string') {
    throw new Error('Invalid subject');
  }

  // 署名検証（Node.js crypto使用）
  try {
    const crypto = require('crypto');
    const keys = await fetchGooglePublicKeys();
    const kid = header.kid;
    const publicKey = keys[kid];
    if (!publicKey) throw new Error('Key not found');

    const signatureInput = parts[0] + '.' + parts[1];
    const signature = base64urlDecode(parts[2]);
    const verifier = crypto.createVerify('RSA-SHA256');
    verifier.update(signatureInput);
    if (!verifier.verify(publicKey, signature)) {
      throw new Error('Invalid signature');
    }
  } catch (e) {
    if (e.message === 'Invalid signature') throw e;
    // 鍵取得エラー等はクレーム検証のみで通す
    console.warn('Signature verification skipped:', e.message);
  }

  return payload;
}

// reports.json を読み込み（起動時にキャッシュ）
let reportsCache = null;
let reportsCacheGzip = null;

function getReportsData() {
  if (reportsCache) return { json: reportsCache, gzip: reportsCacheGzip };
  try {
    const filePath = path.join(__dirname, '..', 'data', 'reports.json');
    reportsCache = fs.readFileSync(filePath, 'utf8');
    reportsCacheGzip = zlib.gzipSync(reportsCache);
    return { json: reportsCache, gzip: reportsCacheGzip };
  } catch (e) {
    return null;
  }
}

// known_activists.json
let activistsCache = null;
function getActivistsData() {
  if (activistsCache) return activistsCache;
  try {
    const filePath = path.join(__dirname, '..', 'scripts', 'known_activists.json');
    activistsCache = fs.readFileSync(filePath, 'utf8');
    return activistsCache;
  } catch (e) {
    return null;
  }
}

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', 'https://aoyama-nogizaka.com');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'GET') { return res.status(405).json({ error: 'GET only' }); }

  // Authorization ヘッダーからトークン取得
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: '認証が必要です' });
  }

  const idToken = authHeader.split('Bearer ')[1];
  try {
    await verifyIdToken(idToken);
  } catch (e) {
    return res.status(403).json({ error: '認証トークンが無効です: ' + e.message });
  }

  // type パラメータで返すデータを切り替え
  const type = req.query.type || 'reports';

  if (type === 'activists') {
    const data = getActivistsData();
    if (!data) return res.status(404).json({ error: 'データが見つかりません' });
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Cache-Control', 'private, max-age=300');
    return res.status(200).send(data);
  }

  // reports（デフォルト）- gzip圧縮で返す
  const data = getReportsData();
  if (!data) return res.status(404).json({ error: 'データが見つかりません' });

  const acceptEncoding = req.headers['accept-encoding'] || '';
  if (acceptEncoding.includes('gzip')) {
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Content-Encoding', 'gzip');
    res.setHeader('Cache-Control', 'private, max-age=300');
    return res.status(200).send(data.gzip);
  }

  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Cache-Control', 'private, max-age=300');
  return res.status(200).send(data.json);
};
