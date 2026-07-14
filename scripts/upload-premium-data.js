#!/usr/bin/env node
/**
 * プレミアムコンテンツデータを Firestore にアップロードするスクリプト
 * Firebase CLI のログイントークンを使用（サービスアカウント不要）
 *
 * 使い方: node scripts/upload-premium-data.js
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

const PROJECT_ID = 'aoyama-nogizaka-activist';

// Firebase CLI のトークンを取得
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

// Firestore REST API でドキュメントを書き込み
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
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(JSON.parse(data));
        } else {
          reject(new Error(`Firestore write failed (${res.statusCode}): ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// JavaScript値をFirestoreのValue形式に変換
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

// HTML からデータを抽出
function extractDataFromHtml(htmlPath, varName) {
  const html = fs.readFileSync(htmlPath, 'utf8');

  // companies 配列を抽出
  if (varName === 'companies') {
    const match = html.match(new RegExp(`const ${varName}\\s*=\\s*(\\[[\\s\\S]*?\\]);\\s*\\n`));
    if (match) {
      try {
        return eval(match[1]);
      } catch (e) {
        console.error(`  Failed to parse ${varName}:`, e.message);
        return null;
      }
    }
  }

  // 単純な定数を抽出
  const match = html.match(new RegExp(`const ${varName}\\s*=\\s*([\\s\\S]*?);\\s*\\n`, 'm'));
  if (match) {
    try {
      return eval(`(${match[1]})`);
    } catch (e) {
      return null;
    }
  }
  return null;
}

async function main() {
  console.log('=== Firebase アクセストークンを取得中... ===');
  const token = await getAccessToken();
  console.log('  トークン取得成功\n');

  const basePath = path.resolve(__dirname, '..');

  // === SaaS ダッシュボード ===
  console.log('=== SaaS ダッシュボードデータをアップロード中... ===');
  const saasJsonPath = path.join(basePath, 'data', 'saas-companies.json');
  if (fs.existsSync(saasJsonPath)) {
    const saasCompanies = JSON.parse(fs.readFileSync(saasJsonPath, 'utf8'));
    await firestoreWrite(token, 'premiumContent', 'saas-companies', {
      companies: saasCompanies,
      count: saasCompanies.length,
      updatedAt: new Date().toISOString(),
    });
    console.log(`  saas-companies: ${saasCompanies.length}社のデータをアップロード`);
  } else {
    console.log('  ⚠ data/saas-companies.json が見つかりません（スキップ）');
  }

  // === 外食産業ダッシュボード ===
  console.log('\n=== 外食産業ダッシュボードデータをアップロード中... ===');
  const foodJsonPath = path.join(basePath, 'data', 'food-companies.json');
  if (fs.existsSync(foodJsonPath)) {
    const foodCompanies = JSON.parse(fs.readFileSync(foodJsonPath, 'utf8'));
    await firestoreWrite(token, 'premiumContent', 'food-companies', {
      companies: foodCompanies,
      count: foodCompanies.length,
      updatedAt: new Date().toISOString(),
    });
    console.log(`  food-companies: ${foodCompanies.length}社のデータをアップロード`);
  } else {
    console.log('  ⚠ data/food-companies.json が見つかりません（スキップ）');
  }

  // === 広告セクター ダッシュボード ===
  console.log('\n=== 広告セクターダッシュボードデータをアップロード中... ===');
  const adJsonPath = path.join(basePath, 'data', 'ad-agency-companies.json');
  if (fs.existsSync(adJsonPath)) {
    const adCompanies = JSON.parse(fs.readFileSync(adJsonPath, 'utf8'));
    await firestoreWrite(token, 'premiumContent', 'ad-agency-companies', {
      companies: adCompanies,
      count: adCompanies.length,
      updatedAt: new Date().toISOString(),
    });
    console.log(`  ad-agency-companies: ${adCompanies.length}社のデータをアップロード`);
  } else {
    console.log('  ⚠ data/ad-agency-companies.json が見つかりません（スキップ）');
  }

  // === エンタメセクター ダッシュボード ===
  console.log('\n=== エンタメセクターダッシュボードデータをアップロード中... ===');
  const entJsonPath = path.join(basePath, 'data', 'entertainment-companies.json');
  if (fs.existsSync(entJsonPath)) {
    const entCompanies = JSON.parse(fs.readFileSync(entJsonPath, 'utf8'));
    await firestoreWrite(token, 'premiumContent', 'entertainment-companies', {
      companies: entCompanies,
      count: entCompanies.length,
      updatedAt: new Date().toISOString(),
    });
    console.log(`  entertainment-companies: ${entCompanies.length}社のデータをアップロード`);
  } else {
    console.log('  ⚠ data/entertainment-companies.json が見つかりません（スキップ）');
  }

  // === デジタルメディアセクター ダッシュボード ===
  console.log('\n=== デジタルメディアセクターダッシュボードデータをアップロード中... ===');
  const mediaJsonPath = path.join(basePath, 'data', 'digital-media-companies.json');
  if (fs.existsSync(mediaJsonPath)) {
    const mediaCompanies = JSON.parse(fs.readFileSync(mediaJsonPath, 'utf8'));
    await firestoreWrite(token, 'premiumContent', 'digital-media-companies', {
      companies: mediaCompanies,
      count: mediaCompanies.length,
      updatedAt: new Date().toISOString(),
    });
    console.log(`  digital-media-companies: ${mediaCompanies.length}社のデータをアップロード`);
  } else {
    console.log('  ⚠ data/digital-media-companies.json が見つかりません（スキップ）');
  }

  console.log('\n=== 全データのアップロードが完了しました ===');
}

main().catch(e => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
