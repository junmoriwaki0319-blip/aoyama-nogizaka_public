#!/usr/bin/env node
/**
 * 外食産業 月次既存店売上高 過去年度データを Firestore にアップロード
 *
 * 全データは以下の公開ソースから取得した実数値のみ:
 *   - 各社IR月次開示（ゼンショー zensho.co.jp、マクドナルド mcd-holdings.co.jp）
 *   - Nautical Star Strategy & Analysis 既存店前年比速報 (nauticalstar-sa.com)
 * 推定値は一切含まない
 *
 * 使い方: node scripts/upload-food-monthly-history.js
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

const PROJECT_ID = 'aoyama-nogizaka-activist';

// ============================================================
// FY2023 (2023年4月〜2024年3月) 既存店売上高前年比データ
// 配列順: [4月, 5月, 6月, 7月, 8月, 9月, 10月, 11月, 12月, 1月, 2月, 3月]
// 全データ出所: Nautical Star + 各社IR
// ============================================================
const SSS_FY2023 = {
  // マクドナルド - 出所: mcd-holdings.co.jp IR
  '2702': [109.1, 105.2, 105.7, 106.6, 108.4, 108.6, 103.9, 103.3, 108.5, 105.4, 105.8, 109.6],
  // ゼンショー(すき家) - 出所: zensho.co.jp IR 2024年3月期
  '7550': [122.0, 115.2, 119.8, 116.9, 118.8, 116.6, 112.5, 117.2, 110.3, 112.7, 115.0, 112.7],
  // FOOD&LIFE(スシロー) - 出所: Nautical Star 各月レポート
  '3563': [84.8, 89.0, 95.0, 104.2, 114.4, 109.9, 111.8, 130.3, 125.4, 117.7, 104.0, 109.3],
  // すかいらーく(ガスト等) - 出所: Nautical Star 各月レポート
  '3197': [118.9, 114.6, 112.1, 118.4, 118.6, 119.5, 111.6, 113.8, 113.0, 110.5, 114.5, 114.4],
  // トリドール(丸亀製麺) - 出所: Nautical Star 各月レポート
  '3397': [115.7, 110.8, 108.2, 117.6, 112.7, 107.3, 109.6, 108.5, 107.0, 109.5, 113.6, 113.5],
  // サイゼリヤ - 出所: Nautical Star 各月レポート
  '7581': [114.7, 112.8, 113.7, 122.7, 126.4, 121.7, 118.7, 121.3, 123.4, 122.6, 124.0, 129.1],
  // 吉野家 - 出所: Nautical Star 各月レポート
  '9861': [107.6, 107.9, 116.6, 111.0, 114.9, 116.2, 107.7, 108.0, 106.9, 110.4, 108.7, 106.8],
  // コロワイド(CW全ブランド) - 出所: Nautical Star 各月レポート
  '7616': [117.1, 110.8, 110.8, 116.6, 112.5, 111.2, 108.5, 107.4, 109.9, 105.7, 106.6, 106.6],
  // 王将フードサービス(餃子の王将) - 出所: Nautical Star 各月レポート
  '9936': [112.6, 104.5, 108.5, 109.9, 110.9, 112.0, 105.5, 109.9, 105.6, 105.4, 107.6, 106.4],
  // 物語コーポレーション(焼肉きんぐ) - 出所: Nautical Star 各月レポート
  '3097': [118.4, 110.7, 111.5, 110.4, 116.2, 109.5, 110.5, 108.0, 111.3, 109.4, 110.5, 111.6],
  // くら寿司 - 出所: Nautical Star 各月レポート
  '2695': [101.8, 103.8, 98.8, 109.5, 110.1, 110.3, 105.8, 98.6, 102.5, 103.2, 106.6, 123.0],
  // ロイヤル(ロイヤルホスト) - 出所: Nautical Star 各月レポート
  '8179': [121.1, 108.0, 112.4, 119.1, 112.0, 110.6, 108.6, 106.0, 100.5, 104.6, 108.8, 105.4],
  // 壱番屋(CoCo壱番屋) - 出所: Nautical Star 各月レポート
  '7630': [114.4, 110.0, 111.9, 115.1, 115.7, 118.1, 113.6, 113.6, 106.1, 107.5, 106.7, 110.5],
  // コメダ - 出所: Nautical Star 各月レポート
  '3543': [111.3, 119.9, 112.1, 121.5, 117.0, 124.9, 112.0, 110.5, 108.5, 112.5, 110.0, 109.4],
  // クリエイト・レストランツ - データなし(Nautical Starに掲載なし)
};

// ============================================================
// FY2022 (2022年4月〜2023年3月) 既存店売上高前年比データ
// 全データ出所: Nautical Star + 各社IR
// ============================================================
const SSS_FY2022 = {
  // マクドナルド - 出所: Nautical Star 各月 + mcd-holdings.co.jp IR(1-3月)
  '2702': [111.3, 105.1, 110.2, 108.1, 103.3, 104.9, 109.0, 113.3, 115.2, 114.6, 103.0, 106.4],
  // ゼンショー(すき家) - 出所: zensho.co.jp IR 2023年3月期
  '7550': [108.3, 106.3, 109.2, 107.4, 108.1, 113.4, 111.3, 111.3, 105.4, 107.9, 109.8, 113.4],
  // FOOD&LIFE(スシロー) - 出所: Nautical Star 各月レポート
  '3563': [105.3, 107.8, 97.5, 89.8, 93.7, 91.9, 81.5, 74.8, 77.7, 89.6, 99.8, 87.4],
  // すかいらーく(ガスト等) - 出所: Nautical Star 各月レポート
  '3197': [110.6, 122.0, 126.0, 118.0, 126.0, 131.4, 119.4, 107.6, 102.1, 121.7, 138.8, 126.6],
  // トリドール(丸亀製麺) - 出所: Nautical Star 各月レポート
  '3397': [105.7, 114.5, 107.5, 102.1, 118.6, 110.2, 109.3, 106.6, 117.3, 114.8, 123.1, 113.0],
  // サイゼリヤ - 出所: Nautical Star 各月レポート
  '7581': [124.4, 138.8, 136.1, 124.0, 138.4, 147.6, 120.8, 110.0, 107.0, 119.3, 133.7, 117.2],
  // 吉野家 - 出所: Nautical Star 各月レポート
  '9861': [111.6, 109.8, 106.3, 110.5, 99.4, 102.2, 108.1, 103.0, 104.5, 106.4, 108.1, 107.4],
  // コロワイド(CW全ブランド) - 出所: Nautical Star 各月レポート
  '7616': [117.2, 123.8, 121.1, 113.5, 127.9, 137.8, 113.3, 103.3, 105.6, 117.6, 136.5, 123.8],
  // 王将フードサービス(餃子の王将) - 出所: Nautical Star 各月レポート
  '9936': [106.0, 117.5, 107.1, 106.7, 110.2, 112.5, 108.1, 104.7, 105.5, 108.5, 105.0, 112.5],
  // 物語コーポレーション(焼肉きんぐ) - 出所: Nautical Star 各月レポート
  '3097': [121.2, 138.2, 134.9, 120.9, 131.7, 142.2, 116.2, 100.6, 101.9, 119.6, 136.7, 118.0],
  // くら寿司 - 出所: Nautical Star 各月レポート
  '2695': [110.0, 113.6, 108.6, 101.3, 106.8, 118.3, 106.1, 96.4, 94.9, 99.0, 114.7, 97.8],
  // ロイヤル(ロイヤルホスト) - 出所: Nautical Star 各月レポート (FRカテゴリ Royal HD)
  // Note: 2022年4月はすかいらーく内にRoyal別掲あり、一部月はRoyal HD表記
  '8179': [null, null, null, null, null, null, 123.3, 111.3, 120.9, 114.2, 139.2, 124.6],
  // 壱番屋(CoCo壱番屋) - 出所: Nautical Star 各月レポート
  '7630': [99.0, 103.9, 107.8, 108.1, 108.8, 109.3, 107.7, 106.4, 105.8, 105.9, 119.7, 115.2],
  // コメダ - 出所: Nautical Star 各月レポート
  '3543': [103.7, 109.0, 107.9, 96.1, 105.8, 110.2, 105.3, 106.6, 106.1, 113.0, 110.8, 111.1],
};

// ============================================================
// Main
// ============================================================
async function main() {
  console.log('=== Firebase アクセストークンを取得中... ===');
  const token = await getAccessToken();
  console.log('  トークン取得成功\n');

  // 現在のFirestoreから企業データを取得
  console.log('=== Firestore企業データを取得中... ===');
  const compData = await firestoreRead(token, 'premiumContent', 'food-companies');
  if (!compData || !compData.companies) {
    throw new Error('企業データが見つかりません');
  }

  const companies = fromFirestoreValue(compData.companies);
  console.log(`  ${companies.length}社のデータを取得\n`);

  // 各企業にmonthlyHistoryを追加（実データのみ）
  let updatedCount = 0;
  companies.forEach(c => {
    const fy23 = SSS_FY2023[c.code];
    const fy22 = SSS_FY2022[c.code];
    if (fy23 || fy22) {
      c.monthlyHistory = {};
      if (fy22) {
        c.monthlyHistory['FY2022'] = { sameStoreSales: fy22 };
      }
      if (fy23) {
        c.monthlyHistory['FY2023'] = { sameStoreSales: fy23 };
      }
      updatedCount++;
      const nullCount22 = fy22 ? fy22.filter(v => v === null).length : 12;
      const nullCount23 = fy23 ? fy23.filter(v => v === null).length : 12;
      console.log(`  ✓ ${c.code} ${c.name}: FY2022(${fy22 ? 12 - nullCount22 : 0}ヶ月) FY2023(${fy23 ? 12 - nullCount23 : 0}ヶ月)`);
    } else {
      // monthlyHistoryをクリア（推定データ削除）
      if (c.monthlyHistory) {
        delete c.monthlyHistory;
        console.log(`  ✗ ${c.code} ${c.name}: monthlyHistory削除（実データなし）`);
      }
    }
  });

  console.log(`\n  ${updatedCount}/${companies.length}社にmonthlyHistoryを設定\n`);

  // Firestoreに書き戻し
  console.log('=== Firestoreにアップロード中... ===');
  await firestoreWrite(token, 'premiumContent', 'food-companies', {
    companies: companies,
    count: companies.length
  });
  console.log('  premiumContent/food-companies を更新しました');

  console.log('\n=== 完了 ===');
}

// ============================================================
// Firebase Helpers (same as before)
// ============================================================
function getAccessToken() {
  const configPath = path.join(process.env.HOME || process.env.USERPROFILE, '.config', 'configstore', 'firebase-tools.json');
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const refreshToken = config.tokens.refresh_token;
  return new Promise((resolve, reject) => {
    const postData = `grant_type=refresh_token&refresh_token=${encodeURIComponent(refreshToken)}&client_id=563584335869-fgrhgmd47bqnekij5i8b5pr03ho849e6.apps.googleusercontent.com&client_secret=j9iVZfS8kkCEFUPaAeJV0sAi`;
    const req = https.request({
      hostname: 'oauth2.googleapis.com', path: '/token', method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) }
    }, (res) => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => { const j = JSON.parse(data); j.access_token ? resolve(j.access_token) : reject(new Error('Token refresh failed')); });
    });
    req.on('error', reject); req.write(postData); req.end();
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
  if (Array.isArray(val)) return { arrayValue: { values: val.map(toFirestoreValue) } };
  if (typeof val === 'object') {
    const fields = {};
    for (const [k, v] of Object.entries(val)) fields[k] = toFirestoreValue(v);
    return { mapValue: { fields } };
  }
  return { stringValue: String(val) };
}

function fromFirestoreValue(val) {
  if (val.nullValue !== undefined) return null;
  if (val.booleanValue !== undefined) return val.booleanValue;
  if (val.integerValue !== undefined) return parseInt(val.integerValue);
  if (val.doubleValue !== undefined) return val.doubleValue;
  if (val.stringValue !== undefined) return val.stringValue;
  if (val.arrayValue) return (val.arrayValue.values || []).map(fromFirestoreValue);
  if (val.mapValue) {
    const obj = {};
    for (const [k, v] of Object.entries(val.mapValue.fields || {})) obj[k] = fromFirestoreValue(v);
    return obj;
  }
  return null;
}

function firestoreRead(accessToken, collection, docId) {
  return new Promise((resolve, reject) => {
    const docPath = `projects/${PROJECT_ID}/databases/(default)/documents/${collection}/${docId}`;
    https.get({
      hostname: 'firestore.googleapis.com', path: `/v1/${docPath}`,
      headers: { 'Authorization': `Bearer ${accessToken}` }
    }, (res) => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          const doc = JSON.parse(data);
          const result = {};
          for (const [k, v] of Object.entries(doc.fields || {})) result[k] = v;
          resolve(result);
        } else reject(new Error(`Read failed (${res.statusCode}): ${data.substring(0,200)}`));
      });
    }).on('error', reject);
  });
}

function firestoreWrite(accessToken, collection, docId, data) {
  return new Promise((resolve, reject) => {
    const fields = {};
    for (const [key, value] of Object.entries(data)) fields[key] = toFirestoreValue(value);
    const body = JSON.stringify({ fields });
    const docPath = `projects/${PROJECT_ID}/databases/(default)/documents/${collection}/${docId}`;
    const req = https.request({
      hostname: 'firestore.googleapis.com', path: `/v1/${docPath}`, method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
    }, (res) => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(JSON.parse(data));
        else reject(new Error(`Write failed (${res.statusCode}): ${data.substring(0,200)}`));
      });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

main().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
