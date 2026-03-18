const https = require('https');
const zlib = require('zlib');

/**
 * 土地明細抽出API
 * 有価証券報告書のXBRL注記（有形固定資産等明細表）から土地の個別明細を抽出し、
 * 国土交通省の地価公示データで時価推定を行う
 * GET /api/edinet/land-parcels/:docID?apiKey=xxx
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const docID = req.query.docID;
  const apiKey = req.query.apiKey || process.env.EDINET_API_KEY;

  if (!docID) return res.status(400).json({ success: false, error: 'docIDが必要です' });
  if (!apiKey) return res.status(400).json({ success: false, error: 'EDINET APIキーが必要です' });

  try {
    // XBRL ZIP ダウンロード
    const zipBuffer = await downloadDoc(docID, apiKey);
    const entries = readZipEntries(zipBuffer);

    // HTML注記ファイルから固定資産明細表を探す
    const noteEntries = entries.filter(e =>
      /\.htm(l)?$/i.test(e.name) && /PublicDoc/i.test(e.name)
    );

    let parcels = [];
    let landBookValue = null;

    // XBRLから土地簿価を取得
    const xbrlEntry = entries.find(e =>
      /\.xbrl$/i.test(e.name) &&
      !e.name.includes('AuditDoc') &&
      (e.name.includes('jpcrp') || e.name.includes('jplvh') || /PublicDoc/i.test(e.name))
    ) || entries.find(e => /\.xbrl$/i.test(e.name) && /PublicDoc/i.test(e.name));

    if (xbrlEntry) {
      const xbrlXml = extractEntry(zipBuffer, xbrlEntry).toString('utf8');
      landBookValue = findXbrlValue(xbrlXml, 'Land', 'Instant');
      if (landBookValue != null) landBookValue = landBookValue / 1000000; // 百万円
    }

    // HTML注記から固定資産明細表を解析
    for (const ne of noteEntries) {
      try {
        const html = extractEntry(zipBuffer, ne).toString('utf8');
        const found = parseLandParcels(html);
        if (found.length > 0) {
          parcels = parcels.concat(found);
        }
      } catch {}
    }

    // 重複除去（同じ所在地のものをマージ）
    parcels = deduplicateParcels(parcels);

    // ヒューリスティック: 簿価単位の自動修正
    // 簿価/面積が異常に大きい場合、千円単位が百万円として誤認されている可能性が高い
    for (const p of parcels) {
      if (p.bookValue && p.area && p.area > 0) {
        const bvPerSqm = p.bookValue * 1000000 / p.area; // 百万円→円に変換して/㎡
        if (bvPerSqm > 50000000) { // 5000万円/㎡超は異常（銀座でも最大3000万円/㎡程度）
          p.bookValue = p.bookValue / 1000; // 千円→百万円に変換
          p._unitCorrected = true;
        }
      }
    }

    // 国土交通省 地価公示データで時価推定
    if (parcels.length > 0) {
      await estimateLandPrices(parcels);
    }

    // 合計計算
    let totalBookValue = 0;
    let totalEstimatedValue = 0;
    let estimatedCount = 0;
    for (const p of parcels) {
      if (p.bookValue) totalBookValue += p.bookValue;
      if (p.estimatedValue) {
        totalEstimatedValue += p.estimatedValue;
        estimatedCount++;
      }
    }

    // 全体の土地含み益推定（明細表の簿価合計とXBRLの土地簿価を比較して按分）
    let totalEstimatedGain = null;
    let gainMethod = null;
    if (estimatedCount > 0 && totalBookValue > 0) {
      const parcelGain = totalEstimatedValue - totalBookValue;
      if (landBookValue && landBookValue > totalBookValue) {
        // 明細表に載っていない土地もあるので按分
        const ratio = totalEstimatedValue / totalBookValue;
        totalEstimatedGain = Math.round(landBookValue * ratio - landBookValue);
        gainMethod = `都道府県別公示地価×用途別係数による概算（${estimatedCount}件, 明細簿価${Math.round(totalBookValue)}→全土地${Math.round(landBookValue)}百万円に按分）`;
      } else {
        totalEstimatedGain = Math.round(parcelGain);
        gainMethod = `都道府県別公示地価×用途別係数による概算（${estimatedCount}/${parcels.length}件推定）`;
      }
    }

    res.json({
      success: true,
      data: {
        landBookValue,
        parcels,
        totalBookValue: Math.round(totalBookValue),
        totalEstimatedValue: Math.round(totalEstimatedValue),
        totalEstimatedGain,
        gainMethod,
        parcelCount: parcels.length,
        estimatedCount,
      }
    });

  } catch (err) {
    console.error('Land parcels error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
};

// ═══════════════════════════════════════════════════════════════
// 固定資産明細表パーサー
// ═══════════════════════════════════════════════════════════════

function parseLandParcels(html) {
  // まず「主要な設備の状況」テーブルから抽出を試行（最も構造化されたデータ）
  const facilityParcels = parseFacilityTable(html);
  if (facilityParcels.length > 0) {
    return facilityParcels;
  }

  // 「有形固定資産等明細表」から抽出を試行
  const patterns = [
    '有形固定資産等明細表',
    '有形固定資産明細表',
    '固定資産明細表',
  ];

  let sectionStart = -1;
  for (const pat of patterns) {
    sectionStart = html.indexOf(pat);
    if (sectionStart !== -1) break;
  }
  if (sectionStart === -1) return [];

  const searchEnd = Math.min(html.length, sectionStart + 30000);
  const searchHtml = html.substring(sectionStart, searchEnd);

  const landIdx = searchHtml.indexOf('土地');
  if (landIdx === -1) return [];

  // 固定資産明細表から直接抽出を試行
  const directParcels = parseDirectLandEntries(searchHtml, landIdx);
  if (directParcels.length > 0) {
    return directParcels;
  }

  return [];
}

/**
 * 「主要な設備の状況」テーブルから土地データを抽出
 * このテーブルには事業所ごとの土地面積と帳簿価額が含まれる
 */
function parseFacilityTable(html) {
  const allParcels = [];

  // 「主要な設備の状況」セクションを探す
  const facilityPatterns = ['主要な設備の状況', '設備の状況'];
  let fIdx = -1;
  for (const p of facilityPatterns) {
    fIdx = html.indexOf(p);
    if (fIdx !== -1) break;
  }
  if (fIdx === -1) return allParcels;

  // セクション内のテーブルを全て取得
  const searchRange = html.substring(fIdx, Math.min(html.length, fIdx + 80000));
  const tableRegex = /<table[\s\S]*?<\/table>/gi;
  let tableMatch;

  while ((tableMatch = tableRegex.exec(searchRange)) !== null) {
    const table = tableMatch[0];
    if (!table.includes('土地')) continue;

    const rows = table.match(/<tr[\s\S]*?<\/tr>/gi);
    if (!rows) continue;

    // ヘッダー解析（2段ヘッダー対応）
    // メインヘッダー例: [事業所名(所在地), セグメント, 設備内容, 帳簿価額(百万円), 従業員数]
    // サブヘッダー例:   [土地(面積千㎡), 建物, 機械装置, 合計]
    // → サブヘッダーは「帳簿価額」列の中に展開される
    let colMap = {};
    let headerFound = false;
    let areaUnit = 1;
    // テーブル全体から帳簿価額の単位をまず検出（ヘッダーのrowspan等で分散している場合対応）
    const tableText = table.replace(/<[^>]*>/g, '');
    let bvUnitDivisor = tableText.includes('千円') ? 1000 : 1; // 帳簿価額の単位: 百万円=1, 千円=1000
    let landDataCol = null; // データ行での土地列インデックス

    for (let ri = 0; ri < rows.length; ri++) {
      const cells = extractCells(rows[ri]);
      const cellTexts = cells.map(c => cleanCellText(c));

      if (!headerFound) {
        // メインヘッダー行の解析
        for (let ci = 0; ci < cellTexts.length; ci++) {
          const t = cellTexts[ci];
          if (t.includes('事業所名') || t.includes('事業所') ||
              (t.includes('子会社') && t.includes('所在地'))) {
            colMap.name = ci;
            if (t.includes('所在地')) colMap.nameIncludesAddress = true;
          }
          if ((t.includes('所在地') || t.includes('住所')) && !t.includes('事業所')) {
            colMap.address = ci;
          }
          if (t.includes('帳簿価額')) {
            colMap.bookValueCol = ci;
            // 帳簿価額の単位を検出（千円 or 百万円）
            if (t.includes('千円')) bvUnitDivisor = 1000;
            else if (t.includes('百万円')) bvUnitDivisor = 1;
          }
        }

        // サブヘッダー行：「土地(面積千㎡)」を含む
        const landCellIdx = cellTexts.findIndex(t => t.includes('土地'));
        if (landCellIdx !== -1) {
          const landText = cellTexts[landCellIdx];
          if (landText.includes('千㎡') || landText.includes('千m')) areaUnit = 1000;
          // サブヘッダー行にも単位表記がある場合（「千円」がヘッダー全体に含まれるか確認）
          const rowText = cellTexts.join('');
          if (bvUnitDivisor === 1 && rowText.includes('千円')) bvUnitDivisor = 1000;

          // データ行での列位置を計算
          // サブヘッダーは帳簿価額列の展開なので: landDataCol = bookValueCol + landCellIdx
          if (colMap.bookValueCol != null) {
            landDataCol = colMap.bookValueCol + landCellIdx;
          } else {
            // 帳簿価額列が見つからない場合、メインヘッダー列数とサブヘッダー列数の差分で推定
            // データ行の列数 = メインヘッダー非展開列数 + サブヘッダー列数
            landDataCol = landCellIdx + (colMap.name != null ? colMap.name + 1 : 0);
            // フォールバック：名前列より後ろの最初の数値列を探す（後でデータ行で調整）
          }
        }

        if (colMap.name != null && landDataCol != null) {
          headerFound = true;
          continue;
        }
        continue;
      }

      // データ行を解析
      // パターンA: 1〜2セル行（面積のみ）→ 前のparcelに面積を追加
      if (cellTexts.length <= 2) {
        const singleText = cellTexts.join('');
        const areaMatch = singleText.match(/[\(（]\s*([0-9,]+(?:\.[0-9]+)?)\s*[\)）]/);
        if (areaMatch && allParcels.length > 0) {
          const lastParcel = allParcels[allParcels.length - 1];
          if (lastParcel.area == null) {
            lastParcel.area = parseFloat(areaMatch[1].replace(/,/g, '')) * areaUnit;
          }
        }
        continue;
      }

      if (cellTexts.length < 3) continue;

      // 事業所名と所在地を取得
      let name = '';
      let address = '';
      const rawName = colMap.name != null && colMap.name < cellTexts.length ? cellTexts[colMap.name] : '';

      if (colMap.nameIncludesAddress) {
        const addrInParen = rawName.match(/(.+?)(?:（|[\(])((?:北海道|東京都|大阪府|京都府|.{2,3}県)[^）\)]*?)(?:）|[\)])/);
        if (addrInParen) {
          name = addrInParen[1];
          address = addrInParen[2];
        } else {
          name = rawName;
          const prefM = rawName.match(/(北海道|東京都|大阪府|京都府|.{2,3}県)[^）\)（\(]*/);
          if (prefM) address = prefM[0];
        }
      } else {
        name = rawName;
        address = colMap.address != null && colMap.address < cellTexts.length ? cellTexts[colMap.address] : '';
      }

      if (!address && !name) continue;
      if ((address + name).includes('事業所') || (address + name).includes('セグメント')) continue;

      const prefMatch = (address || name).match(/(北海道|東京都|大阪府|京都府|.{2,3}県)/);
      if (!prefMatch) continue; // 海外拠点スキップ

      // 土地の帳簿価額と面積を取得
      let bookValue = null;
      let area = null;

      if (landDataCol != null && landDataCol < cellTexts.length) {
        const landText = cellTexts[landDataCol];

        // パターンB: 「109,381(386)(※118)」— 帳簿価額と面積が同一セル
        const combinedMatch = landText.match(/^([0-9,]+)[\(（]\s*([0-9,]+(?:\.[0-9]+)?)\s*[\)）]/);
        if (combinedMatch) {
          bookValue = parseFloat(combinedMatch[1].replace(/,/g, ''));
          area = parseFloat(combinedMatch[2].replace(/,/g, '')) * areaUnit;
        } else {
          // パターンA: 帳簿価額のみ（面積は次の1セル行）
          const bvStr = landText.replace(/,/g, '').replace(/[^0-9.△\-]/g, '');
          bookValue = parseFloat(bvStr.replace('△', '-'));
        }
        if (isNaN(bookValue)) bookValue = null;
        if (isNaN(area)) area = null;
        // 単位を百万円に統一
        if (bookValue != null && bvUnitDivisor > 1) {
          bookValue = bookValue / bvUnitDivisor;
        }
      }

      allParcels.push({
        name: name || '',
        address: address || name,
        prefecture: prefMatch[1],
        area: area,
        bookValue: bookValue,
        estimatedValue: null,
        estimatedPricePerSqm: null,
        source: '主要な設備の状況',
      });
    }
  }

  return allParcels;
}

function cleanCellText(cellHtml) {
  return cellHtml.replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, ' ').replace(/&#160;/g, ' ').replace(/\s+/g, '').trim();
}

/**
 * 固定資産明細表から直接的に土地エントリを抽出
 */
function parseDirectLandEntries(html, landIdx) {
  const parcels = [];

  // 土地の行の後にある注記番号の注記から所在地を抽出
  // 注記パターン: ※1 土地の内容は以下のとおり...
  const noteSection = html.substring(landIdx);

  // 所在地パターンを直接探す
  const addressRegex = /((?:北海道|東京都|大阪府|京都府|.{2,3}県)[^\s<,、]{2,50})/g;
  let addrMatch;
  const addresses = [];
  const checkRange = noteSection.substring(0, 5000);

  while ((addrMatch = addressRegex.exec(checkRange)) !== null) {
    addresses.push(addrMatch[1]);
  }

  // 数値（面積・金額）を近くから抽出
  for (const addr of addresses.slice(0, 20)) { // 最大20件
    const prefMatch = addr.match(/(北海道|東京都|大阪府|京都府|.{2,3}県)/);
    parcels.push({
      name: '',
      address: addr,
      prefecture: prefMatch ? prefMatch[1] : null,
      area: null,
      bookValue: null,
      estimatedValue: null,
      estimatedPricePerSqm: null,
      source: '固定資産明細表注記',
    });
  }

  return parcels;
}

function extractCells(rowHtml) {
  const cells = [];
  const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  let m;
  while ((m = cellRegex.exec(rowHtml)) !== null) {
    cells.push(m[1]);
  }
  return cells;
}

function deduplicateParcels(parcels) {
  const seen = new Map();
  for (const p of parcels) {
    const key = (p.address || p.name).replace(/\s+/g, '');
    if (!seen.has(key)) {
      seen.set(key, p);
    } else {
      // より詳細なデータで上書き
      const existing = seen.get(key);
      if (!existing.area && p.area) existing.area = p.area;
      if (!existing.bookValue && p.bookValue) existing.bookValue = p.bookValue;
    }
  }
  return Array.from(seen.values());
}

// ═══════════════════════════════════════════════════════════════
// 国土交通省 地価公示データによる時価推定
// ═══════════════════════════════════════════════════════════════

// 都道府県別平均公示地価（円/㎡, 2025年基準地価ベース）
const PREF_LAND_PRICES = {
  '北海道':50774,'青森県':19915,'岩手県':30061,'宮城県':125030,'秋田県':16718,
  '山形県':24876,'福島県':28324,'茨城県':41669,'栃木県':41272,'群馬県':42254,
  '埼玉県':162186,'千葉県':133177,'東京都':1301762,'神奈川県':345138,'新潟県':36066,
  '富山県':46816,'石川県':73132,'福井県':40650,'山梨県':26100,'長野県':33548,
  '岐阜県':46058,'静岡県':84459,'愛知県':240567,'三重県':37752,'滋賀県':59729,
  '京都府':316989,'大阪府':439556,'兵庫県':175693,'奈良県':75505,'和歌山県':44961,
  '鳥取県':23745,'島根県':23525,'岡山県':50913,'広島県':115781,'山口県':29819,
  '徳島県':35422,'香川県':40309,'愛媛県':47576,'高知県':41487,'福岡県':178296,
  '佐賀県':30437,'長崎県':47062,'熊本県':62059,'大分県':35474,'宮崎県':30477,
  '鹿児島県':42099,'沖縄県':114814,
};

// 用途推定ルール（事業所名・所在地からキーワードマッチ）
// 工場キーワードを先に判定（「本社工場」のようなケースで工場を優先）
const LAND_USE_RULES = [
  { key: 'factory',     label: '工場・倉庫', rate: 0.5, keywords: ['工場','製造所','製作所','倉庫','物流','配送','プラント','製鉄所','製油所','精製所'] },
  { key: 'commercial',  label: '商業施設',   rate: 0.8, keywords: ['店舗','ショッピング','商業','モール','販売','百貨店'] },
  { key: 'residential', label: '住宅・寮',   rate: 1.0, keywords: ['社宅','寮','住宅','マンション','レジデンス'] },
  { key: 'idle',        label: '遊休地',     rate: 0.6, keywords: ['遊休','未利用','跡地'] },
  { key: 'office',      label: 'オフィス',   rate: 0.9, keywords: ['本社','本店','事務所','オフィス','事業所','支店','支社','営業所'] },
];

function guessLandUse(name, address, area) {
  const text = ((name || '') + ' ' + (address || '')).toLowerCase();

  // 面積が10万㎡超は工場・大規模施設とみなす（オフィスビルの敷地は通常1万㎡以下）
  if (area && area > 100000) {
    return { key: 'factory', label: '工場・倉庫', rate: 0.5, keywords: [] };
  }

  for (const r of LAND_USE_RULES) {
    for (const kw of r.keywords) {
      if (text.includes(kw)) return r;
    }
  }
  return { key: 'other', label: 'その他', rate: 0.5, keywords: [] };
}

async function estimateLandPrices(parcels) {
  // 都道府県別平均公示地価（円/㎡, 2025年基準地価）を使って推定
  for (const parcel of parcels) {
    // 用途推定（面積も考慮）
    const use = guessLandUse(parcel.name, parcel.address, parcel.area);
    parcel.useType = use.key;
    parcel.useLabel = use.label;
    parcel.adjustmentRate = use.rate;

    if (!parcel.prefecture || !parcel.area) continue;
    const pricePerSqm = PREF_LAND_PRICES[parcel.prefecture];
    if (!pricePerSqm) continue;

    parcel.basePricePerSqm = pricePerSqm; // 調整前の全用途平均単価
    const adjustedPrice = Math.round(pricePerSqm * use.rate);
    parcel.estimatedPricePerSqm = adjustedPrice;
    parcel.estimatedValue = Math.round(parcel.area * adjustedPrice / 1000000); // 百万円
  }
}


// ═══════════════════════════════════════════════════════════════
// XBRL ユーティリティ（[docID].jsと共通）
// ═══════════════════════════════════════════════════════════════

function findXbrlValue(xml, elementName, typeHint) {
  const regex = new RegExp(
    `<[^>]*?:?${elementName}[^>]*contextRef="([^"]*)"[^>]*>([^<]+)<`, 'gi'
  );
  let bestVal = null;
  let bestPriority = -1;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    const ctx = match[1];
    const rawVal = match[2].replace(/,/g, '').trim();
    const num = parseFloat(rawVal);
    if (isNaN(num)) continue;
    let priority = 0;
    if (typeHint && ctx.toLowerCase().includes(typeHint.toLowerCase())) priority += 10;
    if (/CurrentYear/i.test(ctx) || /Current(?!.*Prior)/i.test(ctx)) priority += 5;
    if (!/NonConsolidated/i.test(ctx)) priority += 3;
    if (!/Member/i.test(ctx) || /ConsolidatedMember/i.test(ctx)) priority += 1;
    if (priority > bestPriority) { bestPriority = priority; bestVal = num; }
  }
  return bestVal;
}

// ═══════════════════════════════════════════════════════════════
// ZIP パーサー
// ═══════════════════════════════════════════════════════════════

function readZipEntries(buffer) {
  let eocdOffset = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocdOffset = i; break; }
  }
  if (eocdOffset === -1) throw new Error('ZIPファイルではありません');
  const cdOffset = buffer.readUInt32LE(eocdOffset + 16);
  const cdEntries = buffer.readUInt16LE(eocdOffset + 10);
  const entries = [];
  let pos = cdOffset;
  for (let i = 0; i < cdEntries && pos < eocdOffset; i++) {
    if (buffer.readUInt32LE(pos) !== 0x02014b50) break;
    const compression = buffer.readUInt16LE(pos + 10);
    const compSize = buffer.readUInt32LE(pos + 20);
    const uncompSize = buffer.readUInt32LE(pos + 24);
    const nameLen = buffer.readUInt16LE(pos + 28);
    const extraLen = buffer.readUInt16LE(pos + 30);
    const commentLen = buffer.readUInt16LE(pos + 32);
    const localOffset = buffer.readUInt32LE(pos + 42);
    const name = buffer.toString('utf8', pos + 46, pos + 46 + nameLen);
    entries.push({ name, compression, compSize, uncompSize, localOffset });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function extractEntry(buffer, entry) {
  const lh = entry.localOffset;
  if (buffer.readUInt32LE(lh) !== 0x04034b50) throw new Error('不正なローカルヘッダ');
  const nameLen = buffer.readUInt16LE(lh + 26);
  const extraLen = buffer.readUInt16LE(lh + 28);
  const dataStart = lh + 30 + nameLen + extraLen;
  const compressed = buffer.slice(dataStart, dataStart + entry.compSize);
  if (entry.compression === 0) return compressed;
  if (entry.compression === 8) return zlib.inflateRawSync(compressed);
  throw new Error('未対応の圧縮形式: ' + entry.compression);
}

function downloadDoc(docID, apiKey) {
  const url = `https://api.edinet-fsa.go.jp/api/v2/documents/${docID}?type=1&Subscription-Key=${apiKey}`;
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/octet-stream' },
      timeout: 30000,
    }, (resp) => {
      if (resp.statusCode >= 300 && resp.statusCode < 400 && resp.headers.location) {
        downloadDoc(resp.headers.location.includes('http') ? resp.headers.location : `https://api.edinet-fsa.go.jp${resp.headers.location}`, apiKey)
          .then(resolve).catch(reject);
        resp.resume();
        return;
      }
      if (resp.statusCode !== 200) { resp.resume(); return reject(new Error(`EDINET returned HTTP ${resp.statusCode}`)); }
      const chunks = [];
      resp.on('data', chunk => chunks.push(chunk));
      resp.on('end', () => resolve(Buffer.concat(chunks)));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('ダウンロードタイムアウト')); });
  });
}
