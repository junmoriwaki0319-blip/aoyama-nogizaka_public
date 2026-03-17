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
        gainMethod = `固定資産明細表+地価公示（${estimatedCount}件推定, 明細簿価${Math.round(totalBookValue)}→全土地${Math.round(landBookValue)}百万円に按分）`;
      } else {
        totalEstimatedGain = Math.round(parcelGain);
        gainMethod = `固定資産明細表+地価公示（${estimatedCount}/${parcels.length}件推定）`;
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
  const parcels = [];

  // 「有形固定資産等明細表」「有形固定資産明細表」セクションを探す
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
  if (sectionStart === -1) return parcels;

  // セクション周辺のテーブルを取得（前後の範囲を広めに）
  const searchEnd = Math.min(html.length, sectionStart + 30000);
  const searchHtml = html.substring(sectionStart, searchEnd);

  // 「土地」の行を探す
  const landIdx = searchHtml.indexOf('土地');
  if (landIdx === -1) return parcels;

  // 土地セクション周辺のテーブルデータを解析
  // 注記にリンクがある場合（「※1」「注1」等）、注記から詳細を取得
  const noteRefMatch = searchHtml.substring(landIdx, landIdx + 500).match(/[※注]\s*(\d+)/);

  // 土地の主要物件注記を探す
  // パターン1: 「土地の主要物件」「主な土地」等のセクション
  const landDetailPatterns = [
    '主要な土地',
    '土地の主要物件',
    '主な土地',
    '主要物件',
    '主要な設備',
    '設備の状況',
  ];

  // 注記セクションまたは「主要な設備の状況」から土地データを抽出
  // 「主要な設備の状況」テーブルは通常、事業所名・所在地・面積・帳簿価額を含む
  const facilityParcels = parseFacilityTable(html);
  if (facilityParcels.length > 0) {
    return facilityParcels;
  }

  // 固定資産明細表から直接抽出を試行
  const directParcels = parseDirectLandEntries(searchHtml, landIdx);
  if (directParcels.length > 0) {
    return directParcels;
  }

  return parcels;
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
          }
        }

        // サブヘッダー行：「土地(面積千㎡)」を含む
        const landCellIdx = cellTexts.findIndex(t => t.includes('土地'));
        if (landCellIdx !== -1) {
          const landText = cellTexts[landCellIdx];
          if (landText.includes('千㎡') || landText.includes('千m')) areaUnit = 1000;

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

// 都道府県コード
const PREF_CODES = {
  '北海道':'01','青森県':'02','岩手県':'03','宮城県':'04','秋田県':'05',
  '山形県':'06','福島県':'07','茨城県':'08','栃木県':'09','群馬県':'10',
  '埼玉県':'11','千葉県':'12','東京都':'13','神奈川県':'14','新潟県':'15',
  '富山県':'16','石川県':'17','福井県':'18','山梨県':'19','長野県':'20',
  '愛知県':'23','三重県':'24','滋賀県':'25','京都府':'26','大阪府':'27',
  '兵庫県':'28','奈良県':'29','和歌山県':'30','鳥取県':'31','島根県':'32',
  '岡山県':'33','広島県':'34','山口県':'35','徳島県':'36','香川県':'37',
  '愛媛県':'38','高知県':'39','福岡県':'40','佐賀県':'41','長崎県':'42',
  '熊本県':'43','大分県':'44','宮崎県':'45','鹿児島県':'46','沖縄県':'47',
  '岐阜県':'21','静岡県':'22',
};

async function estimateLandPrices(parcels) {
  // 都道府県ごとにグループ化
  const prefGroups = {};
  for (const p of parcels) {
    const pref = p.prefecture;
    if (!pref) continue;
    if (!prefGroups[pref]) prefGroups[pref] = [];
    prefGroups[pref].push(p);
  }

  // 各都道府県の地価公示データを取得
  const promises = Object.entries(prefGroups).map(async ([pref, prefParcels]) => {
    const prefCode = PREF_CODES[pref];
    if (!prefCode) return;

    try {
      // 国土交通省 不動産取引価格情報API
      const currentYear = new Date().getFullYear();
      const fromQ = `${currentYear - 2}1`; // 2年前の第1四半期から
      const toQ = `${currentYear}4`;

      const tradeData = await fetchLandPriceAPI(prefCode, fromQ, toQ);

      if (tradeData && tradeData.length > 0) {
        // 住所の市区町村でマッチング
        for (const parcel of prefParcels) {
          const matchedPrice = findBestPriceMatch(parcel, tradeData);
          if (matchedPrice && parcel.area) {
            parcel.estimatedPricePerSqm = matchedPrice;
            parcel.estimatedValue = Math.round(parcel.area * matchedPrice / 1000000); // 百万円
          } else if (matchedPrice && parcel.bookValue) {
            // 面積不明の場合、簿価ベースで倍率推定
            parcel.estimatedPricePerSqm = matchedPrice;
          }
        }
      }
    } catch (err) {
      console.error(`地価取得エラー（${pref}）:`, err.message);
    }
  });

  await Promise.all(promises);
}

function findBestPriceMatch(parcel, tradeData) {
  const address = parcel.address;

  // 市区町村を抽出
  const cityMatch = address.match(/(?:北海道|東京都|大阪府|京都府|.{2,3}県)((?:[^市区町村]{1,5}[市区町村]){1,2})/);
  const city = cityMatch ? cityMatch[1] : '';

  // 町丁目を抽出
  const townMatch = address.match(/[市区町村](.{1,10}?)\d/);
  const town = townMatch ? townMatch[1] : '';

  let bestMatch = null;
  let bestScore = 0;

  for (const trade of tradeData) {
    if (trade.TradePrice == null || trade.Area == null) continue;
    const pricePerSqm = parseInt(trade.TradePrice) / parseFloat(trade.Area);
    if (isNaN(pricePerSqm) || pricePerSqm <= 0) continue;

    // 用途が土地でないものをスキップ（建物付きは含む）
    if (trade.Type && !trade.Type.includes('土地') && !trade.Type.includes('宅地')) continue;

    let score = 0;
    const tradeAddr = (trade.Municipality || '') + (trade.DistrictName || '');

    // 市区町村マッチ
    if (city && tradeAddr.includes(city)) score += 10;
    // 町丁目マッチ
    if (town && tradeAddr.includes(town)) score += 20;
    // 用途マッチ（工業地域 vs 商業地域 etc）
    if (trade.Use && (trade.Use.includes('工場') || trade.Use.includes('事務所'))) score += 2;

    if (score > bestScore) {
      bestScore = score;
      bestMatch = pricePerSqm;
    }
  }

  // マッチスコアが低すぎる場合は都道府県平均を使用
  if (bestScore < 10 && tradeData.length > 0) {
    // 都道府県の中央値を計算
    const prices = tradeData
      .filter(t => t.TradePrice && t.Area && parseFloat(t.Area) > 0)
      .filter(t => !t.Type || t.Type.includes('土地') || t.Type.includes('宅地'))
      .map(t => parseInt(t.TradePrice) / parseFloat(t.Area))
      .filter(p => !isNaN(p) && p > 0)
      .sort((a, b) => a - b);

    if (prices.length > 0) {
      bestMatch = prices[Math.floor(prices.length / 2)]; // 中央値
    }
  }

  return bestMatch ? Math.round(bestMatch) : null;
}

function fetchLandPriceAPI(prefCode, from, to) {
  // 国土交通省 不動産取引価格情報検索API（無料・登録不要）
  const url = `https://www.land.mlit.go.jp/webland/api/TradeListSearch?from=${from}&to=${to}&area=${prefCode}&city=`;

  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    https.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
    }, (resp) => {
      if (resp.statusCode !== 200) {
        resp.resume();
        return resolve([]);
      }
      const chunks = [];
      resp.on('data', chunk => chunks.push(chunk));
      resp.on('end', () => {
        try {
          const json = JSON.parse(Buffer.concat(chunks).toString('utf8'));
          // 土地関連の取引のみ（最大200件でサンプリング）
          const data = (json.data || [])
            .filter(d => d.Type && (d.Type.includes('土地') || d.Type.includes('宅地')))
            .slice(0, 200);
          resolve(data);
        } catch { resolve([]); }
      });
    }).on('error', () => resolve([]));
  });
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
