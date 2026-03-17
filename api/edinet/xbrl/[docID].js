const https = require('https');
const zlib = require('zlib');

/**
 * EDINET XBRL解析API
 * 有価証券報告書のXBRLを解析し、財務データ（含み益・土地等）を返す
 * GET /api/edinet/xbrl/:docID?apiKey=xxx
 */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const docID = req.query.docID;
  const apiKey = req.query.apiKey || process.env.EDINET_API_KEY;

  if (!docID) {
    return res.status(400).json({ success: false, error: 'docIDが必要です' });
  }
  if (!apiKey) {
    return res.status(400).json({ success: false, error: 'EDINET APIキーが必要です' });
  }

  try {
    // XBRL ZIP ダウンロード
    const zipBuffer = await downloadDoc(docID, apiKey);

    // ZIP 解析
    const entries = readZipEntries(zipBuffer);

    // インスタンスドキュメント（XBRL）を探す
    const xbrlEntry = entries.find(e =>
      /\.xbrl$/i.test(e.name) &&
      !e.name.includes('AuditDoc') &&
      (e.name.includes('jpcrp') || e.name.includes('jplvh') || /PublicDoc/i.test(e.name))
    ) || entries.find(e => /\.xbrl$/i.test(e.name) && /PublicDoc/i.test(e.name));

    if (!xbrlEntry) {
      return res.json({ success: false, error: 'XBRLファイルが見つかりません' });
    }

    const xbrlXml = extractEntry(zipBuffer, xbrlEntry).toString('utf8');

    // XBRL 解析
    const data = parseXbrl(xbrlXml);

    // 注記テキストブロックから投資不動産・有価証券データを抽出
    const noteEntries = entries.filter(e =>
      /\.htm(l)?$/i.test(e.name) && /PublicDoc/i.test(e.name)
    );
    for (const ne of noteEntries) {
      try {
        const html = extractEntry(zipBuffer, ne).toString('utf8');
        parseNotes(html, data);
      } catch {}
    }

    // 土地含み益推定
    estimateLandGain(data);

    res.json({ success: true, data });
  } catch (err) {
    console.error('EDINET XBRL error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
};

// ═══════════════════════════════════════════════════════════════
// XBRL パーサー
// ═══════════════════════════════════════════════════════════════

function parseXbrl(xml) {
  const result = {};

  // contextRef の優先順位：連結 > 個別、当期 > 前期
  // CurrentYearInstant（連結BS当期末）を優先的に取得

  // ── BS項目（Instant = 期末残高）──
  const bsItems = {
    cashAndDeposits: ['CashAndDeposits'],
    shortTermSecurities: ['ShortTermInvestmentSecurities', 'SecuritiesMKT'],
    investmentSecurities: ['InvestmentSecurities'],
    land: ['Land'],
    netAssets: ['NetAssets', 'TotalNetAssets'],
    shareholdersEquity: ['ShareholdersEquity', 'TotalShareholdersEquity'],
    treasuryShares: ['TreasuryShares', 'TreasurySharesStock', 'TreasuryStock'],
    shortTermBorrowings: ['ShortTermBorrowings', 'ShortTermLoansPayable'],
    currentPortionLongTermDebt: ['CurrentPortionOfLongTermLoansPayable', 'CurrentPortionOfLongTermDebt'],
    currentPortionBonds: ['CurrentPortionOfBondsPayable', 'CurrentPortionOfBonds'],
    longTermBorrowings: ['LongTermLoansPayable', 'LongTermBorrowings', 'LongTermDebt'],
    bondsPayable: ['BondsPayable'],
    landRevaluationReserve: ['RevaluationReserveForLand', 'LandRevaluationExcess', 'LandRevaluationDifference'],
  };

  for (const [key, elements] of Object.entries(bsItems)) {
    for (const el of elements) {
      const val = findXbrlValue(xml, el, 'Instant');
      if (val !== null) {
        // XBRL値はJPY（円）→ 百万円に変換
        result[key] = val / 1000000;
        break;
      }
    }
  }

  // ── 株式数（単位は株、変換不要）──
  const sharesVal = findXbrlValue(xml, 'NumberOfSharesIssued', 'Instant')
    || findXbrlValue(xml, 'TotalNumberOfIssuedShares', 'Instant');
  if (sharesVal !== null) result.sharesIssued = sharesVal;

  const treasurySharesCount = findXbrlValue(xml, 'NumberOfTreasuryShares', 'Instant')
    || findXbrlValue(xml, 'TreasurySharesShares', 'Instant');
  if (treasurySharesCount !== null) result.treasurySharesCount = treasurySharesCount;

  // ── ガバナンス指標 ──
  const foreignPct = findXbrlValue(xml, 'RatioOfForeignShareholding', 'Instant')
    || findXbrlValue(xml, 'ForeignShareholdingRatio', 'Instant')
    || findXbrlValue(xml, 'PercentageOfForeignShareholders', 'Instant');
  if (foreignPct !== null) {
    result.foreignOwnership = foreignPct > 1 ? foreignPct : foreignPct * 100;
  }

  const outsidePct = findXbrlValue(xml, 'RatioOfOutsideDirectors', 'Instant')
    || findXbrlValue(xml, 'OutsideDirectorRatio', 'Instant');
  if (outsidePct !== null) {
    result.outsideDirectorRatio = outsidePct > 1 ? outsidePct : outsidePct * 100;
  }

  return result;
}

/**
 * XBRLから指定要素の値を取得
 * contextRefにtypeHint（"Instant" or "Duration"）を含むものを優先
 * 連結（Consolidated未指定 or NonConsolidated未指定）を優先
 */
function findXbrlValue(xml, elementName, typeHint) {
  // 名前空間プレフィックスが異なる場合があるので柔軟にマッチ
  const regex = new RegExp(
    `<[^>]*?:?${elementName}[^>]*contextRef="([^"]*)"[^>]*>([^<]+)<`,
    'gi'
  );

  let bestVal = null;
  let bestPriority = -1;
  let match;

  while ((match = regex.exec(xml)) !== null) {
    const ctx = match[1];
    const rawVal = match[2].replace(/,/g, '').trim();
    const num = parseFloat(rawVal);
    if (isNaN(num)) continue;

    // 優先度計算
    let priority = 0;

    // typeHintに合致
    if (typeHint && ctx.toLowerCase().includes(typeHint.toLowerCase())) priority += 10;

    // 当期を優先
    if (/CurrentYear/i.test(ctx) || /Current(?!.*Prior)/i.test(ctx)) priority += 5;

    // 連結を優先（NonConsolidatedを含まない）
    if (!/NonConsolidated/i.test(ctx)) priority += 3;

    // 個別メンバーでないものを優先
    if (!/Member/i.test(ctx) || /ConsolidatedMember/i.test(ctx)) priority += 1;

    if (priority > bestPriority) {
      bestPriority = priority;
      bestVal = num;
    }
  }

  return bestVal;
}

// ═══════════════════════════════════════════════════════════════
// 注記テキスト解析（投資不動産・有価証券）
// ═══════════════════════════════════════════════════════════════

function parseNotes(html, data) {
  // 投資不動産の注記を探す
  if (data.investmentPropertyBookValue == null) {
    parseInvestmentPropertyNote(html, data);
  }

  // 有価証券の注記を探す
  if (data.securitiesBookValue == null) {
    parseSecuritiesNote(html, data);
  }

  // 主要な設備の状況から土地明細を探す（含み益推定の補助データ）
  if (!data._landParcelsFound) {
    parseLandFromFacilities(html, data);
  }
}

function parseInvestmentPropertyNote(html, data) {
  // "投資不動産" を含むセクションを探す
  const ipIdx = html.indexOf('投資不動産');
  if (ipIdx === -1) return;

  // 周辺のテーブルを解析
  const searchRange = html.substring(Math.max(0, ipIdx - 500), Math.min(html.length, ipIdx + 5000));

  // 「貸借対照表計上額」「時価」のパターンを探す
  // 一般的なパターン：BS計上額 → 時価 の順で数値が並ぶ
  const bsMatch = searchRange.match(/貸借対照表計上額[\s\S]{0,200}?([0-9,]+)/);
  const fvMatch = searchRange.match(/(?:時価|公正価値)[\s\S]{0,200}?([0-9,]+)/);

  if (bsMatch && fvMatch) {
    const bv = parseFloat(bsMatch[1].replace(/,/g, ''));
    const fv = parseFloat(fvMatch[1].replace(/,/g, ''));
    if (!isNaN(bv) && !isNaN(fv) && bv > 0 && fv > 0) {
      data.investmentPropertyBookValue = bv;
      data.investmentPropertyFairValue = fv;
    }
  }

  // 別パターン：テーブルセルから数値を連続抽出
  if (data.investmentPropertyBookValue == null) {
    const tableMatch = searchRange.match(/<table[\s\S]*?<\/table>/i);
    if (tableMatch) {
      const cells = tableMatch[0].match(/<td[^>]*>[\s\S]*?<\/td>/gi);
      if (cells) {
        const nums = [];
        for (const cell of cells) {
          const text = cell.replace(/<[^>]*>/g, '').replace(/\s+/g, '').replace(/,/g, '');
          const n = parseFloat(text);
          if (!isNaN(n) && n > 0) nums.push(n);
        }
        // 2つ以上の数値があればBS計上額と時価と推定
        if (nums.length >= 2) {
          data.investmentPropertyBookValue = nums[0];
          data.investmentPropertyFairValue = nums[1];
        }
      }
    }
  }
}

function parseSecuritiesNote(html, data) {
  // "その他有価証券" を含むセクションを探す
  const secIdx = html.indexOf('その他有価証券');
  if (secIdx === -1) return;

  const searchRange = html.substring(Math.max(0, secIdx - 200), Math.min(html.length, secIdx + 5000));

  // 「取得原価」「貸借対照表計上額」のパターン
  const costMatch = searchRange.match(/取得原価[\s\S]{0,300}?([0-9,]+)/);
  const bvMatch = searchRange.match(/貸借対照表計上額[\s\S]{0,300}?([0-9,]+)/);

  if (costMatch && bvMatch) {
    const cost = parseFloat(costMatch[1].replace(/,/g, ''));
    const bv = parseFloat(bvMatch[1].replace(/,/g, ''));
    if (!isNaN(cost) && !isNaN(bv) && cost > 0 && bv > 0) {
      data.securitiesBookValue = cost;
      data.securitiesMarketValue = bv;
    }
  }
}

// ═══════════════════════════════════════════════════════════════
// 主要な設備の状況から土地データ抽出
// ═══════════════════════════════════════════════════════════════

function parseLandFromFacilities(html, data) {
  // 「主要な設備の状況」セクションを探す
  const facilityIdx = html.indexOf('主要な設備の状況');
  if (facilityIdx === -1) return;

  const searchRange = html.substring(facilityIdx, Math.min(html.length, facilityIdx + 50000));

  // テーブルから土地面積と所在地を抽出
  const tableRegex = /<table[\s\S]*?<\/table>/gi;
  let tableMatch;
  const parcels = [];

  while ((tableMatch = tableRegex.exec(searchRange)) !== null) {
    const table = tableMatch[0];
    if (!table.includes('土地') || (!table.includes('所在地') && !table.includes('住所'))) continue;

    // 都道府県を含む住所を抽出
    const addrRegex = /((?:北海道|東京都|大阪府|京都府|.{2,3}県)[^\s<,、]{2,40})/g;
    let am;
    while ((am = addrRegex.exec(table)) !== null) {
      const addr = am[1];
      const prefMatch = addr.match(/(北海道|東京都|大阪府|京都府|.{2,3}県)/);
      parcels.push({
        address: addr,
        prefecture: prefMatch ? prefMatch[1] : null,
      });
    }

    if (parcels.length > 0) break;
  }

  if (parcels.length > 0) {
    data._landParcels = parcels;
    data._landParcelsFound = true;
    data.landParcelCount = parcels.length;
    // 都道府県の分布を記録
    const prefCounts = {};
    for (const p of parcels) {
      if (p.prefecture) prefCounts[p.prefecture] = (prefCounts[p.prefecture] || 0) + 1;
    }
    data.landPrefectures = Object.entries(prefCounts)
      .sort((a, b) => b[1] - a[1])
      .map(([pref, count]) => `${pref}(${count}件)`)
      .join(', ');
  }
}

// ═══════════════════════════════════════════════════════════════
// 土地含み益推定
// ═══════════════════════════════════════════════════════════════

function estimateLandGain(data) {
  data.estimatedLandGain = null;
  data.landGainMethod = null;

  const landBV = data.land; // 土地簿価（百万円）
  if (!landBV || landBV <= 0) return;

  // 方法1: 投資不動産の時価/簿価比率を全土地に適用
  if (data.investmentPropertyBookValue > 0 && data.investmentPropertyFairValue > 0) {
    const ipRatio = data.investmentPropertyFairValue / data.investmentPropertyBookValue;
    // 投資不動産の含み益率を保守的に70%で適用（全土地が同じ含み益率とは限らない）
    const conservativeRatio = 1 + (ipRatio - 1) * 0.7;
    const estimatedFV = landBV * conservativeRatio;
    data.estimatedLandGain = Math.round(estimatedFV - landBV);
    data.landGainMethod = `投資不動産比率準用（時価/簿価=${ipRatio.toFixed(2)}倍→保守的70%適用）`;
    return;
  }

  // 方法2: 土地再評価差額金がある場合
  if (data.landRevaluationReserve != null && data.landRevaluationReserve !== 0) {
    // 土地再評価差額金は税効果後の値なので、税率30%で逆算
    const grossGain = data.landRevaluationReserve / 0.7;
    data.estimatedLandGain = Math.round(grossGain);
    data.landGainMethod = `土地再評価差額金から逆算（税効果30%戻し）`;
    return;
  }

  // 方法3: 推定データ不足 → 土地明細分析を案内
  if (data._landParcelsFound && data.landParcelCount > 0) {
    data.landGainMethod = `推定には詳細分析が必要です（${data.landParcelCount}拠点検出: ${data.landPrefectures || ''}）。「土地明細分析」ボタンで地価公示データによる推定が可能です。`;
  } else {
    data.landGainMethod = '推定データ不足（投資不動産注記・再評価差額金なし）。「土地明細分析」ボタンまたは手動入力をご利用ください。';
  }
}

// ═══════════════════════════════════════════════════════════════
// ZIP パーサー（Node.js組込みのみ使用）
// ═══════════════════════════════════════════════════════════════

function readZipEntries(buffer) {
  // End of Central Directoryを末尾から探す
  let eocdOffset = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
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

// ═══════════════════════════════════════════════════════════════
// HTTP ユーティリティ
// ═══════════════════════════════════════════════════════════════

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
      if (resp.statusCode !== 200) {
        resp.resume();
        return reject(new Error(`EDINET returned HTTP ${resp.statusCode}`));
      }
      const chunks = [];
      resp.on('data', chunk => chunks.push(chunk));
      resp.on('end', () => resolve(Buffer.concat(chunks)));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('ダウンロードタイムアウト')); });
  });
}
