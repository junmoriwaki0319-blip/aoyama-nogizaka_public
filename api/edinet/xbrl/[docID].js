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
  if (!/^[A-Za-z0-9]{8,12}$/.test(docID)) {
    return res.status(400).json({ success: false, error: 'docIDの形式が不正です' });
  }
  const apiKey = req.query.apiKey || process.env.EDINET_API_KEY;

  if (!docID) {
    return res.status(400).json({ success: false, error: 'docIDが必要です' });
  }
  if (!/^[A-Za-z0-9]{8,12}$/.test(docID)) {
    return res.status(400).json({ success: false, error: 'docIDの形式が不正です' });
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
      /\.(htm(l)?|xhtml)$/i.test(e.name) && /PublicDoc/i.test(e.name)
    );
    for (const ne of noteEntries) {
      try {
        const html = extractEntry(zipBuffer, ne).toString('utf8');
        parseNotes(html, data);
      } catch {}
    }

    // XBRLインラインドキュメント自体からも大株主データを探す
    // （大株主の状況はHTMLノートではなくXBRL本体に含まれることが多い）
    if (!data._majorShareholdersParsed) {
      parseMajorShareholders(xbrlXml, data);
    }

    // 土地含み益推定
    estimateLandGain(data);

    // 内部フラグを除去
    delete data._majorShareholdersParsed;
    delete data._policyHoldingsParsed;
    delete data._landParcelsFound;

    res.json({ success: true, data });
  } catch (err) {
    console.error('EDINET XBRL error:', err.message);
    res.status(500).json({ success: false, error: 'XBRL解析に失敗しました' });
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

  // ── PL項目（Duration = 期間）──
  for (const el of ['OperatingIncome', 'OperatingProfit']) {
    const val = findXbrlValue(xml, el, 'Duration');
    if (val !== null) { result.operatingIncome = val / 1000000; break; }
  }

  // 減価償却費：総額タグを優先、なければ売上原価+販管費の個別を合算
  const daTotal = findXbrlValue(xml, 'DepreciationAndAmortization', 'Duration')
    || findXbrlValue(xml, 'Depreciation', 'Duration');
  if (daTotal !== null) {
    result.depreciationAndAmortization = daTotal / 1000000;
  } else {
    let daSum = 0; let found = false;
    for (const el of ['DepreciationCostOfSales', 'DepreciationSGA']) {
      const v = findXbrlValue(xml, el, 'Duration');
      if (v !== null) { daSum += v; found = true; }
    }
    if (found) result.depreciationAndAmortization = daSum / 1000000;
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

  // 政策保有株式（特定投資株式）の解析
  if (!data._policyHoldingsParsed) {
    parsePolicyHoldings(html, data);
  }

  // 主要な設備の状況から土地明細を探す（含み益推定の補助データ）
  if (!data._landParcelsFound) {
    parseLandFromFacilities(html, data);
  }

  // 大株主の状況を解析
  if (!data._majorShareholdersParsed) {
    parseMajorShareholders(html, data);
  }
}

function parseInvestmentPropertyNote(html, data) {
  // 「賃貸等不動産」「投資不動産」「賃貸不動産」を含むセクションを探す
  // 日本基準では「賃貸等不動産」が正式名称
  const searchTerms = ['賃貸等不動産', '賃貸不動産', '投資不動産'];
  let ipIdx = -1;
  let matchedTerm = '';
  for (const term of searchTerms) {
    ipIdx = html.indexOf(term);
    if (ipIdx !== -1) { matchedTerm = term; break; }
  }
  if (ipIdx === -1) return;

  // 周辺のテーブルを解析（前後を広めに取る）
  const searchRange = html.substring(Math.max(0, ipIdx - 500), Math.min(html.length, ipIdx + 8000));
  const debugInfo = { matchedTerm, ipIdx };

  // 賃貸等不動産の典型的なテーブル構造:
  // 日本基準の注記では以下の2つの表がある:
  //   表1: 期首残高、期中増減額、期末残高（BS計上額ベース）
  //   表2: 貸借対照表計上額（期末）、時価（期末）、差額
  // ここでは表2（期末のBS計上額と時価）を抽出したい

  // HTMLタグを除去したクリーンテキスト（数値位置の特定用）
  const cleanText = searchRange.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/\s+/g, ' ');

  // パターン1: テーブルから期末の「貸借対照表計上額」「時価」「差額」を探す
  // 典型的な構造: ... 貸借対照表計上額 ... 時価 ... 差額 ...
  //              ... 1,234,567 ... 2,345,678 ... 1,111,111 ...
  // 数値は百万円単位が多い（大企業は百万円、中小は千円）

  // 期末の貸借対照表計上額と時価を含む行を探す
  // 「貸借対照表計上額」と「時価」が近くにある部分を見つけ、その下の数値を取得
  const bsFvPattern = cleanText.match(/貸借対照表計上額[\s（(]*(?:百万円|千円|円)*[\s）)]*[\s,]*(?:時価|当期末の時価)[\s（(]*(?:百万円|千円|円)*[\s）)]*[\s,]*(?:差額)/);
  if (bsFvPattern) {
    // ヘッダー行の後の数値行を探す（5桁以上 or カンマ付き数値のみ）
    const afterHeader = cleanText.substring(cleanText.indexOf(bsFvPattern[0]) + bsFvPattern[0].length);
    const numsInRow = afterHeader.match(/([0-9]{1,3},[0-9]{3}(?:,[0-9]{3})*|[0-9]{5,})/g);
    if (numsInRow && numsInRow.length >= 2) {
      const bv = parseFloat(numsInRow[0].replace(/,/g, ''));
      const fv = parseFloat(numsInRow[1].replace(/,/g, ''));
      if (bv > 100 && fv > 100) {
        data.investmentPropertyBookValue = bv;
        data.investmentPropertyFairValue = fv;
        data._ipDebug = { pattern: '1a-header-then-nums', bv, fv };
        return;
      }
    }
  }

  // パターン1b: 「貸借対照表計上額」の後の最初の大きな数値、「時価」の後の最初の大きな数値
  // 年号(2023,2024,2025等)を除外するため、5桁以上 or カンマ付き4桁以上を対象
  const numPattern = '([0-9]{1,3},[0-9]{3}(?:,[0-9]{3})*|[0-9]{5,})';
  const bsMatch = cleanText.match(new RegExp('貸借対照表計上額[\\s\\S]{0,300}?' + numPattern));
  const fvMatch = cleanText.match(new RegExp('(?:当期末の時価|時価|公正価値)[\\s\\S]{0,300}?' + numPattern));
  if (bsMatch && fvMatch) {
    const bv = parseFloat(bsMatch[1].replace(/,/g, ''));
    const fv = parseFloat(fvMatch[1].replace(/,/g, ''));
    if (bv > 100 && fv > 100) {
      data.investmentPropertyBookValue = bv;
      data.investmentPropertyFairValue = fv;
      data._ipDebug = { pattern: '1b-bs-fv-keywords', bv, fv };
      return;
    }
  }

  // パターン2: テーブルの最終行付近から「期末」の数値を取得
  // テーブル構造: 期首 | 増加 | 減少 | 期末(BS) | 期末(時価) | 差額
  if (data.investmentPropertyBookValue == null) {
    // テーブル内のすべての行を解析
    const tables = searchRange.match(/<table[\s\S]*?<\/table>/gi);
    if (tables) {
      for (const table of tables) {
        const rows = table.match(/<tr[\s\S]*?<\/tr>/gi);
        if (!rows) continue;

        // ヘッダーに「貸借対照表計上額」「時価」が含まれるテーブルを探す
        const tableText = table.replace(/<[^>]*>/g, ' ');
        if (!tableText.includes('貸借対照表計上額') && !tableText.includes('時価')) continue;

        // 各行のセルを解析し、4桁以上の数値を持つ行を収集
        for (const row of rows) {
          const cells = row.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi);
          if (!cells) continue;
          const cellTexts = cells.map(c => c.replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, ' ').replace(/\s+/g, '').trim());

          // 「期末」「当期末」を含む行を探す
          const isEndRow = cellTexts.some(t => /期末|当期末/.test(t));
          if (!isEndRow) continue;

          // この行からカンマ付き数値 or 5桁以上の数値を収集（年号除外）
          const nums = [];
          for (const t of cellTexts) {
            // カンマ付き数値（例: 1,234,567）か5桁以上を対象
            if (/[0-9]{1,3},[0-9]{3}/.test(t) || /[0-9]{5,}/.test(t.replace(/,/g, ''))) {
              const cleaned = t.replace(/,/g, '').replace(/[△\-]/g, '');
              const n = parseFloat(cleaned);
              if (!isNaN(n) && n >= 100) nums.push(n);
            }
          }

          if (nums.length >= 2) {
            // 最後の2つまたは3つの数値が「BS計上額」「時価」「差額」
            // 差額 = 時価 - BS計上額 なので検証可能
            if (nums.length >= 3) {
              const last3 = nums.slice(-3);
              // last3[2] ≈ last3[1] - last3[0] なら正しい
              if (Math.abs(last3[2] - (last3[1] - last3[0])) < last3[0] * 0.01) {
                data.investmentPropertyBookValue = last3[0];
                data.investmentPropertyFairValue = last3[1];
                data._ipDebug = { pattern: '2-table-end-row-3nums', nums: last3 };
                return;
              }
            }
            // 2つだけの場合は最初がBS、次が時価
            data.investmentPropertyBookValue = nums[nums.length - 2];
            data.investmentPropertyFairValue = nums[nums.length - 1];
            data._ipDebug = { pattern: '2-table-end-row-2nums', nums };
            return;
          }
        }
      }
    }
  }

  // パターン3: 「期末」キーワード付近から大きな数値ペアを取得（フォールバック）
  if (data.investmentPropertyBookValue == null) {
    const bigNumPat = '([0-9]{1,3},[0-9]{3}(?:,[0-9]{3})*|[0-9]{5,})';
    const endBalMatch = cleanText.match(new RegExp('期末[\\s\\S]{0,300}?' + bigNumPat + '[\\s\\S]{0,200}?' + bigNumPat));
    if (endBalMatch) {
      const v1 = parseFloat(endBalMatch[1].replace(/,/g, ''));
      const v2 = parseFloat(endBalMatch[2].replace(/,/g, ''));
      if (v1 > 100 && v2 > 100) {
        data.investmentPropertyBookValue = v1;
        data.investmentPropertyFairValue = v2;
        data._ipDebug = { pattern: '3-period-end-fallback', v1, v2 };
      }
    }
  }

  if (!data._ipDebug) {
    // デバッグ: マッチしなかった場合、周辺テキストを返す
    data._ipDebug = { pattern: 'no-match', searchSnippet: cleanText.substring(0, 800) };
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
// 政策保有株式（特定投資株式）の解析
// ═══════════════════════════════════════════════════════════════

function parsePolicyHoldings(html, data) {
  // 「特定投資株式」セクションを探す
  const idx = html.indexOf('特定投資株式');
  if (idx === -1) return;

  // 特定投資株式の後の最初のテーブルを取得（大企業は120KB超のテーブルがある）
  const afterSection = html.substring(idx, Math.min(html.length, idx + 300000));
  const tableMatch = afterSection.match(/<table[\s\S]*?<\/table>/i);
  if (!tableMatch) return;

  const table = tableMatch[0];

  // 全行を取得
  const rowRegex = /<tr[\s\S]*?<\/tr>/gi;
  const rows = [];
  let rm;
  while ((rm = rowRegex.exec(table)) !== null) {
    rows.push(rm[0]);
  }
  if (rows.length < 3) return;

  // セルを取得するヘルパー
  function getCells(rowHtml) {
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    const cells = [];
    let cm;
    while ((cm = cellRegex.exec(rowHtml)) !== null) {
      cells.push(cm[1].replace(/<[^>]*>/g, '').replace(/[\s\u00a0]+/g, '').replace(/,/g, ''));
    }
    return cells;
  }

  // ヘッダー行を解析して構造・単位を把握
  // 単位: 百万円 or 千円（会社規模により異なる）
  // テーブル全体のテキストから単位を検出
  let unitDivisor = 1; // 百万円単位ならそのまま、千円単位なら÷1000
  const tableText = table.replace(/<[^>]*>/g, '');
  if (/計上額[^百千]*千円/.test(tableText) || /計上額[（(][^)）]*千円/.test(tableText)) {
    unitDivisor = 1000; // 千円→百万円に変換
  }
  // 百万円が明記されている場合はそのまま（デフォルト）

  const holdings = [];
  let totalMarketValue = 0;

  // まずヘッダーを飛ばしてデータ行を探す
  let dataStartIdx = 0;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const cells = getCells(rows[i]);
    const rowText = cells.join('');
    if (rowText.includes('貸借対照表計上額') || rowText.includes('株式数') || rowText.includes('銘柄')) {
      dataStartIdx = i + 1;
    }
  }

  // パターン判定: 最初のデータ行に銘柄名らしきテキストがあるか
  if (dataStartIdx >= rows.length) return;

  const firstDataCells = getCells(rows[dataStartIdx]);

  // パターンA: 交互行（銘柄名が含まれ、次の行にBS計上額）
  // トヨタ等の大企業で典型的
  let patternA = false;
  if (dataStartIdx + 1 < rows.length) {
    const nextCells = getCells(rows[dataStartIdx + 1]);
    // 最初のデータ行に漢字/カナが多くて銘柄名っぽく、次の行が数値中心なら交互行パターン
    const hasName = /[ぁ-んァ-ヶー一-龠Ａ-Ｚ㈱㈲]/.test(firstDataCells[0] || '');
    const nextIsNumeric = nextCells.length > 0 && /^[\d△▲\-−]+$/.test(nextCells[0] || '');
    if (hasName && nextIsNumeric) patternA = true;
  }

  if (patternA) {
    // 交互行パターン: 奇数行=銘柄, 偶数行=金額
    for (let i = dataStartIdx; i + 1 < rows.length; i += 2) {
      const nameCells = getCells(rows[i]);
      const valueCells = getCells(rows[i + 1]);
      if (nameCells.length === 0) continue;

      const name = nameCells[0].replace(/[\s\u3000]/g, '');
      if (!name || /^[\d,.\-]+$/.test(name)) break; // 数値のみなら銘柄名ではない

      // 株式数（当事業年度）は nameCells[1]
      const shares = parseFloat((nameCells[1] || '').replace(/[△▲\-−]/g, '-'));
      // BS計上額（当事業年度）は valueCells[0]、単位を百万円に統一
      const bsRaw = parseFloat((valueCells[0] || '').replace(/[△▲\-−]/g, '-'));
      const bsValue = !isNaN(bsRaw) ? bsRaw / unitDivisor : NaN;
      if (!isNaN(bsValue) && bsValue > 0) {
        holdings.push({ name, shares: isNaN(shares) ? null : shares, marketValue: Math.round(bsValue * 100) / 100 });
        totalMarketValue += bsValue;
      }
    }
  } else {
    // パターンB: 1行に全情報（銘柄・株式数・BS計上額が同じ行）
    for (let i = dataStartIdx; i < rows.length; i++) {
      const cells = getCells(rows[i]);
      if (cells.length < 2) continue;

      const name = cells[0].replace(/[\s\u3000]/g, '');
      if (!name || /^[\d,.\-]+$/.test(name)) continue;
      if (name.includes('合計') || name.includes('計')) continue;

      // 数値セルからBS計上額を探す（百万円単位の数値）
      for (let j = 1; j < cells.length; j++) {
        const val = parseFloat(cells[j]);
        if (!isNaN(val) && val > 0) {
          // 株式数（大きな値）ではなくBS計上額（相対的に小さい値）を使う
          // 株式数は通常万単位以上、BS計上額は百万円単位
          // 2番目に見つかった正の数値をBS計上額とみなす（最初は株式数の可能性）
          const nums = [];
          for (let k = 1; k < cells.length; k++) {
            const n = parseFloat(cells[k]);
            if (!isNaN(n) && n > 0) nums.push(n);
          }
          // 当事業年度のBS計上額（通常は株式数の次）、単位を百万円に統一
          if (nums.length >= 2) {
            const mv = nums[1] / unitDivisor;
            holdings.push({ name, shares: nums[0], marketValue: Math.round(mv * 100) / 100 });
            totalMarketValue += mv;
          }
          break;
        }
      }
    }
  }

  if (holdings.length > 0) {
    data._policyHoldingsParsed = true;
    data.policyHoldingsCount = holdings.length;
    data.policyHoldingsMarketValue = Math.round(totalMarketValue * 100) / 100;
    // 上位20銘柄を記録（株式数含む）
    data.policyHoldingsTop = holdings
      .sort((a, b) => b.marketValue - a.marketValue)
      .slice(0, 20)
      .map(h => ({ name: h.name, shares: h.shares, marketValue: h.marketValue }));

    // securitiesMarketValue が未設定の場合、政策保有株式の合計を設定
    if (data.securitiesMarketValue == null) {
      data.securitiesMarketValue = totalMarketValue;
      // securitiesBookValue は投資有価証券BS計上額を使う（近似）
      if (data.securitiesBookValue == null && data.investmentSecurities != null) {
        data.securitiesBookValue = data.investmentSecurities;
      }
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
// 大株主の状況
// ═══════════════════════════════════════════════════════════════

function parseMajorShareholders(html, data) {
  // 「大株主の状況」セクションを探す
  const idx = html.indexOf('大株主の状況');
  if (idx === -1) return;

  const afterSection = html.substring(idx, Math.min(html.length, idx + 100000));
  const tableMatch = afterSection.match(/<table[\s\S]*?<\/table>/i);
  if (!tableMatch) return;

  const table = tableMatch[0];

  // 全行を取得
  const rowRegex = /<tr[\s\S]*?<\/tr>/gi;
  const rows = [];
  let rm;
  while ((rm = rowRegex.exec(table)) !== null) {
    rows.push(rm[0]);
  }
  if (rows.length < 3) return;

  function getCells(rowHtml) {
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    const cells = [];
    let cm;
    while ((cm = cellRegex.exec(rowHtml)) !== null) {
      cells.push(cm[1].replace(/<[^>]*>/g, '').replace(/[\s\u00a0\u3000]+/g, ' ').replace(/,/g, '').trim());
    }
    return cells;
  }

  // ヘッダーを飛ばしてデータ行を探す
  let dataStartIdx = 0;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const cells = getCells(rows[i]);
    const rowText = cells.join('');
    if (rowText.includes('所有株式数') || rowText.includes('持株比率') || rowText.includes('発行済株式')) {
      dataStartIdx = i + 1;
    }
  }

  // 単位判定（千株 or 株）
  const tableText = table.replace(/<[^>]*>/g, '');
  let shareUnit = 1; // 株
  if (/千株/.test(tableText)) shareUnit = 1000;
  if (/百万株/.test(tableText)) shareUnit = 1000000;

  const shareholders = [];
  let totalRatio = 0;

  for (let i = dataStartIdx; i < rows.length; i++) {
    const cells = getCells(rows[i]);
    if (cells.length < 2) continue;

    const name = cells[0].trim();
    if (!name || name === '計' || name.includes('合計')) break;
    // 「自己株式」行もスキップしない（表示用に残す）

    // 数値を探す（所有株式数、持株比率）
    const nums = [];
    for (let j = 1; j < cells.length; j++) {
      const cleaned = cells[j].replace(/[%％]/g, '');
      const val = parseFloat(cleaned);
      if (!isNaN(val)) nums.push(val);
    }

    if (nums.length >= 2) {
      // 一般的に: [所有株式数, 持株比率(%)]
      const shares = nums[0] * shareUnit;
      const ratio = nums[nums.length - 1]; // 最後の数値が持株比率（%）

      // 持株比率は通常0〜100の範囲
      if (ratio > 0 && ratio <= 100) {
        shareholders.push({ name, shares: Math.round(shares), ratio: Math.round(ratio * 100) / 100 });
        totalRatio += ratio;
      }
    } else if (nums.length === 1) {
      // 持株比率のみの場合
      const ratio = nums[0];
      if (ratio > 0 && ratio <= 100) {
        shareholders.push({ name, shares: null, ratio: Math.round(ratio * 100) / 100 });
        totalRatio += ratio;
      }
    }
  }

  if (shareholders.length > 0) {
    data._majorShareholdersParsed = true;
    data.majorShareholders = shareholders;
    data.majorShareholdersCount = shareholders.length;
    data.majorShareholdersTotalRatio = Math.round(totalRatio * 100) / 100;

    // 株主分類（簡易カテゴリ分け）
    const categories = { trust: 0, foreign: 0, insurance: 0, bank: 0, fund: 0, treasury: 0, other: 0 };
    for (const sh of shareholders) {
      const n = sh.name;
      if (/自己株式|自社株/.test(n)) { categories.treasury += sh.ratio; }
      else if (/信託|トラスト|マスタートラスト|日本カストディ|CUSTODY|TRUST|資産管理/.test(n)) { categories.trust += sh.ratio; }
      else if (/生命保険|損害保険|保険/.test(n)) { categories.insurance += sh.ratio; }
      else if (/銀行|バンク|BANK/.test(n)) { categories.bank += sh.ratio; }
      else if (/ファンド|FUND|キャピタル|CAPITAL|インベストメント|INVESTMENT|パートナーズ|PARTNERS|アセット|ASSET|LLC|LLP|L\.P\.|MANAGEMENT/i.test(n)) { categories.fund += sh.ratio; }
      else if (/[A-Z].*[A-Z]/.test(n) || /CORPORATION|COMPANY|LIMITED|INC|NOMINEE|CLEARING|BANK OF|CITIBANK|GOLDMAN|MORGAN|CHASE/i.test(n)) { categories.foreign += sh.ratio; }
      else { categories.other += sh.ratio; }
    }
    data.shareholderCategories = {};
    for (const [k, v] of Object.entries(categories)) {
      if (v > 0) data.shareholderCategories[k] = Math.round(v * 100) / 100;
    }
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
