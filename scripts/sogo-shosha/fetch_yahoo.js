#!/usr/bin/env node
/**
 * 商社セクター・ダッシュボード フェーズ1 — Yahoo Finance データ取得
 *
 * 7社 (8058 8031 8001 8053 8002 2768 8015) について、
 * yahoo-finance2 の quote / quoteSummary から取れる範囲のファンダメンタル指標を取得し、
 * data/sogo-shosha/raw/<ticker>/yahoo.json に保存する。
 *
 * EDINET / J-Quants の API キーが現状未取得のため、本スクリプトはキー不要で動作する
 * Yahoo Finance を一次ソースとする。詳細は data/sogo-shosha/known-issues.md 参照。
 *
 * 使い方: node scripts/sogo-shosha/fetch_yahoo.js
 */

const fs = require('fs');
const path = require('path');
const YahooFinance = require('yahoo-finance2').default;
const yahooFinance = new YahooFinance({ suppressNotices: ['yahooSurvey'] });

const BASE = path.resolve(__dirname, '..', '..');
const RAW_DIR = path.join(BASE, 'data', 'sogo-shosha', 'raw');

const TARGETS = [
  { ticker: '8058', name: '三菱商事',     tier: 'big5' },
  { ticker: '8031', name: '三井物産',     tier: 'big5' },
  { ticker: '8001', name: '伊藤忠商事',   tier: 'big5' },
  { ticker: '8053', name: '住友商事',     tier: 'big5' },
  { ticker: '8002', name: '丸紅',         tier: 'big5' },
  { ticker: '2768', name: '双日',         tier: 'mid' },
  { ticker: '8015', name: '豊田通商',     tier: 'mid' },
];

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function fetchOne(t) {
  const symbol = `${t.ticker}.T`;
  const out = { ticker: t.ticker, name: t.name, tier: t.tier, symbol, fetchedAt: new Date().toISOString() };

  try {
    const [quote, summary] = await Promise.all([
      yahooFinance.quote(symbol),
      yahooFinance.quoteSummary(symbol, {
        modules: ['financialData', 'defaultKeyStatistics', 'summaryDetail', 'price'],
      }).catch(e => { out._summaryError = e.message; return null; }),
    ]);

    out.quote = {
      regularMarketPrice: quote.regularMarketPrice ?? null,
      marketCap: quote.marketCap ?? null,
      sharesOutstanding: quote.sharesOutstanding ?? null,
      currency: quote.currency ?? null,
      bookValue: quote.bookValue ?? null,
      epsTrailingTwelveMonths: quote.epsTrailingTwelveMonths ?? null,
      epsForward: quote.epsForward ?? null,
      trailingPE: quote.trailingPE ?? null,
      forwardPE: quote.forwardPE ?? null,
      priceToBook: quote.priceToBook ?? null,
    };

    if (summary) {
      out.financialData = summary.financialData || null;
      out.defaultKeyStatistics = summary.defaultKeyStatistics || null;
      out.summaryDetail = summary.summaryDetail || null;
      out.price = summary.price ? {
        regularMarketPrice: summary.price.regularMarketPrice ?? null,
        marketCap: summary.price.marketCap ?? null,
        currency: summary.price.currency ?? null,
        currencySymbol: summary.price.currencySymbol ?? null,
        longName: summary.price.longName ?? null,
        shortName: summary.price.shortName ?? null,
      } : null;
    }
  } catch (e) {
    out._error = e.message;
  }

  return out;
}

async function main() {
  console.log(`=== 商社7社 Yahoo Finance fetch ===`);

  for (const t of TARGETS) {
    process.stdout.write(`  ${t.ticker} ${t.name} ... `);
    const data = await fetchOne(t);
    const dir = path.join(RAW_DIR, t.ticker);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'yahoo.json'), JSON.stringify(data, null, 2), 'utf8');
    if (data._error) {
      console.log(`ERROR: ${data._error}`);
    } else {
      const fd = data.financialData || {};
      const ks = data.defaultKeyStatistics || {};
      console.log(`OK  (PER=${ks.trailingPE?.toFixed(1) ?? 'na'}, PBR=${ks.priceToBook?.toFixed(2) ?? 'na'}, ROE=${fd.returnOnEquity != null ? (fd.returnOnEquity*100).toFixed(1)+'%' : 'na'})`);
    }
    await sleep(500);
  }

  console.log(`\n=== 完了 ===`);
  console.log(`  → data/sogo-shosha/raw/<ticker>/yahoo.json に保存`);
}

main().catch(e => {
  console.error('FATAL:', e);
  process.exit(1);
});
