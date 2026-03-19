#!/usr/bin/env node
/**
 * edinet-financials.json にinvestmentPropertyBookValue/FairValueを追加
 * 各社のdocIDでXBRL APIを呼び、結果をマージする
 */
const https = require('https');
const fs = require('fs');
const path = require('path');

const DATA_PATH = path.join(__dirname, '..', 'data', 'edinet-financials.json');
const API_BASE = 'https://aoyama-nogizakapublic.vercel.app/api/edinet/xbrl/';

function fetchJSON(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { 'User-Agent': 'batch-updater' } }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`JSON parse error: ${data.substring(0, 200)}`)); }
      });
    }).on('error', reject);
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function main() {
  const json = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
  const companies = json.companies;
  const entries = Object.entries(companies).filter(([, v]) => v._docID);

  console.log(`Total companies with docID: ${entries.length}`);
  let updated = 0, errors = 0, skipped = 0;

  // Process in batches of 3 to avoid rate limits
  for (let i = 0; i < entries.length; i += 3) {
    const batch = entries.slice(i, i + 3);
    const promises = batch.map(async ([code, company]) => {
      try {
        const result = await fetchJSON(`${API_BASE}${company._docID}`);
        if (result.success && result.data) {
          const d = result.data;
          if (d.investmentPropertyBookValue && d.investmentPropertyFairValue) {
            company.investmentPropertyBookValue = d.investmentPropertyBookValue;
            company.investmentPropertyFairValue = d.investmentPropertyFairValue;
            // Recalculate land gain if using investment property ratio
            if (company.land && company.land > 0) {
              const ratio = d.investmentPropertyFairValue / d.investmentPropertyBookValue;
              const conservative = 1 + (ratio - 1) * 0.7;
              if (!company.landRevaluationReserve && !company.landRevaluationExcess) {
                company.estimatedLandGain = Math.round(company.land * conservative - company.land);
                company.landGainMethod = `投資不動産比率準用（時価/簿価=${ratio.toFixed(2)}倍→保守的70%適用）`;
              }
            }
            updated++;
            console.log(`  ✓ ${code} ${company._filerName}: BV=${d.investmentPropertyBookValue}, FV=${d.investmentPropertyFairValue}`);
          } else {
            skipped++;
          }
        } else {
          skipped++;
        }
      } catch (e) {
        errors++;
        console.log(`  ✗ ${code} ${company._filerName}: ${e.message}`);
      }
    });
    await Promise.all(promises);
    if (i + 3 < entries.length) await sleep(1000); // Rate limit
    process.stdout.write(`\r[${Math.min(i + 3, entries.length)}/${entries.length}]`);
  }

  console.log(`\nDone: ${updated} updated, ${skipped} skipped (no IP data), ${errors} errors`);

  json.last_updated = new Date().toISOString().replace('Z', '+09:00');
  fs.writeFileSync(DATA_PATH, JSON.stringify(json, null, 2) + '\n');
  console.log('Saved to', DATA_PATH);
}

main().catch(console.error);
