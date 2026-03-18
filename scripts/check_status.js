const fs = require('fs');
const data = JSON.parse(fs.readFileSync('data/reports.json', 'utf8'));

const maki = data.activists['maki_hiroyuki'];
console.log('牧寛之:', maki ? maki.report_count + ' reports' : 'NOT FOUND');
if (maki && maki.holdings) {
  maki.holdings.slice(0, 5).forEach(h => {
    const ratio = h.holding_ratio != null ? h.holding_ratio + '%' : '-';
    console.log('  ', h.target_company || h.sec_code, ratio, h.date);
  });
}

console.log('\n--- Top 33 投資家ランキング ---');
Object.entries(data.activists).forEach(([id, a], i) => {
  console.log((i + 1) + '.', a.name, '-', a.report_count, 'reports', '(' + a.type + ')');
});

console.log('\nTotal reports:', data.reports.length);
console.log('Activist reports:', data.activist_reports);
