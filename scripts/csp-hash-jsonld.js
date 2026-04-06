#!/usr/bin/env node
/**
 * csp-hash-jsonld.js
 * 全HTMLファイルからJSON-LDスクリプトを抽出しSHA-256ハッシュを計算する
 */
const fs = require('fs');
const crypto = require('crypto');
const path = require('path');

function findHtml(dir, results) {
  for (const f of fs.readdirSync(dir)) {
    if (f === 'node_modules' || f === 'functions' || f === '.git') continue;
    const full = path.join(dir, f);
    if (fs.statSync(full).isDirectory()) findHtml(full, results);
    else if (f.endsWith('.html')) results.push(full);
  }
  return results;
}

const root = path.resolve(__dirname, '..');
const files = findHtml(root, []);
const hashes = new Set();
const details = [];

for (const file of files) {
  const html = fs.readFileSync(file, 'utf8');
  const re = /<script\s+type="application\/ld\+json"\s*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const content = m[1];
    const hash = crypto.createHash('sha256').update(content, 'utf8').digest('base64');
    hashes.add(hash);
    const relPath = path.relative(root, file).replace(/\\/g, '/');
    details.push({ file: relPath, hash, len: content.length });
  }
}

console.log('=== ' + details.length + ' JSON-LD blocks, ' + hashes.size + ' unique hashes ===\n');
for (const d of details) {
  console.log("'sha256-" + d.hash + "'  " + d.file + ' (' + d.len + ' chars)');
}
console.log('\n=== Deduplicated for CSP script-src ===');
for (const h of hashes) {
  console.log("'sha256-" + h + "'");
}
