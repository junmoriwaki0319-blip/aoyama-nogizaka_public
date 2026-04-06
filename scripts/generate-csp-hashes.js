// generate-csp-hashes.js — Calculate SHA-256 hashes for inline JSON-LD scripts
// Usage: node scripts/generate-csp-hashes.js

const { readFileSync, readdirSync, statSync } = require('fs');
const { createHash } = require('crypto');
const { join, relative } = require('path');

const rootDir = join(__dirname, '..');
const hashes = new Set();
const fileHashes = [];

function walkDir(dir) {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    if (entry === 'node_modules' || entry === '.git' || entry === 'scripts') continue;
    if (statSync(full).isDirectory()) {
      walkDir(full);
    } else if (full.endsWith('.html')) {
      processHtml(full);
    }
  }
}

function processHtml(filePath) {
  const html = readFileSync(filePath, 'utf-8');

  // Extract all <script type="application/ld+json">...</script> content
  const regex = /<script\s+type="application\/ld\+json"\s*>([\s\S]*?)<\/script>/gi;
  let match;

  while ((match = regex.exec(html)) !== null) {
    const content = match[1];
    const hash = createHash('sha256').update(content, 'utf-8').digest('base64');
    const directive = `'sha256-${hash}'`;
    hashes.add(directive);
    fileHashes.push({
      file: relative(rootDir, filePath),
      hash: directive,
      preview: content.trim().substring(0, 80) + '...'
    });
  }
}

walkDir(rootDir);

console.log('=== JSON-LD CSP Hashes ===\n');

for (const item of fileHashes) {
  console.log(`${item.file}`);
  console.log(`  ${item.hash}`);
  console.log(`  ${item.preview}`);
  console.log('');
}

console.log('=== Unique Hashes for script-src ===\n');
const allHashes = [...hashes].join(' ');
console.log(allHashes);
console.log(`\n(${hashes.size} unique hashes)\n`);

console.log('=== Ready-to-use script-src directive ===\n');
console.log(`script-src 'self' ${allHashes};`);
console.log('');
