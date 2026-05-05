#!/usr/bin/env node
// Reads a unified diff on stdin and drops hunks whose "old start line" matches any provided --drop-old-start value.
// Pass-through for all header/file lines; mechanical hunk removal only.
// Usage: node filter-hunks.js --drop-old-start=225 < input.diff > output.diff

const fs = require('fs');

const args = process.argv.slice(2);
const drops = new Set();
for (const a of args) {
  const m = a.match(/^--drop-old-start=(\d+)$/);
  if (m) drops.add(Number(m[1]));
}
if (drops.size === 0) {
  console.error('ERROR: at least one --drop-old-start=<N> required');
  process.exit(2);
}

const input = fs.readFileSync(0, 'utf8');
const lines = input.split(/\r?\n/);

const out = [];
let i = 0;
let droppedHunks = [];
let keptHunks = [];

while (i < lines.length) {
  const line = lines[i];
  const m = line.match(/^@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@/);
  if (!m) {
    out.push(line);
    i++;
    continue;
  }
  const oldStart = Number(m[1]);
  const oldCount = m[2] === undefined ? 1 : Number(m[2]);
  const newCount = m[4] === undefined ? 1 : Number(m[4]);

  // Body line count: each context line counts against both old and new;
  // each '-' counts against old only; each '+' counts against new only.
  // Consume body until both old and new quotas exhausted.
  let oldLeft = oldCount;
  let newLeft = newCount;
  const bodyStart = i + 1;
  let j = bodyStart;
  while (j < lines.length && (oldLeft > 0 || newLeft > 0)) {
    const bl = lines[j];
    if (bl.startsWith('@@') || bl.startsWith('diff --git')) break;
    if (bl.startsWith(' ')) { oldLeft--; newLeft--; }
    else if (bl.startsWith('-')) { oldLeft--; }
    else if (bl.startsWith('+')) { newLeft--; }
    else if (bl === '') {
      // Empty line = treat as context (git sometimes emits this for blank context)
      oldLeft--; newLeft--;
    } else if (bl.startsWith('\\')) {
      // "\ No newline at end of file" — doesn't count
    } else {
      break;
    }
    j++;
  }

  const hunkBodyEnd = j; // exclusive
  if (drops.has(oldStart)) {
    droppedHunks.push({ oldStart, oldCount, newCount, bodyLines: j - bodyStart });
    i = hunkBodyEnd;
    continue;
  } else {
    keptHunks.push({ oldStart, oldCount, newCount, bodyLines: j - bodyStart });
    for (let k = i; k < hunkBodyEnd; k++) out.push(lines[k]);
    i = hunkBodyEnd;
  }
}

process.stdout.write(out.join('\n'));

// Telemetry to stderr (won't pollute stdout patch)
console.error(`kept hunks: ${keptHunks.length} (starts: ${keptHunks.map(h => h.oldStart).join(',')})`);
console.error(`dropped hunks: ${droppedHunks.length} (starts: ${droppedHunks.map(h => h.oldStart).join(',')})`);
