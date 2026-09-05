#!/usr/bin/env node
// Read-only git history extractor. Never runs a write git command.
// Usage: node extract.js <path-to-repo> [output.json]

const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const repoPath = process.argv[2];
const outPath = process.argv[3] || path.join(__dirname, 'history.json');

if (!repoPath) {
  console.error('Usage: node extract.js <path-to-repo> [output.json]');
  process.exit(1);
}

const SEP = '\x1f'; // unit separator, unlikely in commit fields
const format = ['%H', '%P', '%D', '%at', '%an', '%s'].join(SEP);

const raw = execFileSync(
  'git',
  ['-C', repoPath, 'log', '--all', `--format=${format}`],
  { encoding: 'utf8', maxBuffer: 1024 * 1024 * 64 }
);

const commits = raw
  .split('\n')
  .filter(Boolean)
  .map((line) => {
    const [hash, parentsRaw, refsRaw, tsRaw, author, subject] = line.split(SEP);
    return {
      hash,
      parents: parentsRaw ? parentsRaw.split(' ').filter(Boolean) : [],
      refs: refsRaw ? refsRaw.split(', ').filter(Boolean) : [],
      timestamp: Number(tsRaw),
      author,
      message: subject,
    };
  });

fs.writeFileSync(outPath, JSON.stringify(commits, null, 2));
console.log(`Wrote ${commits.length} commits to ${outPath}`);
