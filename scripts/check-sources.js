#!/usr/bin/env node
'use strict';

/**
 * Watches the primary tariff source pages for changes.
 *
 * For each entry in data/source-watch.json: fetch the URL, SHA-256 the response
 * body, and compare against the stored hash. When a page has changed since the
 * previous check, the matching entries in data/tariff-data.json are flagged for
 * manual review. The script never edits the tariff figures themselves -- the
 * rate, status, effective and source fields are read-only here, by design.
 *
 * No dependencies beyond Node's built-in fetch and crypto (Node 18+).
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.resolve(__dirname, '..');
const WATCH_FILE = path.join(ROOT, 'data', 'source-watch.json');
const TARIFF_FILE = path.join(ROOT, 'data', 'tariff-data.json');

// Fields the script must never write to, whatever a source page says.
const PROTECTED_FIELDS = ['rate', 'status', 'effective', 'source'];

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function writeJson(file, value) {
  fs.writeFileSync(file, JSON.stringify(value, null, 2) + '\n');
}

function isoDate(value) {
  const d = value ? new Date(value) : new Date();
  return (isNaN(d.getTime()) ? new Date() : d).toISOString().slice(0, 10);
}

function sha256(text) {
  return crypto.createHash('sha256').update(text).digest('hex');
}

/** Collects every object that looks like a dashboard entry (i.e. has a url). */
function collectEntries(node, out) {
  if (Array.isArray(node)) {
    node.forEach((child) => collectEntries(child, out));
  } else if (node && typeof node === 'object') {
    if (typeof node.url === 'string') out.push(node);
    Object.values(node).forEach((child) => collectEntries(child, out));
  }
  return out;
}

function urlsMatch(a, b) {
  const x = a.trim();
  const y = b.trim();
  return x.includes(y) || y.includes(x);
}

async function main() {
  const watch = readJson(WATCH_FILE);
  const sources = Array.isArray(watch) ? watch : watch.sources;
  if (!Array.isArray(sources)) {
    throw new Error('source-watch.json has no "sources" array');
  }

  let tariff = null;
  let tariffEntries = [];
  if (fs.existsSync(TARIFF_FILE)) {
    tariff = readJson(TARIFF_FILE);
    tariffEntries = collectEntries(tariff, []);
  } else {
    console.log(`note: ${path.relative(ROOT, TARIFF_FILE)} not found - review flags will be skipped this run`);
  }

  const now = new Date().toISOString();
  let watchChanged = false;
  let tariffChanged = false;

  for (const source of sources) {
    let body;
    try {
      const res = await fetch(source.url, {
        headers: { 'user-agent': 'report-tools-source-watch/1.0' },
        redirect: 'follow',
      });
      if (!res.ok) {
        console.log(`WARN  ${source.key}: HTTP ${res.status} - skipped, hash left unchanged`);
        continue;
      }
      body = await res.text();
    } catch (err) {
      console.log(`WARN  ${source.key}: fetch failed (${err.message}) - skipped, hash left unchanged`);
      continue;
    }

    const hash = sha256(body);
    const previousHash = source.lastHash;
    const previousChecked = source.lastChecked;

    if (previousHash === null || previousHash === undefined) {
      console.log(`INIT  ${source.key}: first check, baseline hash recorded`);
    } else if (previousHash === hash) {
      console.log(`OK    ${source.key}: unchanged`);
    } else {
      const since = isoDate(previousChecked);
      console.log(`CHANGE ${source.key}: page changed since last check on ${since}`);

      const matches = tariffEntries.filter((entry) => urlsMatch(entry.url, source.url));
      if (matches.length === 0) {
        console.log(`       no tariff-data.json entry matches ${source.url} - nothing flagged`);
      }
      for (const entry of matches) {
        const before = JSON.stringify(PROTECTED_FIELDS.map((f) => entry[f]));
        entry.reviewNeeded = true;
        entry.reviewNote = `source page changed since last check on ${since}, verify manually`;
        // Guard rail: the protected fields must be byte-identical afterwards.
        if (JSON.stringify(PROTECTED_FIELDS.map((f) => entry[f])) !== before) {
          throw new Error(`refusing to write: protected fields mutated on "${entry.name}"`);
        }
        tariffChanged = true;
        console.log(`       flagged: ${entry.name || entry.url}`);
      }
    }

    // The hash and timestamp advance on every successful fetch, changed or not.
    source.lastHash = hash;
    source.lastChecked = now;
    watchChanged = true;
  }

  if (tariffChanged) writeJson(TARIFF_FILE, tariff);
  if (watchChanged) writeJson(WATCH_FILE, watch);

  console.log(
    `done: source-watch.json ${watchChanged ? 'updated' : 'unchanged'}, ` +
      `tariff-data.json ${tariffChanged ? 'flagged for review' : 'unchanged'}`
  );
}

main().catch((err) => {
  console.error(`check-sources failed: ${err.message}`);
  process.exit(1);
});
