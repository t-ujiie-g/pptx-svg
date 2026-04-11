/**
 * Shared helpers for node:test category files under test_fixtures/tests/.
 *
 * Each *.test.mjs file imports these helpers, calls `resetAssertions()` at the
 * start of its single top-level `test(...)` block, records assertions via
 * `expect(label, cond)`, and at the end calls `finishAssertions()` which
 * throws an Error summarizing any failures (so node:test marks the test fail).
 *
 * ZIP extraction is delegated to the library's own lib/zip.ts (compiled to
 * dist/zip.js). That removes the duplicated minimal-ZIP parser the old
 * monolithic test_node.mjs carried.
 */

import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { extractZip } from '../../dist/zip.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
export const FIXTURES_DIR = dirname(__dirname);
export const DIST_DIR = join(FIXTURES_DIR, '..', 'dist');

// ── Assertion state ─────────────────────────────────────────────────────────

let failures = [];
let total = 0;

export function resetAssertions() {
  failures = [];
  total = 0;
}

export function expect(label, condition, detail = '') {
  total++;
  if (!condition) {
    failures.push(detail ? `${label} — ${detail}` : label);
  }
}

// Back-compat alias used by mechanically moved sections.
export const assert = expect;

// No-op label marker (kept so moved section() calls still compile).
export function section(_name) {}

export function finishAssertions() {
  console.log(`    ${total - failures.length}/${total} assertions passed`);
  if (failures.length > 0) {
    const lines = failures.slice(0, 50).map((m) => `  ✗ ${m}`).join('\n');
    const more = failures.length > 50 ? `\n  ... (${failures.length - 50} more)` : '';
    throw new Error(`${failures.length}/${total} assertion(s) failed:\n${lines}${more}`);
  }
}

// ── PPTX loading ────────────────────────────────────────────────────────────

function loadPptxAb(filename) {
  const path = join(FIXTURES_DIR, filename);
  if (!existsSync(path)) {
    throw new Error(`fixture not found: ${path}`);
  }
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

export function pptxExists(filename) {
  return existsSync(join(FIXTURES_DIR, filename));
}

export async function loadPptxEntries(filename) {
  return extractZip(loadPptxAb(filename));
}

let featuresCache;
export async function loadFeatures() {
  if (!featuresCache) {
    featuresCache = await loadPptxEntries('test_features.pptx');
  }
  return featuresCache;
}

// ── XML helpers ─────────────────────────────────────────────────────────────

export function countSlideIds(xml) {
  const patterns = ['<p:sldId ', '<p:sldId\t', '<p:sldId\n', '<p:sldId/>'];
  let n = 0;
  for (const pat of patterns) {
    let pos = 0;
    while (true) {
      const idx = xml.indexOf(pat, pos);
      if (idx === -1) break;
      n++;
      pos = idx + pat.length;
    }
  }
  return n;
}

export function hasTag(xml, tag) {
  if (!xml) return false;
  return xml.includes(`<${tag}`) || xml.includes(`<${tag}/`);
}

export function findRelTarget(relsXml, typeSuffix) {
  if (!relsXml) return [];
  const targets = [];
  const re = /<Relationship[^>]*Type="[^"]*\/([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/g;
  let m;
  while ((m = re.exec(relsXml)) !== null) {
    if (m[1] === typeSuffix || m[0].includes(typeSuffix)) {
      targets.push(m[2]);
    }
  }
  const re2 = /<Relationship[^>]*Target="([^"]*)"[^>]*Type="[^"]*\/([^"]*)"[^>]*\/?>/g;
  while ((m = re2.exec(relsXml)) !== null) {
    if (m[2] === typeSuffix || m[0].includes(typeSuffix)) {
      if (!targets.includes(m[1])) targets.push(m[1]);
    }
  }
  return targets;
}
