/**
 * Node.js test script for pptx-render
 *
 * Tests the JavaScript host layer (ZIP extraction) and validates PPTX
 * structure (XML content, relationships, master/layout chain) independently
 * of the Wasm module. Wasm-GC requires a browser to run.
 *
 * Run: node test_fixtures/test_node.mjs
 */

import { readFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── Minimal ZIP extractor (mirrors host.js logic) ─────────────────────────────

async function extractZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const textFiles = new Map();
  const binaryFiles = new Map();
  const decoder = new TextDecoder('utf-8');

  async function inflate(compressed) {
    const { DecompressionStream } = await import('node:stream/web');
    const stream = new DecompressionStream('deflate-raw');
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();
    writer.write(compressed);
    writer.close();
    const chunks = [];
    let totalLen = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      totalLen += value.length;
    }
    const result = new Uint8Array(totalLen);
    let pos = 0;
    for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
    return result;
  }

  let offset = 0;
  while (offset < bytes.length - 4) {
    const sig = view.getUint32(offset, true);
    if (sig !== 0x04034b50) break;

    const method          = view.getUint16(offset + 8, true);
    const compressedSize  = view.getUint32(offset + 18, true);
    const fileNameLen     = view.getUint16(offset + 26, true);
    const extraLen        = view.getUint16(offset + 28, true);

    const name = decoder.decode(bytes.slice(offset + 30, offset + 30 + fileNameLen));
    const dataOffset = offset + 30 + fileNameLen + extraLen;
    const compressedData = bytes.slice(dataOffset, dataOffset + compressedSize);

    let decompressed;
    if (method === 0) {
      decompressed = compressedData;
    } else if (method === 8) {
      decompressed = await inflate(compressedData);
    } else {
      offset = dataOffset + compressedSize;
      continue;
    }

    const lower = name.toLowerCase();
    const isText = lower.endsWith('.xml') || lower.endsWith('.rels');
    if (isText) {
      textFiles.set(name, decoder.decode(decompressed));
    }
    binaryFiles.set(name, decompressed);

    offset = dataOffset + compressedSize;
  }
  return { textFiles, binaryFiles };
}

// ── Slide count logic (mirrors main.mbt) ─────────────────────────────────────

function countSlideIds(xml) {
  const patterns = ['<p:sldId ', '<p:sldId\t', '<p:sldId\n', '<p:sldId/>'];
  let total = 0;
  for (const pat of patterns) {
    let pos = 0;
    while (true) {
      const idx = xml.indexOf(pat, pos);
      if (idx === -1) break;
      total++;
      pos = idx + pat.length;
    }
  }
  return total;
}

// ── XML helpers ──────────────────────────────────────────────────────────────

/** Simple attribute extraction (no full XML parsing needed) */
function getAttrValue(xml, tag, attr) {
  const re = new RegExp(`<${tag}[^>]*\\s${attr}="([^"]*)"`, 's');
  const m = xml.match(re);
  return m ? m[1] : null;
}

/** Check if a tag exists in XML */
function hasTag(xml, tag) {
  return xml.includes(`<${tag}`) || xml.includes(`<${tag}/`);
}

/** Find Target attribute in Relationship elements matching a type suffix */
function findRelTarget(relsXml, typeSuffix) {
  const targets = [];
  const re = /<Relationship[^>]*Type="[^"]*\/([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/g;
  let m;
  while ((m = re.exec(relsXml)) !== null) {
    if (m[1] === typeSuffix || m[0].includes(typeSuffix)) {
      targets.push(m[2]);
    }
  }
  // Also try with Target before Type
  const re2 = /<Relationship[^>]*Target="([^"]*)"[^>]*Type="[^"]*\/([^"]*)"[^>]*\/?>/g;
  while ((m = re2.exec(relsXml)) !== null) {
    if (m[2] === typeSuffix || m[0].includes(typeSuffix)) {
      if (!targets.includes(m[1])) targets.push(m[1]);
    }
  }
  return targets;
}

// ── Test framework ───────────────────────────────────────────────────────────

let totalTests = 0;
let passedTests = 0;
let failedTests = 0;

function assert(label, condition, detail = '') {
  totalTests++;
  if (condition) {
    passedTests++;
    console.log(`  [${totalTests}] ✓ ${label}`);
  } else {
    failedTests++;
    const extra = detail ? ` — ${detail}` : '';
    console.log(`  [${totalTests}] ✗ FAIL: ${label}${extra}`);
  }
}

function section(name) {
  console.log(`\n── ${name} ──`);
}

function loadPptx(filename) {
  const path = join(__dirname, filename);
  if (!existsSync(path)) {
    console.error(`File not found: ${path}`);
    process.exit(1);
  }
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

// ── Tests: minimal.pptx ─────────────────────────────────────────────────────

async function testMinimalPptx() {
  section('minimal.pptx — basic ZIP extraction');
  const { textFiles } = await extractZip(loadPptx('minimal.pptx'));

  const prsXml = textFiles.get('ppt/presentation.xml');
  assert('presentation.xml exists', !!prsXml);

  const slideCount = countSlideIds(prsXml ?? '');
  assert('slide count = 2', slideCount === 2, `got ${slideCount}`);

  const slide1 = textFiles.get('ppt/slides/slide1.xml');
  assert('slide1.xml exists', !!slide1);
  assert('slide1.xml contains title text', slide1?.includes('Hello from MoonBit') ?? false);

  const slide2 = textFiles.get('ppt/slides/slide2.xml');
  assert('slide2.xml exists', !!slide2);
}

// ── Tests: test_features.pptx ────────────────────────────────────────────────

async function testFeaturesPptx() {
  const path = join(__dirname, 'test_features.pptx');
  if (!existsSync(path)) {
    section('test_features.pptx — SKIPPED (file not found)');
    return;
  }

  const { textFiles } = await extractZip(loadPptx('test_features.pptx'));

  // ── Basic structure ──
  section('test_features.pptx — basic structure');
  const prsXml = textFiles.get('ppt/presentation.xml');
  assert('presentation.xml exists', !!prsXml);

  const slideCount = countSlideIds(prsXml ?? '');
  assert('slide count = 8', slideCount === 8, `got ${slideCount}`);

  // Verify all 8 slides exist
  for (let i = 1; i <= 8; i++) {
    const path = `ppt/slides/slide${i}.xml`;
    assert(`slide${i}.xml exists`, textFiles.has(path));
  }

  // ── Slide .rels ──
  section('test_features.pptx — slide relationships');
  for (let i = 1; i <= 8; i++) {
    const relsPath = `ppt/slides/_rels/slide${i}.xml.rels`;
    const relsXml = textFiles.get(relsPath);
    assert(`slide${i} .rels exists`, !!relsXml);
    if (relsXml) {
      const layoutTargets = findRelTarget(relsXml, 'slideLayout');
      assert(`slide${i} references a slideLayout`, layoutTargets.length > 0,
        layoutTargets.length > 0 ? layoutTargets[0] : 'none');
    }
  }

  // ── Master/Layout chain ──
  section('test_features.pptx — master/layout chain');
  const prsRels = textFiles.get('ppt/_rels/presentation.xml.rels');
  assert('presentation.xml.rels exists', !!prsRels);

  // Check that theme exists
  if (prsRels) {
    const themeTargets = findRelTarget(prsRels, 'theme');
    assert('presentation references a theme', themeTargets.length > 0);
  }

  // Check that slideMaster1.xml exists and has content
  const masterXml = textFiles.get('ppt/slideMasters/slideMaster1.xml');
  assert('slideMaster1.xml exists', !!masterXml);

  if (masterXml) {
    // Verify master has background
    assert('master has p:bg element', hasTag(masterXml, 'p:bg'));
    assert('master has solidFill in bg', masterXml.includes('1B3A6B'));

    // Verify master has txStyles
    assert('master has p:txStyles', hasTag(masterXml, 'p:txStyles'));
    assert('master has p:titleStyle', hasTag(masterXml, 'p:titleStyle'));
    assert('master has p:bodyStyle', hasTag(masterXml, 'p:bodyStyle'));
    assert('master has p:otherStyle', hasTag(masterXml, 'p:otherStyle'));

    // Verify titleStyle has 44pt
    assert('titleStyle has sz="4400" (44pt)', masterXml.includes('sz="4400"'));
  }

  // Check a layout exists and references the master
  const layoutPath = 'ppt/slideLayouts/slideLayout2.xml';
  const layoutXml = textFiles.get(layoutPath);
  assert('slideLayout2.xml exists (title+content)', !!layoutXml);

  const layoutRels = textFiles.get('ppt/slideLayouts/_rels/slideLayout2.xml.rels');
  if (layoutRels) {
    const masterTargets = findRelTarget(layoutRels, 'slideMaster');
    assert('layout2 references slideMaster', masterTargets.length > 0);
  }

  // ── Slide 6: No explicit background ──
  section('test_features.pptx — Slide 6: background inheritance');
  const slide6 = textFiles.get('ppt/slides/slide6.xml');
  if (slide6) {
    // Slide 6 should NOT have its own p:bg (relying on master inheritance)
    const hasBg = hasTag(slide6, 'p:bg');
    assert('slide6 has NO explicit p:bg (relies on master)', !hasBg);
  }

  // ── Slide 7: Placeholder shapes ──
  section('test_features.pptx — Slide 7: placeholder shapes');
  const slide7 = textFiles.get('ppt/slides/slide7.xml');
  if (slide7) {
    assert('slide7 has p:ph elements', hasTag(slide7, 'p:ph'));
    assert('slide7 has title placeholder', slide7.includes('type="title"'));
    // Content placeholder has idx="1" (no explicit type)
    assert('slide7 has body placeholder (idx="1")', slide7.includes('idx="1"'));
    // Check that the title text is present
    assert('slide7 has title text content', slide7.includes('Placeholder Title'));
    // Check level attribute on body paragraph
    assert('slide7 has lvl="1" paragraph', slide7.includes('lvl="1"'));
  }

  // Verify slide7 uses layout2 (Title and Content)
  const slide7Rels = textFiles.get('ppt/slides/_rels/slide7.xml.rels');
  if (slide7Rels) {
    const layouts = findRelTarget(slide7Rels, 'slideLayout');
    assert('slide7 uses slideLayout2 (title+content)',
      layouts.some(t => t.includes('slideLayout2')),
      layouts.join(', '));
  }

  // ── Slide 8: Explicit background overrides master ──
  section('test_features.pptx — Slide 8: explicit background');
  const slide8 = textFiles.get('ppt/slides/slide8.xml');
  if (slide8) {
    assert('slide8 has explicit p:bg', hasTag(slide8, 'p:bg'));
    assert('slide8 bg is green (#228B22)', slide8.includes('228B22'));
  }

  // ── Slide 1-5: existing features regression ──
  section('test_features.pptx — Slides 1-5: feature regression');
  const slide1 = textFiles.get('ppt/slides/slide1.xml');
  if (slide1) {
    assert('slide1 has bodyPr anchor attribute', slide1.includes('anchor='));
    assert('slide1 has lIns/tIns attributes', slide1.includes('lIns='));
  }

  const slide2 = textFiles.get('ppt/slides/slide2.xml');
  if (slide2) {
    assert('slide2 has spcBef spacing', hasTag(slide2, 'a:spcBef'));
    assert('slide2 has spcAft spacing', hasTag(slide2, 'a:spcAft'));
    assert('slide2 has marL indent', slide2.includes('marL='));
  }

  const slide3 = textFiles.get('ppt/slides/slide3.xml');
  if (slide3) {
    assert('slide3 has buChar bullets', hasTag(slide3, 'a:buChar'));
    assert('slide3 has buAutoNum', hasTag(slide3, 'a:buAutoNum'));
  }

  const slide4 = textFiles.get('ppt/slides/slide4.xml');
  if (slide4) {
    assert('slide4 has underline (u=)', slide4.includes(' u='));
    assert('slide4 has strikethrough', slide4.includes('strike='));
    assert('slide4 has baseline (super/subscript)', slide4.includes('baseline='));
  }

  const slide5 = textFiles.get('ppt/slides/slide5.xml');
  if (slide5) {
    assert('slide5 has explicit dark bg', hasTag(slide5, 'p:bg'));
    assert('slide5 bg is navy (#1B3A6B)', slide5.includes('1B3A6B'));
  }

  // ── Layout placeholder transforms ──
  section('test_features.pptx — layout placeholder transforms');
  if (layoutXml) {
    // Layout should have placeholder shapes with transforms
    assert('layout2 has p:ph elements', hasTag(layoutXml, 'p:ph'));
    assert('layout2 has a:xfrm transforms', hasTag(layoutXml, 'a:xfrm'));
    assert('layout2 has a:off (position)', hasTag(layoutXml, 'a:off'));
    assert('layout2 has a:ext (size)', hasTag(layoutXml, 'a:ext'));
  }

  // ── Theme ──
  section('test_features.pptx — theme');
  const themeXml = textFiles.get('ppt/theme/theme1.xml');
  assert('theme1.xml exists', !!themeXml);
  if (themeXml) {
    assert('theme has a:clrScheme', hasTag(themeXml, 'a:clrScheme'));
    assert('theme has a:fontScheme', hasTag(themeXml, 'a:fontScheme'));
    assert('theme has a:majorFont', hasTag(themeXml, 'a:majorFont'));
    assert('theme has a:minorFont', hasTag(themeXml, 'a:minorFont'));
  }
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  console.log('=== pptx-render Node.js Tests ===');

  await testMinimalPptx();
  await testFeaturesPptx();

  console.log(`\n${'═'.repeat(50)}`);
  console.log(`Total: ${totalTests}  Passed: ${passedTests}  Failed: ${failedTests}`);
  if (failedTests > 0) {
    console.log('*** SOME TESTS FAILED ***');
    process.exit(1);
  } else {
    console.log('All tests passed.');
  }
  console.log(`\nFor full Wasm pipeline testing, open in Chrome/Firefox:`);
  console.log('  http://localhost:8765/web/index.html');
}

main().catch(err => { console.error(err); process.exit(1); });
