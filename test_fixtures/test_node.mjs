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
  assert('slide count = 58', slideCount === 58, `got ${slideCount}`);

  // Verify all slides exist
  for (let i = 1; i <= 58; i++) {
    const path = `ppt/slides/slide${i}.xml`;
    assert(`slide${i}.xml exists`, textFiles.has(path));
  }

  // ── Slide .rels ──
  section('test_features.pptx — slide relationships');
  for (let i = 1; i <= 58; i++) {
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
    // EA font in theme
    assert('theme majorFont has a:ea', themeXml.includes('<a:ea') && themeXml.includes('Yu Gothic'));
  }

  // ── Slide 9: East Asian fonts + font theme references ──
  section('test_features.pptx — Slide 9: EA fonts & theme refs');
  const slide9 = textFiles.get('ppt/slides/slide9.xml');
  if (slide9) {
    // Explicit EA font
    assert('slide9 has a:ea element', hasTag(slide9, 'a:ea'));
    assert('slide9 has MS PGothic EA font', slide9.includes('MS PGothic'));
    // Theme font references
    assert('slide9 has +mj-ea reference', slide9.includes('+mj-ea'));
    assert('slide9 has +mj-lt reference', slide9.includes('+mj-lt'));
    assert('slide9 has +mn-ea reference', slide9.includes('+mn-ea'));
    assert('slide9 has +mn-lt reference', slide9.includes('+mn-lt'));
    // Explicit EA font: Meiryo
    assert('slide9 has Meiryo EA font', slide9.includes('Meiryo'));
    // Latin fonts alongside EA
    assert('slide9 has a:latin element', hasTag(slide9, 'a:latin'));
  }

  // ── Slide 10: Line spacing ──
  section('test_features.pptx — Slide 10: line spacing');
  const slide10 = textFiles.get('ppt/slides/slide10.xml');
  if (slide10) {
    assert('slide10 has a:lnSpc', hasTag(slide10, 'a:lnSpc'));
    // Percentage: 150%
    assert('slide10 has spcPct 150000', slide10.includes('150000'));
    // Point-based: 36pt = 3600 hundredths
    assert('slide10 has spcPts 3600', slide10.includes('3600'));
    // Tight: 80%
    assert('slide10 has spcPct 80000', slide10.includes('80000'));
  }

  // ── Slide 12: normAutofit ──
  section('test_features.pptx — Slide 12: normAutofit');
  const slide12 = textFiles.get('ppt/slides/slide12.xml');
  if (slide12) {
    // Shape 1: fontScale=80000 only
    assert('slide12 has a:normAutofit', hasTag(slide12, 'a:normAutofit'));
    assert('slide12 has fontScale="80000"', slide12.includes('fontScale="80000"'));
    // Shape 2: fontScale=62500 + lnSpcReduction=20000
    assert('slide12 has fontScale="62500"', slide12.includes('fontScale="62500"'));
    assert('slide12 has lnSpcReduction="20000"', slide12.includes('lnSpcReduction="20000"'));
    // Shape 3: normAutofit with defaults (no fontScale attr)
    // Verify there is a bare <a:normAutofit/> (no attributes)
    assert('slide12 has bare <a:normAutofit/>', slide12.includes('<a:normAutofit/>'));
    // At least one shape has bodyPr without normAutofit (reference shape)
    // The reference shape has bodyPr with insets but no normAutofit child
    assert('slide12 has multiple normAutofit elements', (slide12.match(/normAutofit/g) || []).length >= 3);
  }

  // ── Slide 11: Character spacing + lstStyle ──
  section('test_features.pptx — Slide 11: char spacing & lstStyle');
  const slide11 = textFiles.get('ppt/slides/slide11.xml');
  if (slide11) {
    // Character spacing: spc attribute on rPr
    assert('slide11 has spc="300" (wide)', slide11.includes('spc="300"'));
    assert('slide11 has spc="1000" (very wide)', slide11.includes('spc="1000"'));
    assert('slide11 has spc="-100" (tight)', slide11.includes('spc="-100"'));
    // lstStyle
    assert('slide11 has a:lstStyle with content', slide11.includes('<a:lstStyle>') && slide11.includes('a:lvl1pPr'));
    assert('slide11 lstStyle has sz="2000" (20pt)', slide11.includes('sz="2000"'));
    assert('slide11 lstStyle has Meiryo EA', slide11.includes('Meiryo'));
    assert('slide11 lstStyle has #003366 color', slide11.includes('003366'));
  }

  // ── Slide 13: Text wrapping ──
  section('test_features.pptx — Slide 13: text wrapping');
  const slide13 = textFiles.get('ppt/slides/slide13.xml');
  if (slide13) {
    // Long text that should wrap
    assert('slide13 has long Latin text', slide13.includes('automatically wrap'));
    // CJK text
    assert('slide13 has CJK text', slide13.includes('折り返しテスト'));
    // Mixed Latin + CJK
    assert('slide13 has mixed text', slide13.includes('Mixed'));
    // wrap="none" attribute
    assert('slide13 has wrap="none"', slide13.includes('wrap="none"'));
    // Multiple textboxes
    assert('slide13 has word_wrap bodyPr', hasTag(slide13, 'a:bodyPr'));
  }

  // ── Slide 14: Bullet formatting ──
  section('test_features.pptx — Slide 14: bullet formatting');
  const slide14 = textFiles.get('ppt/slides/slide14.xml');
  if (slide14) {
    // Bullet font
    assert('slide14 has a:buFont', hasTag(slide14, 'a:buFont'));
    assert('slide14 has Wingdings buFont', slide14.includes('Wingdings'));
    assert('slide14 has Symbol buFont', slide14.includes('Symbol'));
    // Bullet size percentage
    assert('slide14 has a:buSzPct', hasTag(slide14, 'a:buSzPct'));
    assert('slide14 has buSzPct 150000', slide14.includes('150000'));
    assert('slide14 has buSzPct 75000', slide14.includes('75000'));
    // Bullet size points
    assert('slide14 has a:buSzPts', hasTag(slide14, 'a:buSzPts'));
    assert('slide14 has buSzPts 3200', slide14.includes('3200'));
    assert('slide14 has buSzPts 800', slide14.includes('800'));
    // Bullet color
    assert('slide14 has a:buClr', hasTag(slide14, 'a:buClr'));
    assert('slide14 has red bullet (FF0000)', slide14.includes('FF0000'));
    assert('slide14 has green bullet (00AA00)', slide14.includes('00AA00'));
    assert('slide14 has blue bullet (0000FF)', slide14.includes('0000FF'));
  }

  // ── Slide 15: Capitalization ──
  section('test_features.pptx — Slide 15: capitalization');
  const slide15 = textFiles.get('ppt/slides/slide15.xml');
  if (slide15) {
    assert('slide15 has cap="all"', slide15.includes('cap="all"'));
    assert('slide15 has cap="small"', slide15.includes('cap="small"'));
    // Text content should be original (not uppercased in XML)
    assert('slide15 has original text (not uppercased)', slide15.includes('This Should Be All Caps'));
    assert('slide15 has small caps text', slide15.includes('Small Caps Text'));
  }

  // ── Slide 16: Color map override ──
  section('test_features.pptx — Slide 16: clrMapOvr');
  const slide16 = textFiles.get('ppt/slides/slide16.xml');
  if (slide16) {
    assert('slide16 has p:clrMapOvr', hasTag(slide16, 'p:clrMapOvr'));
    assert('slide16 has a:overrideClrMapping', hasTag(slide16, 'a:overrideClrMapping'));
    // bg1="dk1" swap
    assert('slide16 has bg1="dk1"', slide16.includes('bg1="dk1"'));
    // tx1="lt1" swap
    assert('slide16 has tx1="lt1"', slide16.includes('tx1="lt1"'));
    // Has scheme color reference tx1
    assert('slide16 has schemeClr val="tx1"', slide16.includes('val="tx1"'));
    // Has dark background
    assert('slide16 has dark bg color', slide16.includes('1A1A2E'));
  }

  // ── Slide 17: CS/Sym fonts + kerning ──
  section('test_features.pptx — Slide 17: CS/Sym fonts + kern');
  const slide17 = textFiles.get('ppt/slides/slide17.xml');
  if (slide17) {
    assert('slide17 has a:cs element', hasTag(slide17, 'a:cs'));
    assert('slide17 has Arial CS font', slide17.includes('Arial'));
    assert('slide17 has a:sym element', hasTag(slide17, 'a:sym'));
    assert('slide17 has Wingdings sym font', slide17.includes('Wingdings'));
    assert('slide17 has kern attribute', slide17.includes('kern="1200"'));
    assert('slide17 has CS text content', slide17.includes('Complex Script'));
    assert('slide17 has Symbol text content', slide17.includes('Symbol Font'));
  }

  // ── Slide 18: Text rotation + tab stops ──
  section('test_features.pptx — Slide 18: rotation + tabs');
  const slide18 = textFiles.get('ppt/slides/slide18.xml');
  if (slide18) {
    assert('slide18 has bodyPr rot', slide18.includes('rot="2700000"'));
    assert('slide18 has a:tabLst', hasTag(slide18, 'a:tabLst'));
    assert('slide18 has a:tab element', hasTag(slide18, 'a:tab'));
    assert('slide18 has tab pos', slide18.includes('pos="2743200"'));
    assert('slide18 has tab algn', slide18.includes('algn="r"'));
    assert('slide18 has rotated text', slide18.includes('Rotated text'));
    assert('slide18 has tab content', slide18.includes('Col1'));
  }

  // ── Slide 19: Vertical text + columns ──
  section('test_features.pptx — Slide 19: vert text + columns');
  const slide19 = textFiles.get('ppt/slides/slide19.xml');
  if (slide19) {
    assert('slide19 has vert="vert"', slide19.includes('vert="vert"'));
    assert('slide19 has vert="eaVert"', slide19.includes('vert="eaVert"'));
    assert('slide19 has numCol="2"', slide19.includes('numCol="2"'));
    assert('slide19 has spcCol', slide19.includes('spcCol="457200"'));
    assert('slide19 has vertical text content', slide19.includes('Vertical text'));
  }

  // ── Slide 20: Hyperlink + RTL ──
  section('test_features.pptx — Slide 20: hyperlink + RTL');
  const slide20 = textFiles.get('ppt/slides/slide20.xml');
  if (slide20) {
    assert('slide20 has a:hlinkClick', hasTag(slide20, 'a:hlinkClick'));
    assert('slide20 has r:id on hlinkClick', slide20.includes('r:id='));
    assert('slide20 has hyperlink text', slide20.includes('Click here'));
    assert('slide20 has rtl="1"', slide20.includes('rtl="1"'));
    assert('slide20 has RTL text', slide20.includes('RTL paragraph'));
  }
  // Check slide20 rels for hyperlink target
  const slide20Rels = textFiles.get('ppt/slides/_rels/slide20.xml.rels');
  if (slide20Rels) {
    assert('slide20 rels has hyperlink', slide20Rels.includes('hyperlink'));
    assert('slide20 rels has example.com', slide20Rels.includes('example.com'));
  }

  // ── Slide 21: Image bullet ──
  section('test_features.pptx — Slide 21: image bullet');
  const slide21 = textFiles.get('ppt/slides/slide21.xml');
  if (slide21) {
    assert('slide21 has a:buBlip', hasTag(slide21, 'a:buBlip'));
    assert('slide21 has a:blip in buBlip', hasTag(slide21, 'a:blip'));
    assert('slide21 has r:embed on blip', slide21.includes('r:embed='));
    assert('slide21 has bullet text', slide21.includes('Image bullet'));
  }

  // ── Slide 22: Hover link + link color ──
  {
    section('test_features.pptx — slide 22 (hover link)');
    const slide22 = textFiles.get('ppt/slides/slide22.xml');
    assert('slide22 exists', !!slide22);
    assert('slide22 has a:hlinkClick', hasTag(slide22, 'a:hlinkClick'));
    assert('slide22 has a:hlinkHover', hasTag(slide22, 'a:hlinkHover'));
    assert('slide22 has Click link text', slide22.includes('Click link'));
    assert('slide22 has Hover link text', slide22.includes('Hover link'));
    assert('slide22 has Both links text', slide22.includes('Both links'));
  }

  // ── Slide 23: Linear gradient fills ──
  {
    section('test_features.pptx — Slide 23: linear gradient fills');
    const slide23 = textFiles.get('ppt/slides/slide23.xml');
    assert('slide23 exists', !!slide23);
    assert('slide23 has a:gradFill', hasTag(slide23, 'a:gradFill'));
    assert('slide23 has a:gsLst', hasTag(slide23, 'a:gsLst'));
    assert('slide23 has a:gs stops', hasTag(slide23, 'a:gs'));
    assert('slide23 has a:lin', hasTag(slide23, 'a:lin'));
    // 3-stop gradient: red→yellow→blue
    assert('slide23 has FF0000 (red)', slide23.includes('FF0000'));
    assert('slide23 has FFFF00 (yellow)', slide23.includes('FFFF00'));
    assert('slide23 has 0000FF (blue)', slide23.includes('0000FF'));
    // Various angles
    assert('slide23 has ang="0" (0°)', slide23.includes('ang="0"'));
    assert('slide23 has ang="5400000" (90°)', slide23.includes('ang="5400000"'));
    assert('slide23 has ang="2700000" (45°)', slide23.includes('ang="2700000"'));
    // rotWithShape attribute
    assert('slide23 has rotWithShape="1"', slide23.includes('rotWithShape="1"'));
    assert('slide23 has rotWithShape="0"', slide23.includes('rotWithShape="0"'));
    assert('slide23 has Linear 0deg text', slide23.includes('Linear 0deg'));
  }

  // ── Slide 24: Radial/path gradient fills ──
  {
    section('test_features.pptx — Slide 24: radial/path gradient fills');
    const slide24 = textFiles.get('ppt/slides/slide24.xml');
    assert('slide24 exists', !!slide24);
    assert('slide24 has a:gradFill', hasTag(slide24, 'a:gradFill'));
    assert('slide24 has a:path (path gradient)', hasTag(slide24, 'a:path'));
    assert('slide24 has path="circle"', slide24.includes('path="circle"'));
    assert('slide24 has path="rect"', slide24.includes('path="rect"'));
    assert('slide24 has a:fillToRect', hasTag(slide24, 'a:fillToRect'));
    // Verify fillToRect values for centered circle
    assert('slide24 has l="50000" (center)', slide24.includes('l="50000"'));
    assert('slide24 has t="50000" (center)', slide24.includes('t="50000"'));
    // Off-center rect gradient
    assert('slide24 has l="25000" (off-center)', slide24.includes('l="25000"'));
    assert('slide24 has b="75000" (off-center)', slide24.includes('b="75000"'));
    // Ellipse shape with gradient
    assert('slide24 has ellipse geometry', slide24.includes('prst="ellipse"'));
    assert('slide24 has 9933FF (purple)', slide24.includes('9933FF'));
    assert('slide24 has Radial circle text', slide24.includes('Radial circle'));
    assert('slide24 has Ellipse grad text', slide24.includes('Ellipse grad'));
  }

  // ── Slide 25: Gradient background ──
  {
    section('test_features.pptx — Slide 25: gradient background');
    const slide25 = textFiles.get('ppt/slides/slide25.xml');
    assert('slide25 exists', !!slide25);
    assert('slide25 has p:bg', hasTag(slide25, 'p:bg'));
    assert('slide25 has p:bgPr', hasTag(slide25, 'p:bgPr'));
    assert('slide25 bg has a:gradFill', hasTag(slide25, 'a:gradFill'));
    assert('slide25 bg has a:gsLst', hasTag(slide25, 'a:gsLst'));
    assert('slide25 bg has a:lin', hasTag(slide25, 'a:lin'));
    // Background gradient colors
    assert('slide25 bg has 1B2838 (dark)', slide25.includes('1B2838'));
    assert('slide25 bg has 2A475E (mid)', slide25.includes('2A475E'));
    assert('slide25 bg has 66C0F4 (light)', slide25.includes('66C0F4'));
    assert('slide25 bg has ang="5400000" (90°)', slide25.includes('ang="5400000"'));
    assert('slide25 has Gradient Background text', slide25.includes('Gradient Background'));
  }

  // ── Slide 26: Alpha/transparency ──
  {
    section('test_features.pptx — Slide 26: alpha/transparency');
    const slide26 = textFiles.get('ppt/slides/slide26.xml');
    assert('slide26 exists', !!slide26);
    assert('slide26 has a:alpha', hasTag(slide26, 'a:alpha'));
    assert('slide26 has alpha val="50000"', slide26.includes('val="50000"'));
    assert('slide26 has FF0000 (red)', slide26.includes('FF0000'));
    // Gradient with alpha
    assert('slide26 has a:gradFill', hasTag(slide26, 'a:gradFill'));
    assert('slide26 has alpha val="80000"', slide26.includes('val="80000"'));
    assert('slide26 has alpha val="20000"', slide26.includes('val="20000"'));
    assert('slide26 has Alpha 50% text', slide26.includes('Alpha 50%'));
  }

  // ── Slide 27: Image fill on AutoShape ──
  {
    section('test_features.pptx — Slide 27: image fill (blipFill)');
    const slide27 = textFiles.get('ppt/slides/slide27.xml');
    assert('slide27 exists', !!slide27);
    // blipFill should be inside p:spPr
    assert('slide27 has a:blipFill', hasTag(slide27, 'a:blipFill'));
    assert('slide27 has a:blip', hasTag(slide27, 'a:blip'));
    assert('slide27 has r:embed on blip', slide27.includes('r:embed='));
    assert('slide27 has a:stretch', hasTag(slide27, 'a:stretch'));
  }

  // ── Slide 28: Pattern fill ──
  {
    section('test_features.pptx — Slide 28: pattern fill');
    const slide28 = textFiles.get('ppt/slides/slide28.xml');
    assert('slide28 exists', !!slide28);
    assert('slide28 has a:pattFill', hasTag(slide28, 'a:pattFill'));
    assert('slide28 has prst="ltDnDiag"', slide28.includes('prst="ltDnDiag"'));
    assert('slide28 has prst="smCheck"', slide28.includes('prst="smCheck"'));
    assert('slide28 has prst="dkHorz"', slide28.includes('prst="dkHorz"'));
    assert('slide28 has a:fgClr', hasTag(slide28, 'a:fgClr'));
    assert('slide28 has a:bgClr', hasTag(slide28, 'a:bgClr'));
    // Pattern colors
    assert('slide28 has 003366 (navy fg)', slide28.includes('003366'));
    assert('slide28 has CCCCCC (gray bg)', slide28.includes('CCCCCC'));
  }

  // ── Slide 29: Gradient tileFlip ──────────────────────────────────────────
  {
    section('test_features.pptx — Slide 29: gradient tileFlip');
    const slide29 = textFiles.get('ppt/slides/slide29.xml');
    assert('slide29 exists', !!slide29);
    assert('slide29 has a:gradFill', hasTag(slide29, 'a:gradFill'));
    assert('slide29 has tileFlip="x"', slide29.includes('tileFlip="x"'));
    assert('slide29 has tileFlip="y"', slide29.includes('tileFlip="y"'));
    assert('slide29 has tileFlip="xy"', slide29.includes('tileFlip="xy"'));
  }

  // ── Slide 30: Additional pattern fills ───────────────────────────────────
  {
    section('test_features.pptx — Slide 30: additional pattern fills');
    const slide30 = textFiles.get('ppt/slides/slide30.xml');
    assert('slide30 exists', !!slide30);
    assert('slide30 has a:pattFill', hasTag(slide30, 'a:pattFill'));
    assert('slide30 has prst="pct50"', slide30.includes('prst="pct50"'));
    assert('slide30 has prst="dnDiag"', slide30.includes('prst="dnDiag"'));
    assert('slide30 has prst="cross"', slide30.includes('prst="cross"'));
    assert('slide30 has prst="lgCheck"', slide30.includes('prst="lgCheck"'));
    assert('slide30 has prst="solidDmnd"', slide30.includes('prst="solidDmnd"'));
    assert('slide30 has prst="trellis"', slide30.includes('prst="trellis"'));
  }

  // ── Slide 31: Image fill tile ────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 31: image fill tile');
    const slide31 = textFiles.get('ppt/slides/slide31.xml');
    assert('slide31 exists', !!slide31);
    assert('slide31 has a:blipFill', hasTag(slide31, 'a:blipFill'));
    assert('slide31 has a:tile', hasTag(slide31, 'a:tile'));
    assert('slide31 has sx="50000"', slide31.includes('sx="50000"'));
    assert('slide31 has sy="50000"', slide31.includes('sy="50000"'));
    assert('slide31 has flip="xy"', slide31.includes('flip="xy"'));
    assert('slide31 has algn="tl"', slide31.includes('algn="tl"'));
  }

  // ── Slide 32: Stroke dash styles ────────────────────────────────────────
  {
    section('test_features.pptx — Slide 32: stroke dash styles');
    const slide32 = textFiles.get('ppt/slides/slide32.xml');
    assert('slide32 exists', !!slide32);
    assert('slide32 has a:ln', hasTag(slide32, 'a:ln'));
    assert('slide32 has a:prstDash', hasTag(slide32, 'a:prstDash'));
    assert('slide32 has val="dash"', slide32.includes('val="dash"'));
    assert('slide32 has val="dot"', slide32.includes('val="dot"'));
    assert('slide32 has val="dashDot"', slide32.includes('val="dashDot"'));
    assert('slide32 has a:custDash', hasTag(slide32, 'a:custDash'));
    assert('slide32 has a:ds (custom dash segment)', hasTag(slide32, 'a:ds'));
  }

  // ── Slide 33: Arrows, line join, line cap ───────────────────────────────
  {
    section('test_features.pptx — Slide 33: arrows/join/cap');
    const slide33 = textFiles.get('ppt/slides/slide33.xml');
    assert('slide33 exists', !!slide33);
    assert('slide33 has a:headEnd', hasTag(slide33, 'a:headEnd'));
    assert('slide33 has a:tailEnd', hasTag(slide33, 'a:tailEnd'));
    assert('slide33 has type="triangle"', slide33.includes('type="triangle"'));
    assert('slide33 has type="stealth"', slide33.includes('type="stealth"'));
    assert('slide33 has a:round', hasTag(slide33, 'a:round'));
    assert('slide33 has a:miter', hasTag(slide33, 'a:miter'));
    assert('slide33 has a:bevel', hasTag(slide33, 'a:bevel'));
    assert('slide33 has cap="rnd"', slide33.includes('cap="rnd"'));
    assert('slide33 has cap="sq"', slide33.includes('cap="sq"'));
    assert('slide33 has cmpd="dbl"', slide33.includes('cmpd="dbl"'));
    assert('slide33 has a:noFill', hasTag(slide33, 'a:noFill'));
  }

  // ── Slide 34: Group shapes ──────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 34: group shapes');
    const slide34 = textFiles.get('ppt/slides/slide34.xml');
    assert('slide34 exists', !!slide34);
    assert('slide34 has p:grpSp', hasTag(slide34, 'p:grpSp'));
    assert('slide34 has a:chOff', hasTag(slide34, 'a:chOff'));
    assert('slide34 has a:chExt', hasTag(slide34, 'a:chExt'));
    // Simple group has two child shapes
    assert('slide34 has FF6B6B (red)', slide34.includes('FF6B6B'));
    assert('slide34 has 4ECDC4 (teal)', slide34.includes('4ECDC4'));
    // Nested group
    assert('slide34 has FFD93D (yellow)', slide34.includes('FFD93D'));
    assert('slide34 has 6C5CE7 (purple)', slide34.includes('6C5CE7'));
    assert('slide34 has prst="ellipse"', slide34.includes('prst="ellipse"'));
    assert('slide34 has prst="roundRect"', slide34.includes('prst="roundRect"'));
  }

  // ── Slide 35: Connectors ────────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 35: connectors');
    const slide35 = textFiles.get('ppt/slides/slide35.xml');
    assert('slide35 exists', !!slide35);
    assert('slide35 has p:cxnSp', hasTag(slide35, 'p:cxnSp'));
    assert('slide35 has straightConnector1', slide35.includes('straightConnector1'));
    assert('slide35 has a:tailEnd', hasTag(slide35, 'a:tailEnd'));
    assert('slide35 has type="triangle"', slide35.includes('type="triangle"'));
    assert('slide35 has type="diamond"', slide35.includes('type="diamond"'));
    assert('slide35 has type="stealth"', slide35.includes('type="stealth"'));
    assert('slide35 has bentConnector3', slide35.includes('bentConnector3'));
    assert('slide35 has curvedConnector3', slide35.includes('curvedConnector3'));
  }

  // ── Slide 36: preset geometry shapes ──
  {
    console.log('\n── test_features.pptx — Slide 36: preset geometry shapes ──');
    const slide36 = textFiles.get('ppt/slides/slide36.xml');
    assert('slide36 exists', !!slide36);
    assert('slide36 has a:prstGeom', hasTag(slide36, 'a:prstGeom'));
    assert('slide36 has prst="triangle"', slide36.includes('prst="triangle"'));
    assert('slide36 has prst="diamond"', slide36.includes('prst="diamond"'));
    assert('slide36 has prst="pentagon"', slide36.includes('prst="pentagon"'));
    assert('slide36 has prst="hexagon"', slide36.includes('prst="hexagon"'));
    assert('slide36 has prst="rightArrow"', slide36.includes('prst="rightArrow"'));
    assert('slide36 has prst="star5"', slide36.includes('prst="star5"'));
    assert('slide36 has prst="heart"', slide36.includes('prst="heart"'));
    assert('slide36 has prst="plus"', slide36.includes('prst="plus"'));
    assert('slide36 has prst="chevron"', slide36.includes('prst="chevron"'));
    assert('slide36 has a:avLst with adj', slide36.includes('name="adj1"') || slide36.includes('name="adj"'));
  }

  // ── Slide 37: Custom geometry ──
  {
    console.log('\n── test_features.pptx — Slide 37: custom geometry ──');
    const slide37 = textFiles.get('ppt/slides/slide37.xml');
    assert('slide37 exists', !!slide37);
    assert('slide37 has a:custGeom', hasTag(slide37, 'a:custGeom'));
    assert('slide37 has a:pathLst', hasTag(slide37, 'a:pathLst'));
    assert('slide37 has a:moveTo', hasTag(slide37, 'a:moveTo'));
    assert('slide37 has a:lnTo', hasTag(slide37, 'a:lnTo'));
    assert('slide37 has a:cubicBezTo', hasTag(slide37, 'a:cubicBezTo'));
    assert('slide37 has a:close', hasTag(slide37, 'a:close'));
    // Guide formula shape
    assert('slide37 has a:avLst', hasTag(slide37, 'a:avLst'));
    assert('slide37 has a:gdLst', hasTag(slide37, 'a:gdLst'));
    assert('slide37 has FFD700 (gold)', slide37.includes('FFD700'));
    assert('slide37 has 87CEEB (sky blue)', slide37.includes('87CEEB'));
    assert('slide37 has 98FB98 (pale green)', slide37.includes('98FB98'));
  }

  // ── Slide 38: Gear shapes ──
  {
    console.log('\n── test_features.pptx — Slide 38: gear shapes ──');
    const slide38 = textFiles.get('ppt/slides/slide38.xml');
    assert('slide38 exists', !!slide38);
    assert('slide38 has gear6', slide38.includes('gear6'));
    assert('slide38 has gear9', slide38.includes('gear9'));
    assert('slide38 has 4472C4 (blue)', slide38.includes('4472C4'));
    assert('slide38 has ED7D31 (orange)', slide38.includes('ED7D31'));
    assert('slide38 has adj fmla', slide38.includes('val 50000'));
  }

  // ── Slide 39: Text rectangles ──
  {
    console.log('\n── test_features.pptx — Slide 39: text rectangles ──');
    const slide39 = textFiles.get('ppt/slides/slide39.xml');
    assert('slide39 exists', !!slide39);
    assert('slide39 has triangle', slide39.includes('triangle'));
    assert('slide39 has diamond', slide39.includes('diamond'));
    assert('slide39 has rightArrow', slide39.includes('rightArrow'));
    assert('slide39 has text body', hasTag(slide39, 'p:txBody'));
    assert('slide39 has FFD700 (gold)', slide39.includes('FFD700'));
  }

  // ── Slide 40: Connection points ──
  {
    console.log('\n── test_features.pptx — Slide 40: connection points ──');
    const slide40 = textFiles.get('ppt/slides/slide40.xml');
    assert('slide40 exists', !!slide40);
    assert('slide40 has p:cxnSp', hasTag(slide40, 'p:cxnSp'));
    assert('slide40 has a:stCxn', hasTag(slide40, 'a:stCxn'));
    assert('slide40 has a:endCxn', hasTag(slide40, 'a:endCxn'));
    assert('slide40 has a:cxnLst', hasTag(slide40, 'a:cxnLst'));
    assert('slide40 has a:pos', hasTag(slide40, 'a:pos'));
    assert('slide40 has a:custGeom', hasTag(slide40, 'a:custGeom'));
    assert('slide40 has a:rect', hasTag(slide40, 'a:rect'));
  }

  // ── Slide 41: Table cell merge + borders + margins + anchor ──
  {
    console.log('\n── test_features.pptx — Slide 41: table cell merge/borders/margins/anchor ──');
    const slide41 = textFiles.get('ppt/slides/slide41.xml');
    assert('slide41 exists', !!slide41);
    assert('slide41 has a:tbl', hasTag(slide41, 'a:tbl'));
    assert('slide41 has gridSpan', slide41.includes('gridSpan'));
    assert('slide41 has rowSpan', slide41.includes('rowSpan'));
    assert('slide41 has vMerge', slide41.includes('vMerge'));
    assert('slide41 has a:lnL', hasTag(slide41, 'a:lnL'));
    assert('slide41 has a:lnR', hasTag(slide41, 'a:lnR'));
    assert('slide41 has a:lnT', hasTag(slide41, 'a:lnT'));
    assert('slide41 has a:lnB', hasTag(slide41, 'a:lnB'));
    assert('slide41 has anchor="ctr"', slide41.includes('anchor="ctr"'));
    assert('slide41 has anchor="b"', slide41.includes('anchor="b"'));
    assert('slide41 has marL', slide41.includes('marL'));
    assert('slide41 has marT', slide41.includes('marT'));
    assert('slide41 has a:noFill in border', slide41.includes('<a:lnL') && slide41.includes('<a:noFill'));
    assert('slide41 has red border FF0000', slide41.includes('FF0000'));
    assert('slide41 has green border 00AA00', slide41.includes('00AA00'));
  }

  // ── Slide 42: Diagonal borders + tblPr flags + tblStyleId ──
  {
    console.log('\n── test_features.pptx — Slide 42: diagonal borders + tblPr flags ──');
    const slide42 = textFiles.get('ppt/slides/slide42.xml');
    assert('slide42 exists', !!slide42);
    assert('slide42 has a:tbl', hasTag(slide42, 'a:tbl'));
    // Diagonal borders
    assert('slide42 has a:lnTlToBr', hasTag(slide42, 'a:lnTlToBr'));
    assert('slide42 has a:lnBlToTr', hasTag(slide42, 'a:lnBlToTr'));
    // tblPr flags
    assert('slide42 has firstRow="1"', slide42.includes('firstRow="1"'));
    assert('slide42 has lastRow="1"', slide42.includes('lastRow="1"'));
    assert('slide42 has bandRow="1"', slide42.includes('bandRow="1"'));
    assert('slide42 has tblStyleId', slide42.includes('tblStyleId'));
    // Diagonal border colors
    assert('slide42 has red diagonal FF0000', slide42.includes('FF0000'));
    assert('slide42 has blue diagonal 0000FF', slide42.includes('0000FF'));
    assert('slide42 has green diagonal 008800', slide42.includes('008800'));
  }

  // ── Slide 43: Image crop + alpha ──
  {
    console.log('\n── test_features.pptx — Slide 43: image crop + alpha ──');
    const slide43 = textFiles.get('ppt/slides/slide43.xml');
    assert('slide43 exists', !!slide43);
    assert('slide43 has p:pic', hasTag(slide43, 'p:pic'));
    assert('slide43 has a:blip', hasTag(slide43, 'a:blip'));
    // Crop attributes
    assert('slide43 has a:srcRect', hasTag(slide43, 'a:srcRect'));
    assert('slide43 has crop l="25000"', slide43.includes('l="25000"'));
    assert('slide43 has crop t="25000"', slide43.includes('t="25000"'));
    // Alpha
    assert('slide43 has a:alphaModFix', hasTag(slide43, 'a:alphaModFix'));
    assert('slide43 has amt="50000"', slide43.includes('amt="50000"'));
    assert('slide43 has amt="75000"', slide43.includes('amt="75000"'));
    // AutoShape with blipFill crop
    assert('slide43 has a:blipFill', hasTag(slide43, 'a:blipFill'));
    assert('slide43 has blipFill crop t="50000"', slide43.includes('t="50000"'));
  }

  // ── Slide 44: External image reference ──
  {
    console.log('\n── test_features.pptx — Slide 44: external image reference ──');
    const slide44 = textFiles.get('ppt/slides/slide44.xml');
    assert('slide44 exists', !!slide44);
    assert('slide44 has p:pic', hasTag(slide44, 'p:pic'));
    const rels44 = textFiles.get('ppt/slides/_rels/slide44.xml.rels');
    assert('slide44 rels exists', !!rels44);
    assert('slide44 has TargetMode="External"', rels44.includes('TargetMode="External"'));
    assert('slide44 has wikimedia URL', rels44.includes('upload.wikimedia.org'));
    assert('slide44 has picsum URL', rels44.includes('picsum.photos'));
  }

  // ── Slide 45: Image effects — brightness/contrast (a:lum) ──
  section('test_features.pptx — Slide 45: brightness/contrast');
  {
    const slide45 = textFiles.get('ppt/slides/slide45.xml') ?? '';
    assert('slide45 has p:pic', hasTag(slide45, 'p:pic'));
    assert('slide45 has a:lum', hasTag(slide45, 'a:lum'));
    assert('slide45 has bright="50000"', slide45.includes('bright="50000"'));
    assert('slide45 has contrast="-30000"', slide45.includes('contrast="-30000"'));
    assert('slide45 has bright="20000"', slide45.includes('bright="20000"'));
    assert('slide45 has contrast="40000"', slide45.includes('contrast="40000"'));
  }

  // ── Slide 46: Duotone + color change ──
  section('test_features.pptx — Slide 46: duotone + clrChange');
  {
    const slide46 = textFiles.get('ppt/slides/slide46.xml') ?? '';
    assert('slide46 has p:pic', hasTag(slide46, 'p:pic'));
    assert('slide46 has a:duotone', hasTag(slide46, 'a:duotone'));
    assert('slide46 has duotone color 000080', slide46.includes('val="000080"'));
    assert('slide46 has duotone color FFFF00', slide46.includes('val="FFFF00"'));
    assert('slide46 has a:clrChange', hasTag(slide46, 'a:clrChange'));
    assert('slide46 has a:clrFrom', hasTag(slide46, 'a:clrFrom'));
    assert('slide46 has a:clrTo', hasTag(slide46, 'a:clrTo'));
  }

  // ── Slide 47: Background pattern fill ──
  section('test_features.pptx — Slide 47: background pattern fill');
  {
    const slide47 = textFiles.get('ppt/slides/slide47.xml') ?? '';
    assert('slide47 has p:bg', hasTag(slide47, 'p:bg'));
    assert('slide47 has p:bgPr', hasTag(slide47, 'p:bgPr'));
    assert('slide47 has a:pattFill', hasTag(slide47, 'a:pattFill'));
    assert('slide47 has ltDnDiag', slide47.includes('ltDnDiag'));
    assert('slide47 has fg color 3366CC', slide47.includes('3366CC'));
  }

  // ── Slide 48: Line gradient/pattern fill ──
  section('test_features.pptx — Slide 48: line gradient/pattern fill');
  {
    const slide48 = textFiles.get('ppt/slides/slide48.xml') ?? '';
    assert('slide48 has a:ln with gradFill', slide48.includes('<a:gradFill>') || slide48.includes('<a:gradFill '));
    assert('slide48 has a:ln with pattFill', slide48.includes('smCheck'));
    assert('slide48 has gradient stop FF0000', slide48.includes('FF0000'));
    assert('slide48 has gradient stop 0000FF', slide48.includes('0000FF'));
  }

  // ── Slide 49: Shape hyperlinks + color modifiers ──
  section('test_features.pptx — Slide 49: shape hyperlinks + color modifiers');
  {
    const slide49 = textFiles.get('ppt/slides/slide49.xml') ?? '';
    assert('slide49 has a:hlinkClick', hasTag(slide49, 'a:hlinkClick'));
    assert('slide49 has a:comp', hasTag(slide49, 'a:comp'));
    assert('slide49 has a:inv', hasTag(slide49, 'a:inv'));
    assert('slide49 has a:hueMod', hasTag(slide49, 'a:hueMod'));
    assert('slide49 has hueMod val 50000', slide49.includes('50000'));
  }

  // ── Slide 50: Shape effects ──
  section('test_features.pptx — Slide 50: shape effects');
  {
    const slide50 = textFiles.get('ppt/slides/slide50.xml') ?? '';
    assert('slide50 has a:effectLst', hasTag(slide50, 'a:effectLst'));
    assert('slide50 has a:outerShdw', hasTag(slide50, 'a:outerShdw'));
    assert('slide50 has a:innerShdw', hasTag(slide50, 'a:innerShdw'));
    assert('slide50 has a:glow', hasTag(slide50, 'a:glow'));
    assert('slide50 has a:softEdge', hasTag(slide50, 'a:softEdge'));
    assert('slide50 outerShdw blurRad', slide50.includes('blurRad="152400"'));
    assert('slide50 outerShdw dir', slide50.includes('dir="5400000"'));
    assert('slide50 glow rad', slide50.includes('rad="228600"'));
  }

  // ── Slide 51: 3D effects + text shadow ──
  section('test_features.pptx — Slide 51: 3D effects + text shadow');
  {
    const slide51 = textFiles.get('ppt/slides/slide51.xml') ?? '';
    assert('slide51 has a:scene3d', hasTag(slide51, 'a:scene3d'));
    assert('slide51 has a:camera', hasTag(slide51, 'a:camera'));
    assert('slide51 has a:lightRig', hasTag(slide51, 'a:lightRig'));
    assert('slide51 has a:sp3d', hasTag(slide51, 'a:sp3d'));
    assert('slide51 has a:bevelT', hasTag(slide51, 'a:bevelT'));
    assert('slide51 has a:extrusionClr', hasTag(slide51, 'a:extrusionClr'));
    assert('slide51 has a:contourClr', hasTag(slide51, 'a:contourClr'));
    assert('slide51 has text effectLst', slide51.includes('<a:effectLst>'));
    assert('slide51 has text outerShdw', slide51.includes('a:outerShdw'));
    assert('slide51 has text glow', slide51.includes('a:glow'));
    assert('slide51 has prstMaterial', slide51.includes('prstMaterial'));
  }

  // ── Slide 52: Column chart ──
  {
    const slide52 = textFiles.get('ppt/slides/slide52.xml') || '';
    const rels52 = textFiles.get('ppt/slides/_rels/slide52.xml.rels') || '';
    section('test_features.pptx — Slide 52: column chart');
    assert('slide52 has p:graphicFrame', slide52.includes('p:graphicFrame'));
    assert('slide52 has a:graphicData', slide52.includes('a:graphicData'));
    assert('slide52 has c:chart', slide52.includes('c:chart'));
    assert('slide52 rels has chart ref', rels52.includes('/chart'));
    // Check chart XML exists
    const chartTarget52 = rels52.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget52) {
      const chartPath = 'ppt/charts/' + chartTarget52[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide52 chart XML exists', chartXml.length > 0);
      assert('slide52 has c:chartSpace', chartXml.includes('c:chartSpace'));
      assert('slide52 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide52 has barDir col', chartXml.includes('val="col"'));
      assert('slide52 has c:ser', chartXml.includes('c:ser'));
      assert('slide52 has c:numCache', chartXml.includes('c:numCache'));
      assert('slide52 has c:strCache', chartXml.includes('c:strCache'));
      assert('slide52 has Q1 category', chartXml.includes('Q1'));
      assert('slide52 has Sales 2024', chartXml.includes('Sales 2024'));
    }
  }

  // ── Slide 53: Line chart ──
  {
    const slide53 = textFiles.get('ppt/slides/slide53.xml') || '';
    section('test_features.pptx — Slide 53: line chart');
    assert('slide53 has p:graphicFrame', slide53.includes('p:graphicFrame'));
    assert('slide53 has c:chart', slide53.includes('c:chart'));
    const rels53 = textFiles.get('ppt/slides/_rels/slide53.xml.rels') || '';
    const chartTarget53 = rels53.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget53) {
      const chartPath = 'ppt/charts/' + chartTarget53[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide53 chart has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide53 chart has Revenue', chartXml.includes('Revenue'));
      assert('slide53 chart has Jan', chartXml.includes('Jan'));
    }
  }

  // ── Slide 54: Pie chart ──
  {
    const slide54 = textFiles.get('ppt/slides/slide54.xml') || '';
    section('test_features.pptx — Slide 54: pie chart');
    assert('slide54 has p:graphicFrame', slide54.includes('p:graphicFrame'));
    const rels54 = textFiles.get('ppt/slides/_rels/slide54.xml.rels') || '';
    const chartTarget54 = rels54.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget54) {
      const chartPath = 'ppt/charts/' + chartTarget54[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide54 chart has c:pieChart', chartXml.includes('c:pieChart'));
      assert('slide54 chart has Desktop', chartXml.includes('Desktop'));
      assert('slide54 chart has Market Share', chartXml.includes('Market Share'));
    }
  }

  // ── Slide 55: Bar + Donut chart ──
  {
    const slide55 = textFiles.get('ppt/slides/slide55.xml') || '';
    section('test_features.pptx — Slide 55: bar + donut chart');
    assert('slide55 has p:graphicFrame', slide55.includes('p:graphicFrame'));
    const rels55 = textFiles.get('ppt/slides/_rels/slide55.xml.rels') || '';
    assert('slide55 has chart refs', rels55.includes('/chart'));
    // Check for both chart files
    const chartTargets55 = [...rels55.matchAll(/Target="([^"]*chart[^"]*)"/g)];
    assert('slide55 has 2 chart refs', chartTargets55.length >= 2);
  }

  // ── Slide 56: Column chart with data labels ──
  {
    const slide56 = textFiles.get('ppt/slides/slide56.xml') || '';
    section('test_features.pptx — Slide 56: column chart with data labels');
    assert('slide56 has p:graphicFrame', slide56.includes('p:graphicFrame'));
    const rels56 = textFiles.get('ppt/slides/_rels/slide56.xml.rels') || '';
    const chartTarget56 = rels56.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget56) {
      const chartPath = 'ppt/charts/' + chartTarget56[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide56 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide56 has c:dLbls', chartXml.includes('c:dLbls'));
      assert('slide56 has showVal', chartXml.includes('showVal'));
      assert('slide56 has North category', chartXml.includes('North'));
    }
  }

  // ── Slide 57: Pie chart with dPt colors + % labels ──
  {
    const slide57 = textFiles.get('ppt/slides/slide57.xml') || '';
    section('test_features.pptx — Slide 57: pie chart with dPt + % labels');
    assert('slide57 has p:graphicFrame', slide57.includes('p:graphicFrame'));
    const rels57 = textFiles.get('ppt/slides/_rels/slide57.xml.rels') || '';
    const chartTarget57 = rels57.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget57) {
      const chartPath = 'ppt/charts/' + chartTarget57[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide57 has c:pieChart', chartXml.includes('c:pieChart'));
      assert('slide57 has c:dPt', chartXml.includes('c:dPt'));
      assert('slide57 has srgbClr', chartXml.includes('srgbClr'));
      assert('slide57 has showPercent', chartXml.includes('showPercent'));
      assert('slide57 has Chrome category', chartXml.includes('Chrome'));
    }
  }

  // ── Slide 58: Line chart with series colors + data labels ──
  {
    const slide58 = textFiles.get('ppt/slides/slide58.xml') || '';
    section('test_features.pptx — Slide 58: line chart with series colors');
    assert('slide58 has p:graphicFrame', slide58.includes('p:graphicFrame'));
    const rels58 = textFiles.get('ppt/slides/_rels/slide58.xml.rels') || '';
    const chartTarget58 = rels58.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget58) {
      const chartPath = 'ppt/charts/' + chartTarget58[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide58 has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide58 has c:dLbls', chartXml.includes('c:dLbls'));
      assert('slide58 has spPr with color', chartXml.includes('srgbClr'));
      assert('slide58 has Week 1', chartXml.includes('Week 1'));
    }
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
