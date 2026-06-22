import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  assert, hasTag,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
  FIXTURES_DIR, DIST_DIR,
} from './_helpers.mjs';

test("regressions (slides 96-98)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

  // ── Slide 96: empty custGeom + fontRef text color ──────────────────────
  {
    console.log('\n── test_features.pptx — Slide 96: regressions ──');
    const slide96 = textFiles.get('ppt/slides/slide96.xml');
    assert('slide96 exists', !!slide96);
    if (!slide96) {
      finishAssertions();
      return;
    }

    // Reference rect: must carry <p:style>/<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
    // so the renderer's font_ref_color fallback gives white text.
    assert('slide96 has <p:style> on the reference rect', hasTag(slide96, 'p:style'));
    assert('slide96 has <a:fontRef> inside <p:style>', hasTag(slide96, 'a:fontRef'));
    assert(
      'slide96 fontRef targets lt1 (white text)',
      slide96.includes('<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'),
    );
    assert(
      'slide96 reference run has no explicit color',
      slide96.includes('<a:r><a:t>Reference rect (prstGeom)</a:t></a:r>'),
    );

    // Empty custGeom: the issue #39 reproduction. The shape must be present
    // in OOXML (so the renderer exercises the suppression path) but have an
    // empty path with no draw commands.
    assert('slide96 has EmptyCustGeom shape', slide96.includes('EmptyCustGeom'));
    assert('slide96 EmptyCustGeom has solidFill 00B4D8', slide96.includes('00B4D8'));
    assert(
      'slide96 EmptyCustGeom path has no moveTo / lnTo / cubicBezTo',
      slide96.includes('<a:pathLst><a:path w="508000" h="508000"/></a:pathLst>'),
    );

    // Valid custGeom: control case — confirms the fix doesn't break normal
    // custGeom rendering.
    assert('slide96 has ValidCustGeom shape', slide96.includes('ValidCustGeom'));
    assert('slide96 ValidCustGeom has fill 06A77D', slide96.includes('06A77D'));
    assert('slide96 ValidCustGeom has a:moveTo', hasTag(slide96, 'a:moveTo'));
    assert('slide96 ValidCustGeom has a:lnTo', hasTag(slide96, 'a:lnTo'));
  }

  // ── Slide 97: header/footer field placeholders (date / footer / slide num) ──
  {
    console.log('\n── test_features.pptx — Slide 97: header/footer fields ──');
    const slide97 = textFiles.get('ppt/slides/slide97.xml');
    assert('slide97 exists', !!slide97);
    if (slide97) {
      // Date + slide-number fields and footer placeholder must be preserved in
      // OOXML (round-trip); the renderer fills the actual values at render time.
      assert('slide97 has date field', slide97.includes('type="datetime1"'));
      assert('slide97 has slide-number field', slide97.includes('type="slidenum"'));
      assert('slide97 has dt placeholder', slide97.includes('type="dt"'));
      assert('slide97 has ftr placeholder', slide97.includes('type="ftr"'));
      assert('slide97 has sldNum placeholder', slide97.includes('type="sldNum"'));
      assert('slide97 footer text preserved', slide97.includes('moon-pptx footer'));
    }
  }

  // ── Slide 98: underline styles + underline color (a:uFill) + double strike ──
  {
    console.log('\n── test_features.pptx — Slide 98: text decoration fidelity ──');
    const slide98 = textFiles.get('ppt/slides/slide98.xml');
    assert('slide98 exists', !!slide98);
    if (slide98) {
      assert('slide98 has u="dbl"', slide98.includes('u="dbl"'));
      assert('slide98 has u="wavy"', slide98.includes('u="wavy"'));
      assert('slide98 has u="dotted"', slide98.includes('u="dotted"'));
      assert('slide98 has underline color a:uFill', slide98.includes('<a:uFill>'));
      assert('slide98 has dblStrike', slide98.includes('strike="dblStrike"'));
    }
  }

  // ── Slide 99: flipH/flipV mirroring (issue #55) ─────────────────────────────
  {
    console.log('\n── test_features.pptx — Slide 99: flipH/flipV mirroring ──');
    const slide99 = textFiles.get('ppt/slides/slide99.xml');
    assert('slide99 exists', !!slide99);
    if (slide99) {
      // OOXML round-trip: the flip attributes survive in the slide XML.
      assert('slide99 has flipH', slide99.includes('flipH="1"'));
      assert('slide99 has flipV', slide99.includes('flipV="1"'));
    }
    // Rendering: the renderer must emit a negative-scale mirror for the flips
    // (the actual issue #55 bug — flips were dropped, only rotate() was kept).
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const svg99 = r.renderSlideSvg(98); // 0-based → slide 99
    assert('slide99 SVG mirrors flipH (scale(-1,1))', svg99.includes('scale(-1,1)'));
    assert('slide99 SVG mirrors flipV (scale(1,-1))', svg99.includes('scale(1,-1)'));
  }

  finishAssertions();
});
