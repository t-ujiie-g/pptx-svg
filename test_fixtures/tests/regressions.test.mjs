import { test } from 'node:test';
import {
  assert, hasTag,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("regressions (slide 96)", async () => {
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

  finishAssertions();
});
