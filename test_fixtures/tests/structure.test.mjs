import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("structure & master/layout/theme inheritance", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

  // ── Basic structure ──
  section('test_features.pptx — basic structure');
  const prsXml = textFiles.get('ppt/presentation.xml');
  assert('presentation.xml exists', !!prsXml);

  const slideCount = countSlideIds(prsXml ?? '');
  assert('slide count = 95', slideCount === 95, `got ${slideCount}`);

  // Verify all slides exist
  for (let i = 1; i <= 95; i++) {
    const path = `ppt/slides/slide${i}.xml`;
    assert(`slide${i}.xml exists`, textFiles.has(path));
  }

  // ── Slide .rels ──
  section('test_features.pptx — slide relationships');
  for (let i = 1; i <= 95; i++) {
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

  finishAssertions();
});
