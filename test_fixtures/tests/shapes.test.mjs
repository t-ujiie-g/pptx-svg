import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("shapes & geometry (slides 34-40)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

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


  finishAssertions();
});
