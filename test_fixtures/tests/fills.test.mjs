import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("fills & strokes (slides 23-33)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();


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


  finishAssertions();
});
