import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("images & shape effects (slides 43-51)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

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

  finishAssertions();
});
