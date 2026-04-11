import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("text features (slides 9-22, 70-72, 88)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

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

  // ── Slide 70: Text outline ──────────────────────────────────────────────────
  {
    const slide70 = textFiles.get('ppt/slides/slide70.xml') || '';
    section('test_features.pptx — Slide 70: Text outline');
    assert('slide70 has a:ln in rPr', slide70.includes('<a:ln'));
    assert('slide70 has w="25400"', slide70.includes('w="25400"'));
    assert('slide70 has FF0000 outline', slide70.includes('FF0000'));
    assert('slide70 has 0000FF outline', slide70.includes('0000FF'));
    assert('slide70 has Red Outlined Text', slide70.includes('Red Outlined Text'));
  }

  // ── Slide 71: Text gradient fill ────────────────────────────────────────────
  {
    const slide71 = textFiles.get('ppt/slides/slide71.xml') || '';
    section('test_features.pptx — Slide 71: Text gradient fill');
    assert('slide71 has a:gradFill', slide71.includes('a:gradFill'));
    assert('slide71 has gradient stops', slide71.includes('a:gs'));
    assert('slide71 has Gradient Text', slide71.includes('Gradient Text'));
    assert('slide71 has FF0000 stop', slide71.includes('FF0000'));
    assert('slide71 has FFFF00 stop', slide71.includes('FFFF00'));
  }

  // ── Slide 72: Text warp ─────────────────────────────────────────────────────
  {
    const slide72 = textFiles.get('ppt/slides/slide72.xml') || '';
    section('test_features.pptx — Slide 72: Text warp');
    assert('slide72 has a:prstTxWarp', slide72.includes('a:prstTxWarp'));
    assert('slide72 has textWave1', slide72.includes('textWave1'));
    assert('slide72 has textArchUp', slide72.includes('textArchUp'));
    assert('slide72 has textDeflate', slide72.includes('textDeflate'));
    assert('slide72 has a:avLst', slide72.includes('a:avLst'));
    assert('slide72 has Wave Text', slide72.includes('Wave Text'));
  }



  // ── Slide 88: Justified text ───────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 88: justified text');
    const slide88 = textFiles.get('ppt/slides/slide89.xml') || '';
    assert('slide88 exists', slide88.length > 0);
    assert('slide88 has algn=just', slide88.includes('algn="just"'));
  }

  finishAssertions();
});
