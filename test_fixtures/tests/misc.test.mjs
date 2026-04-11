import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("misc features (slides 74-87)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

  // ── Slide 74: Speaker notes + comments ──────────────────────────────────────
  {
    section('test_features.pptx — Slide 74: Speaker notes + comments');
    const slide74 = textFiles.get('ppt/slides/slide75.xml') || '';
    assert('slide74 has text content', slide74.includes('speaker notes'));

    // Notes
    const rels74 = textFiles.get('ppt/slides/_rels/slide75.xml.rels') || '';
    const notesRefs = findRelTarget(rels74, 'notesSlide');
    assert('slide74 has notesSlide relationship', notesRefs.length > 0);

    // Find and check notes content
    let notesPath = '';
    for (const ref of notesRefs) {
      const resolved = 'ppt/notesSlides/' + ref.replace('../notesSlides/', '');
      if (textFiles.has(resolved)) { notesPath = resolved; break; }
    }
    assert('notesSlide XML exists', notesPath !== '');
    if (notesPath) {
      const notesXml = textFiles.get(notesPath) || '';
      assert('notes contain speaker notes text', notesXml.includes('speaker notes for slide 74'));
      assert('notes have multiple paragraphs', notesXml.includes('round-trip preservation'));
    }

    // Comments
    const commentRefs = findRelTarget(rels74, 'comments');
    assert('slide74 has comments relationship', commentRefs.length > 0);

    // Check commentAuthors
    const authorsXml = textFiles.get('ppt/commentAuthors.xml') || '';
    assert('commentAuthors.xml exists', authorsXml.length > 0);
    assert('commentAuthors has Test User', authorsXml.includes('name="Test User"'));
    assert('commentAuthors has initials', authorsXml.includes('initials="TU"'));

    // Check comments content
    let commentsPath = '';
    for (const ref of commentRefs) {
      const resolved = 'ppt/comments/' + ref.replace('../comments/', '');
      if (textFiles.has(resolved)) { commentsPath = resolved; break; }
    }
    assert('comments XML exists', commentsPath !== '');
    if (commentsPath) {
      const commentsXml = textFiles.get(commentsPath) || '';
      assert('comments contain test comment', commentsXml.includes('test comment on slide 74'));
      assert('comments contain review feedback', commentsXml.includes('review feedback'));
      assert('comments have position data', commentsXml.includes('<p:pos'));
      assert('comments have 2 entries', (commentsXml.match(/<p:cm\b/g) || []).length === 2);
    }
  }

  // ── Slide 75: SmartArt fallback (mc:AlternateContent) ─────────────────────
  {
    section('test_features.pptx — Slide 75: SmartArt fallback');
    const slide75 = textFiles.get('ppt/slides/slide76.xml') || '';
    assert('slide75 exists', slide75.length > 0);
    assert('slide75 has mc:AlternateContent', slide75.includes('mc:AlternateContent'));
    assert('slide75 has mc:Choice', slide75.includes('mc:Choice'));
    assert('slide75 has mc:Fallback', slide75.includes('mc:Fallback'));
    assert('slide75 has grpSp in fallback', slide75.includes('p:grpSp'));
    // Check that the fallback contains our 3 process boxes
    assert('slide75 has roundRect shapes', slide75.includes('roundRect'));
    const spCount = (slide75.match(/<p:sp\b/g) || []).length;
    assert('slide75 has at least 4 p:sp elements (3 boxes + title)', spCount >= 4);
    // Check text content of SmartArt boxes
    assert('slide75 has "Plan" text', slide75.includes('Plan'));
    assert('slide75 has "Build" text', slide75.includes('Build'));
    assert('slide75 has "Ship" text', slide75.includes('Ship'));
    // Check fill colors
    assert('slide75 has blue fill (4472C4)', slide75.includes('4472C4'));
    assert('slide75 has orange fill (ED7D31)', slide75.includes('ED7D31'));
    assert('slide75 has green fill (70AD47)', slide75.includes('70AD47'));
  }

  // ── Slide 76: OLE embedded object ──────────────────────────────────────────
  {
    section('test_features.pptx — Slide 76: OLE embedded object');
    const slide76 = textFiles.get('ppt/slides/slide77.xml') || '';
    assert('slide76 exists', slide76.length > 0);
    assert('slide76 has p:graphicFrame', slide76.includes('p:graphicFrame'));
    assert('slide76 has p:oleObj', slide76.includes('p:oleObj'));
    assert('slide76 has OLE URI', slide76.includes('presentationml/2006/ole'));
    assert('slide76 has progId Excel.Sheet.12', slide76.includes('Excel.Sheet.12'));
    assert('slide76 has p:embed', slide76.includes('p:embed'));
    assert('slide76 has fallback p:pic', slide76.includes('p:pic'));
    assert('slide76 has blip r:embed', slide76.includes('r:embed'));
    assert('slide76 has OLE name', slide76.includes('Embedded Spreadsheet'));
    // Check rels for image and OLE binary
    const rels76 = textFiles.get('ppt/slides/_rels/slide77.xml.rels') || '';
    assert('slide76 has image relationship', rels76.includes('image'));
    assert('slide76 has oleObject relationship', rels76.includes('oleObject'));
  }

  // ── Slide 77: Media (video with poster frame) ─────────────────────────────
  {
    section('test_features.pptx — Slide 77: Media (video with poster frame)');
    const slide77 = textFiles.get('ppt/slides/slide78.xml') || '';
    assert('slide77 exists', slide77.length > 0);
    assert('slide77 has p:pic', slide77.includes('p:pic'));
    assert('slide77 has a:videoFile', slide77.includes('a:videoFile'));
    assert('slide77 has r:link for video', slide77.includes('r:link'));
    assert('slide77 has p:blipFill for poster', slide77.includes('p:blipFill'));
    assert('slide77 has a:blip r:embed', slide77.includes('r:embed'));
    // Check rels for video and image
    const rels77 = textFiles.get('ppt/slides/_rels/slide78.xml.rels') || '';
    assert('slide77 has image relationship', rels77.includes('image'));
    assert('slide77 has video relationship', rels77.includes('video'));
  }

  // ── Slide 78: Math equation (OMML) ──────────────────────────────────────────
  {
    section('test_features.pptx — Slide 78: Math equation (OMML)');
    const slide78 = textFiles.get('ppt/slides/slide79.xml') || '';
    assert('slide78 exists', slide78.length > 0);
    assert('slide78 has m:oMathPara', slide78.includes('m:oMathPara'));
    assert('slide78 has m:oMath', slide78.includes('m:oMath'));
    assert('slide78 has m:r (math run)', slide78.includes('<m:r>'));
    assert('slide78 has m:t (math text)', slide78.includes('<m:t>'));
    assert('slide78 has m:f (fraction)', slide78.includes('<m:f>'));
    assert('slide78 has m:rad (radical)', slide78.includes('<m:rad>'));
    assert('slide78 has m:sSup (superscript)', slide78.includes('<m:sSup>'));
    assert('slide78 has xmlns:m namespace', slide78.includes('xmlns:m='));
  }

  // ── Slide 79: Transition + Timing (round-trip) ──────────────────────────────
  {
    section('test_features.pptx — Slide 79: Transition + Timing');
    const slide79 = textFiles.get('ppt/slides/slide80.xml') || '';
    assert('slide79 exists', slide79.length > 0);
    assert('slide79 has p:transition', slide79.includes('<p:transition'));
    assert('slide79 transition has fade', slide79.includes('<p:fade'));
    assert('slide79 has p:timing', slide79.includes('<p:timing'));
    assert('slide79 has p:tnLst', slide79.includes('<p:tnLst>'));
  }

  // ── Slide 80: Hidden slide ──────────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 80: Hidden slide');
    const slide80 = textFiles.get('ppt/slides/slide81.xml') || '';
    assert('slide80 exists', slide80.length > 0);
    assert('slide80 has show="0"', slide80.includes('show="0"'));
  }

  // ── Slide 82: OMML — Large operators + delimiters ───────────────────────────
  {
    section('test_features.pptx — Slide 82: OMML nary');
    const slide82 = textFiles.get('ppt/slides/slide83.xml') || '';
    assert('slide82 exists', slide82.length > 0);
    assert('slide82 has m:nary', slide82.includes('m:nary'));
    assert('slide82 has integral char', slide82.includes('\u222B') || slide82.includes('&#x222B;'));
    assert('slide82 has m:limLoc', slide82.includes('m:limLoc'));
    assert('slide82 has m:sSup', slide82.includes('m:sSup'));
  }

  // ── Slide 83: OMML — Matrix + delimiters ──────────────────────────────────
  {
    section('test_features.pptx — Slide 83: OMML matrix');
    const slide83 = textFiles.get('ppt/slides/slide84.xml') || '';
    assert('slide83 exists', slide83.length > 0);
    assert('slide83 has m:m (matrix)', slide83.includes('m:m'));
    assert('slide83 has m:mr (matrix row)', slide83.includes('m:mr'));
    assert('slide83 has m:d (delimiter)', slide83.includes('m:d'));
    assert('slide83 has m:begChr', slide83.includes('m:begChr'));
  }

  // ── Slide 84: OMML — Accent + bar + subsup ───────────────────────────────
  {
    section('test_features.pptx — Slide 84: OMML accent/bar/subsup');
    const slide84 = textFiles.get('ppt/slides/slide85.xml') || '';
    assert('slide84 exists', slide84.length > 0);
    assert('slide84 has m:acc (accent)', slide84.includes('m:acc'));
    assert('slide84 has m:bar', slide84.includes('m:bar'));
    assert('slide84 has m:sSubSup', slide84.includes('m:sSubSup'));
  }

  // ── Slide 85: Blur effect ─────────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 85: blur effect');
    const slide85 = textFiles.get('ppt/slides/slide86.xml') || '';
    assert('slide85 exists', slide85.length > 0);
    assert('slide85 has a:effectLst', hasTag(slide85, 'a:effectLst'));
    assert('slide85 has a:blur', hasTag(slide85, 'a:blur'));
    assert('slide85 blur rad=76200', slide85.includes('rad="76200"'));
  }

  // ── Slide 86: Preset shadow ─────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 86: preset shadow');
    const slide86 = textFiles.get('ppt/slides/slide87.xml') || '';
    assert('slide86 exists', slide86.length > 0);
    assert('slide86 has a:effectLst', hasTag(slide86, 'a:effectLst'));
    assert('slide86 has a:prstShdw', hasTag(slide86, 'a:prstShdw'));
    assert('slide86 has prst=shdw1', slide86.includes('prst="shdw1"'));
    assert('slide86 has prst=shdw2', slide86.includes('prst="shdw2"'));
    assert('slide86 has dist=76200', slide86.includes('dist="76200"'));
  }

  // ── Slide 87: Fill overlay ─────────────────────────────────────────────────
  {
    section('test_features.pptx — Slide 87: fill overlay');
    const slide87 = textFiles.get('ppt/slides/slide88.xml') || '';
    assert('slide87 exists', slide87.length > 0);
    assert('slide87 has a:fillOverlay', hasTag(slide87, 'a:fillOverlay'));
    assert('slide87 has blend=over', slide87.includes('blend="over"'));
    assert('slide87 has blend=mult', slide87.includes('blend="mult"'));
    assert('slide87 has blend=screen', slide87.includes('blend="screen"'));
    assert('slide87 has blend=darken', slide87.includes('blend="darken"'));
    assert('slide87 has blend=lighten', slide87.includes('blend="lighten"'));
  }

  finishAssertions();
});
