import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("tables (slides 41-42)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

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

  finishAssertions();
});
