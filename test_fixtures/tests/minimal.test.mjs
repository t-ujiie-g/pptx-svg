import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';
import { loadPptxEntries } from './_helpers.mjs';

test('minimal.pptx — basic ZIP extraction', async () => {
  resetAssertions();
  const { textFiles } = await loadPptxEntries('minimal.pptx');

  const prsXml = textFiles.get('ppt/presentation.xml');
  expect('presentation.xml exists', !!prsXml);

  const slide1 = textFiles.get('ppt/slides/slide1.xml');
  expect('slide1.xml exists', !!slide1);
  expect('slide1.xml contains title text', slide1?.includes('Hello from MoonBit') ?? false);

  const slide2 = textFiles.get('ppt/slides/slide2.xml');
  expect('slide2.xml exists', !!slide2);

  finishAssertions();
});
