import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { assert, section, resetAssertions, finishAssertions, DIST_DIR, FIXTURES_DIR } from './_helpers.mjs';

async function renderAll(logLevel = 'silent') {
  const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
  const r = new PptxRenderer({ logLevel });
  await r.init(readFileSync(join(DIST_DIR, 'main.wasm')));
  const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
  const { slideCount } = await r.loadPptx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  const svgs = [];
  for (let i = 0; i < slideCount; i++) svgs.push(r.renderSlideSvg(i));
  return svgs;
}

test('Wasm loads without js-string builtins (Safari 18.2–26.1, older Chrome/Firefox)', async () => {
  resetAssertions();
  section('js-string fallback');
  const native = await renderAll();

  // An engine without the builtins ignores the options argument, exactly like this.
  const orig = WebAssembly.instantiate;
  WebAssembly.instantiate = (bytes, imports) => orig.call(WebAssembly, bytes, imports);
  const logs = [];
  const origInfo = console.info, origLog = console.log;
  console.info = console.log = (...a) => logs.push(a.join(' '));
  let fallback = null, error = '';
  try {
    fallback = await renderAll('info');
  } catch (e) {
    error = String(e);
  } finally {
    WebAssembly.instantiate = orig;
    console.info = origInfo;
    console.log = origLog;
  }

  assert('fallback instantiation succeeds', error === '', error);
  assert('fallback path was taken', logs.some((l) => l.includes('js-string in JS')), logs.join(' | '));
  assert('fallback renders every slide', fallback?.length === native.length, `${fallback?.length} vs ${native.length}`);
  assert('fallback SVG is identical to the builtins path',
    fallback?.every((svg, i) => svg === native[i]) === true);
  finishAssertions();
});
