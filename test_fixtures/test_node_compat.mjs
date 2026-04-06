/**
 * Node.js compatibility test — verifies that pptx-svg works on Node.js 22+.
 *
 * Run: node test_fixtures/test_node_compat.mjs
 * Requires: Node.js 22+ (WasmGC support)
 */

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const { PptxRenderer } = await import(join(__dirname, '..', 'dist', 'index.js'));

let passed = 0;
let failed = 0;

function assert(condition, message) {
  if (condition) {
    passed++;
  } else {
    failed++;
    console.error(`  FAIL: ${message}`);
  }
}

console.log('=== Node.js Compatibility Test ===');
console.log(`Node.js ${process.version}, V8 ${process.versions.v8}`);
console.log('');

// --- Test 1: Wasm initialization with ArrayBuffer ---
console.log('Test 1: Wasm init with ArrayBuffer');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmPath = join(__dirname, '..', 'dist', 'main.wasm');
  const wasmBuf = readFileSync(wasmPath);
  // Buffer.buffer may be a shared pool — slice to get standalone ArrayBuffer
  const ab = wasmBuf.buffer.slice(wasmBuf.byteOffset, wasmBuf.byteOffset + wasmBuf.byteLength);
  await renderer.init(ab);
  assert(true, 'init with ArrayBuffer');
  console.log('  OK: Wasm initialized from ArrayBuffer');
}

// --- Test 2: Wasm initialization with Uint8Array (Buffer) ---
console.log('Test 2: Wasm init with Uint8Array/Buffer');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmPath = join(__dirname, '..', 'dist', 'main.wasm');
  const wasmBuf = readFileSync(wasmPath); // Node Buffer extends Uint8Array
  await renderer.init(wasmBuf);
  assert(true, 'init with Uint8Array');
  console.log('  OK: Wasm initialized from Uint8Array/Buffer');
}

// --- Test 3: Load and render PPTX ---
console.log('Test 3: Load and render PPTX');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  const { slideCount } = await renderer.loadPptx(pptxAb);
  assert(slideCount >= 1, `slideCount should be >= 1, got ${slideCount}`);
  console.log(`  OK: Loaded PPTX with ${slideCount} slides`);

  const svg = renderer.renderSlideSvg(0);
  assert(svg.startsWith('<svg'), 'SVG output should start with <svg');
  assert(svg.length > 100, `SVG should have content, got ${svg.length} chars`);
  console.log(`  OK: Rendered slide 0 as SVG (${svg.length} chars)`);
}

// --- Test 4: Custom measureText function ---
console.log('Test 4: Custom measureText');
{
  let measureCalled = false;
  const renderer = new PptxRenderer({
    logLevel: 'silent',
    measureText: (text, font, size) => {
      measureCalled = true;
      return text.length * size * 0.6;
    },
  });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);
  assert(measureCalled, 'custom measureText should have been called');
  console.log(`  OK: Custom measureText was invoked`);
}

// --- Test 5: Export PPTX ---
console.log('Test 5: Export PPTX');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  const exported = await renderer.exportPptx();
  assert(exported instanceof ArrayBuffer, 'exported should be ArrayBuffer');
  assert(exported.byteLength > 0, 'exported should have content');
  console.log(`  OK: Exported PPTX (${exported.byteLength} bytes)`);
}

// --- Test 6: Add slide ---
console.log('Test 6: Add slide');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  assert(renderer.getSlideCount() === 2, 'initial slide count should be 2');

  const { slideCount, insertedIdx } = await renderer.addSlide();
  assert(slideCount === 3, `after addSlide, count should be 3, got ${slideCount}`);
  assert(insertedIdx === 2, `insertedIdx should be 2, got ${insertedIdx}`);
  assert(renderer.getSlideCount() === 3, 'getSlideCount should return 3');

  // The new slide should render without error
  const svg = renderer.renderSlideSvg(2);
  assert(svg.startsWith('<svg'), 'new slide SVG should start with <svg');

  // Export should produce valid PPTX
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > 0, 'exported PPTX should have content');
  console.log(`  OK: Added slide (count=${slideCount}, inserted at ${insertedIdx})`);
}

// --- Test 7: Add slide at beginning ---
console.log('Test 7: Add slide at beginning');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  const { slideCount, insertedIdx } = await renderer.addSlide(-1);
  assert(slideCount === 3, `count should be 3, got ${slideCount}`);
  assert(insertedIdx === 0, `insertedIdx should be 0, got ${insertedIdx}`);

  // Original slide 1 should now be at index 1
  const svg1 = renderer.renderSlideSvg(1);
  assert(svg1.startsWith('<svg'), 'shifted slide should render');
  console.log(`  OK: Added slide at beginning (count=${slideCount})`);
}

// --- Test 8: Delete slide ---
console.log('Test 8: Delete slide');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  assert(renderer.getSlideCount() === 2, 'initial count should be 2');

  const { slideCount } = await renderer.deleteSlide(0);
  assert(slideCount === 1, `after delete, count should be 1, got ${slideCount}`);
  assert(renderer.getSlideCount() === 1, 'getSlideCount should return 1');

  // Remaining slide should render
  const svg = renderer.renderSlideSvg(0);
  assert(svg.startsWith('<svg'), 'remaining slide should render');

  // Export should work
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > 0, 'exported PPTX should have content');
  console.log(`  OK: Deleted slide (count=${slideCount})`);
}

// --- Test 9: Delete last slide should fail ---
console.log('Test 9: Delete last slide should fail');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Delete first, then try to delete the last remaining
  await renderer.deleteSlide(0);
  let threw = false;
  try {
    await renderer.deleteSlide(0);
  } catch (e) {
    threw = true;
  }
  assert(threw, 'deleting last slide should throw');
  console.log('  OK: Correctly prevented deleting last slide');
}

// --- Test 10: Reorder slides ---
console.log('Test 10: Reorder slides');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Reverse order: [1, 0]
  const { slideCount } = await renderer.reorderSlides([1, 0]);
  assert(slideCount === 2, `count should stay 2, got ${slideCount}`);

  // Both slides should still render
  const svg0 = renderer.renderSlideSvg(0);
  const svg1 = renderer.renderSlideSvg(1);
  assert(svg0.startsWith('<svg'), 'reordered slide 0 should render');
  assert(svg1.startsWith('<svg'), 'reordered slide 1 should render');

  // Export should work
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > 0, 'exported PPTX should have content');
  console.log(`  OK: Reordered slides (count=${slideCount})`);
}

// --- Test 11: Reorder with invalid permutation should fail ---
console.log('Test 11: Invalid reorder should fail');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  let threw = false;
  try {
    await renderer.reorderSlides([0, 0]); // not a valid permutation
  } catch (e) {
    threw = true;
  }
  assert(threw, 'invalid permutation should throw');
  console.log('  OK: Correctly rejected invalid permutation');
}

// --- Test 12: Add then delete round-trip ---
console.log('Test 12: Add then delete round-trip');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Add a slide, then delete it
  await renderer.addSlide();
  assert(renderer.getSlideCount() === 3, 'count should be 3 after add');
  await renderer.deleteSlide(2);
  assert(renderer.getSlideCount() === 2, 'count should be 2 after delete');

  // Both original slides should still render
  const svg0 = renderer.renderSlideSvg(0);
  const svg1 = renderer.renderSlideSvg(1);
  assert(svg0.startsWith('<svg'), 'slide 0 should render after round-trip');
  assert(svg1.startsWith('<svg'), 'slide 1 should render after round-trip');
  console.log('  OK: Add then delete round-trip works');
}

// --- Test 13: Add slide in middle position ---
console.log('Test 13: Add slide in middle position');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // First add a third slide at end, then add a fourth between positions 0 and 1
  await renderer.addSlide();         // [0, 1, new2]
  assert(renderer.getSlideCount() === 3, 'count should be 3');
  const { insertedIdx } = await renderer.addSlide(0);  // [0, new, 1, 2]
  assert(insertedIdx === 1, `insertedIdx should be 1, got ${insertedIdx}`);
  assert(renderer.getSlideCount() === 4, 'count should be 4');

  // All 4 slides should render
  for (let i = 0; i < 4; i++) {
    const svg = renderer.renderSlideSvg(i);
    assert(svg.startsWith('<svg'), `slide ${i} should render`);
  }
  console.log('  OK: Added slide in middle position');
}

// --- Test 14: Delete middle slide with 3+ slides ---
console.log('Test 14: Delete middle slide');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  await renderer.addSlide(); // now 3 slides
  await renderer.deleteSlide(1); // delete middle
  assert(renderer.getSlideCount() === 2, 'count should be 2 after deleting middle');

  const svg0 = renderer.renderSlideSvg(0);
  const svg1 = renderer.renderSlideSvg(1);
  assert(svg0.startsWith('<svg'), 'slide 0 should render');
  assert(svg1.startsWith('<svg'), 'slide 1 should render');
  console.log('  OK: Deleted middle slide');
}

// --- Test 15: Reorder with 3 slides ---
console.log('Test 15: Reorder with 3 slides');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  await renderer.addSlide(); // [0, 1, 2]
  // Rotate: [1, 2, 0]
  const { slideCount } = await renderer.reorderSlides([1, 2, 0]);
  assert(slideCount === 3, 'count should stay 3');

  for (let i = 0; i < 3; i++) {
    const svg = renderer.renderSlideSvg(i);
    assert(svg.startsWith('<svg'), `slide ${i} should render after rotate`);
  }
  console.log('  OK: Reordered 3 slides');
}

// --- Test 16: Add multiple slides in sequence ---
console.log('Test 16: Add multiple slides sequentially');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  await renderer.addSlide();
  await renderer.addSlide();
  await renderer.addSlide();
  assert(renderer.getSlideCount() === 5, 'count should be 5');

  for (let i = 0; i < 5; i++) {
    const svg = renderer.renderSlideSvg(i);
    assert(svg.startsWith('<svg'), `slide ${i} should render`);
  }

  // Export and reload to verify structural integrity
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > 0, 'export should produce bytes');

  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  const { slideCount } = await renderer2.loadPptx(exported);
  assert(slideCount === 5, `reloaded should have 5 slides, got ${slideCount}`);
  console.log(`  OK: Added 3 slides, exported, reloaded (${slideCount} slides)`);
}

// --- Test 17: Export after reorder then reload ---
console.log('Test 17: Export after reorder, reload');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  await renderer.reorderSlides([1, 0]);
  const exported = await renderer.exportPptx();

  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  const { slideCount } = await renderer2.loadPptx(exported);
  assert(slideCount === 2, `reloaded should have 2 slides, got ${slideCount}`);

  // Both should render
  const svg0 = renderer2.renderSlideSvg(0);
  const svg1 = renderer2.renderSlideSvg(1);
  assert(svg0.startsWith('<svg'), 'reloaded slide 0 should render');
  assert(svg1.startsWith('<svg'), 'reloaded slide 1 should render');
  console.log('  OK: Export after reorder, reloaded successfully');
}

// --- Test 18: Delete then add ---
console.log('Test 18: Delete then add');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  await renderer.deleteSlide(0);
  assert(renderer.getSlideCount() === 1, 'count should be 1 after delete');
  await renderer.addSlide();
  assert(renderer.getSlideCount() === 2, 'count should be 2 after add');

  const svg0 = renderer.renderSlideSvg(0);
  const svg1 = renderer.renderSlideSvg(1);
  assert(svg0.startsWith('<svg'), 'slide 0 should render');
  assert(svg1.startsWith('<svg'), 'slide 1 should render');
  console.log('  OK: Delete then add works');
}

// --- Summary ---
console.log('');
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
} else {
  console.log('All Node.js compatibility tests passed!');
}
