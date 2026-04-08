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

// --- Test 19: Add shape ---
console.log('Test 19: Add shape');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Add a red rectangle
  const result = renderer.addShape(0, 'rect', 914400, 914400, 1828800, 914400, 255, 0, 0);
  assert(result.startsWith('OK:'), `addShape should return OK, got ${result}`);
  const shapeIdx = parseInt(result.split(':')[1]);
  assert(shapeIdx >= 0, 'shape index should be >= 0');

  // Slide should render with the new shape
  const svg = renderer.renderSlideSvg(0);
  assert(svg.includes('ff0000') || svg.includes('FF0000'), 'SVG should contain red fill');

  // Export and reload
  const exported = await renderer.exportPptx();
  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  await renderer2.loadPptx(exported);
  const svg2 = renderer2.renderSlideSvg(0);
  assert(svg2.includes('ff0000') || svg2.includes('FF0000'), 'exported slide should contain red shape');
  console.log(`  OK: Added shape at idx=${shapeIdx}, round-trip verified`);
}

// --- Test 20: Add shape with no fill ---
console.log('Test 20: Add shape with no fill');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  const result = renderer.addShape(0, 'ellipse', 0, 0, 914400, 914400);
  assert(result.startsWith('OK:'), `addShape (no fill) should return OK, got ${result}`);
  const svg = renderer.renderSlideSvg(0);
  assert(svg.startsWith('<svg'), 'slide should render');
  console.log('  OK: Added ellipse with no fill');
}

// --- Test 21: Delete shape ---
console.log('Test 21: Delete shape');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Render first to populate cache
  renderer.renderSlideSvg(0);

  // Delete shape 0
  const result = renderer.deleteShape(0, 0);
  assert(result === 'OK', `deleteShape should return OK, got ${result}`);

  // Should still render
  const svg = renderer.renderSlideSvg(0);
  assert(svg.startsWith('<svg'), 'slide should render after deletion');

  // Invalid index should error
  const err = renderer.deleteShape(0, 999);
  assert(err.startsWith('ERROR:'), 'deleting invalid index should error');
  console.log('  OK: Deleted shape');
}

// --- Test 22: Duplicate shape ---
console.log('Test 22: Duplicate shape');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Render first to parse slide data, then add a shape and duplicate it
  renderer.renderSlideSvg(0);
  const addResult = renderer.addShape(0, 'rect', 0, 0, 914400, 914400, 0, 0, 255);
  const addedIdx = parseInt(addResult.split(':')[1]);

  const result = renderer.duplicateShape(0, addedIdx);
  assert(result.startsWith('OK:'), `duplicateShape should return OK, got ${result}`);
  const newIdx = parseInt(result.split(':')[1]);

  const svg = renderer.renderSlideSvg(0);
  assert(svg.startsWith('<svg'), 'slide should render with duplicated shape');

  // Export and verify
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > 0, 'export should work');
  console.log(`  OK: Duplicated shape to idx=${newIdx}`);
}

// --- Test 23: Update gradient fill ---
console.log('Test 23: Update gradient fill');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Render first to parse slide data, then add a shape and apply gradient
  renderer.renderSlideSvg(0);
  const addResult = renderer.addShape(0, 'rect', 0, 0, 914400, 914400, 128, 128, 128);
  const addedIdx = parseInt(addResult.split(':')[1]);

  const stops = [
    { pos: 0, r: 255, g: 0, b: 0 },
    { pos: 100000, r: 0, g: 0, b: 255 },
  ];
  const svg = renderer.updateShapeGradientFill(0, addedIdx, 5400000, stops);
  assert(!svg.startsWith('ERROR:'), `gradient fill should succeed, got ${svg}`);
  assert(svg.includes('linearGradient') || svg.includes('radialGradient'),
    'SVG should contain gradient definition');
  console.log('  OK: Applied gradient fill');
}

// --- Test 24: Update stroke ---
console.log('Test 24: Update stroke');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  renderer.renderSlideSvg(0);
  const addResult = renderer.addShape(0, 'rect', 0, 0, 914400, 914400, 200, 200, 200);
  const addedIdx = parseInt(addResult.split(':')[1]);

  // Apply red stroke with dash
  const svg = renderer.updateShapeStroke(0, addedIdx, 255, 0, 0, 25400, 'dash');
  assert(!svg.startsWith('ERROR:'), `stroke update should succeed, got ${svg}`);
  assert(svg.includes('ff0000') || svg.includes('FF0000'), 'SVG should show red stroke');
  console.log('  OK: Applied stroke');
}

// --- Test 25: Remove stroke ---
console.log('Test 25: Remove stroke');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  renderer.renderSlideSvg(0);
  const addResult = renderer.addShape(0, 'rect', 0, 0, 914400, 914400, 200, 200, 200);
  const addedIdx = parseInt(addResult.split(':')[1]);

  // Set stroke, then remove it
  renderer.updateShapeStroke(0, addedIdx, 255, 0, 0, 25400);
  const svg = renderer.updateShapeStroke(0, addedIdx, -1, -1, -1, 0);
  assert(!svg.startsWith('ERROR:'), `remove stroke should succeed, got ${svg}`);
  console.log('  OK: Removed stroke');
}

// --- Test 26: Add line shape (should have default stroke) ---
console.log('Test 26: Add line shape with default stroke');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  const result = renderer.addShape(0, 'line', 914400, 2000000, 3657600, 0);
  assert(result.startsWith('OK:'), `addShape(line) should return OK, got ${result}`);

  // The line should be visible (default black 1pt stroke)
  const svg = renderer.renderSlideSvg(0);
  assert(svg.includes('stroke'), 'line shape SVG should contain stroke attribute');

  // Export and verify the line persists
  const exported = await renderer.exportPptx();
  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  await renderer2.loadPptx(exported);
  const svg2 = renderer2.renderSlideSvg(0);
  assert(svg2.includes('stroke'), 'exported line should have stroke');
  console.log('  OK: Line shape has default stroke');
}

// --- Test 27: Add text to shape ---
console.log('Test 27: Add text to shape');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // Add a shape, then add text to it
  renderer.renderSlideSvg(0);
  const addResult = renderer.addShape(0, 'rect', 914400, 914400, 3657600, 1828800, 200, 200, 200);
  const shapeIdx = parseInt(addResult.split(':')[1]);

  const textResult = renderer.addShapeText(0, shapeIdx, 'Hello World', 1800);
  assert(textResult.startsWith('OK:'), `addShapeText should return OK, got ${textResult}`);
  const paraIdx = parseInt(textResult.split(':')[1]);
  assert(paraIdx === 0, `first paragraph index should be 0, got ${paraIdx}`);

  // Render and verify text appears in SVG (word-wrapped into separate tspans)
  const svg = renderer.renderSlideSvg(0);
  assert(svg.includes('>Hello ') || svg.includes('>Hello<'), 'SVG should contain the added text');

  // Add a second paragraph
  const textResult2 = renderer.addShapeText(0, shapeIdx, 'SecondLine', 1400, 255, 0, 0);
  assert(textResult2.startsWith('OK:'), `second addShapeText should return OK, got ${textResult2}`);
  const paraIdx2 = parseInt(textResult2.split(':')[1]);
  assert(paraIdx2 === 1, `second paragraph index should be 1, got ${paraIdx2}`);

  const svg2 = renderer.renderSlideSvg(0);
  assert(svg2.includes('SecondLine'), 'SVG should contain the second paragraph text');

  // Export and reload to verify round-trip
  const exported = await renderer.exportPptx();
  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  await renderer2.loadPptx(exported);
  const svg3 = renderer2.renderSlideSvg(0);
  assert(svg3.includes('>Hello ') || svg3.includes('>Hello<'), 'exported slide should contain text');
  console.log(`  OK: Added text paragraphs to shape, round-trip verified`);
}

// --- Test 28: Add text to invalid shape ---
console.log('Test 28: Add text to invalid shape');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  renderer.renderSlideSvg(0);

  const result = renderer.addShapeText(0, 999, 'text', 1800);
  assert(result.startsWith('ERROR:'), `invalid shape index should error, got ${result}`);

  const result2 = renderer.addShapeText(99, 0, 'text', 1800);
  assert(result2.startsWith('ERROR:'), `invalid slide index should error, got ${result2}`);
  console.log('  OK: Error handling for invalid indices');
}

// --- Test 29: Text editing E2.5 — paragraph and run CRUD ---
console.log('Test 29: Text editing E2.5 — paragraph and run CRUD');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  // Add a shape to work with
  const addResult = renderer.addShape(0, 'rect', 914400, 914400, 3657600, 1828800, 200, 200, 200);
  const shapeIdx = parseInt(addResult.split(':')[1]);

  // addParagraph
  const p0 = renderer.addParagraph(0, shapeIdx, 'First paragraph', 'ctr');
  assert(p0.startsWith('OK:'), `addParagraph should return OK, got ${p0}`);
  assert(p0 === 'OK:0', `first paragraph index should be 0, got ${p0}`);

  const p1 = renderer.addParagraph(0, shapeIdx, 'Second paragraph', 'r');
  assert(p1 === 'OK:1', `second paragraph index should be 1, got ${p1}`);

  // addRun
  const r0 = renderer.addRun(0, shapeIdx, 0, ' extra run');
  assert(r0.startsWith('OK:'), `addRun should return OK, got ${r0}`);
  assert(r0 === 'OK:1', `second run index should be 1, got ${r0}`);

  // deleteRun
  const dr = renderer.deleteRun(0, shapeIdx, 0, 1);
  assert(dr === 'OK', `deleteRun should return OK, got ${dr}`);

  // deleteParagraph
  const dp = renderer.deleteParagraph(0, shapeIdx, 1);
  assert(dp === 'OK', `deleteParagraph should return OK, got ${dp}`);

  // Error cases
  assert(renderer.addParagraph(0, 999, 'x', '').startsWith('ERROR:'), 'addParagraph invalid shape');
  assert(renderer.addRun(0, shapeIdx, 99, 'x').startsWith('ERROR:'), 'addRun invalid para');
  assert(renderer.deleteRun(0, shapeIdx, 0, 99).startsWith('ERROR:'), 'deleteRun invalid run');
  assert(renderer.deleteParagraph(0, shapeIdx, 99).startsWith('ERROR:'), 'deleteParagraph invalid para');

  console.log('  OK: Paragraph and run CRUD');
}

// --- Test 30: Text editing E2.5 — style, font size, color ---
console.log('Test 30: Text editing E2.5 — style, font size, color');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  const addResult = renderer.addShape(0, 'rect', 0, 0, 3657600, 1828800, 240, 240, 240);
  const shapeIdx = parseInt(addResult.split(':')[1]);
  renderer.addParagraph(0, shapeIdx, 'Styled text', '');

  // updateTextRunStyle (bold/italic)
  const boldSvg = renderer.updateTextRunStyle(0, shapeIdx, 0, 0, 1, -1);
  assert(!boldSvg.startsWith('ERROR:'), `updateTextRunStyle should not error, got ${boldSvg.slice(0,60)}`);
  assert(boldSvg.includes('font-weight="bold"') || boldSvg.includes('data-ooxml-bold="true"'),
    'SVG should reflect bold');

  const italicSvg = renderer.updateTextRunStyle(0, shapeIdx, 0, 0, -1, 1);
  assert(!italicSvg.startsWith('ERROR:'), 'updateTextRunStyle italic should not error');
  assert(italicSvg.includes('font-style="italic"'), 'SVG should reflect italic');

  // updateTextRunFontSize
  const sizeSvg = renderer.updateTextRunFontSize(0, shapeIdx, 0, 0, 3600);
  assert(!sizeSvg.startsWith('ERROR:'), `updateTextRunFontSize should not error`);

  // updateTextRunColor
  const colorSvg = renderer.updateTextRunColor(0, shapeIdx, 0, 0, 255, 0, 0);
  assert(!colorSvg.startsWith('ERROR:'), `updateTextRunColor should not error`);
  assert(colorSvg.includes('#ff0000') || colorSvg.includes('rgb(255'), 'SVG should contain red color');

  // Error cases
  assert(renderer.updateTextRunStyle(0, shapeIdx, 0, 99, 1, 0).startsWith('ERROR:'), 'invalid run idx');
  assert(renderer.updateTextRunFontSize(0, shapeIdx, 99, 0, 1800).startsWith('ERROR:'), 'invalid para idx');

  console.log('  OK: Style, font size, color updates');
}

// --- Test 31: Text editing E2.5 — font family, alignment, decoration ---
console.log('Test 31: Text editing E2.5 — font family, alignment, decoration');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  const addResult = renderer.addShape(0, 'rect', 0, 0, 3657600, 1828800, 240, 240, 240);
  const shapeIdx = parseInt(addResult.split(':')[1]);
  renderer.addParagraph(0, shapeIdx, 'Decorated text', 'l');

  // updateTextRunFont
  const fontSvg = renderer.updateTextRunFont(0, shapeIdx, 0, 0, 'Arial', 'MS Gothic', '');
  assert(!fontSvg.startsWith('ERROR:'), `updateTextRunFont should not error`);
  assert(fontSvg.includes('Arial'), 'SVG should contain font name');

  // updateParagraphAlign
  const alignSvg = renderer.updateParagraphAlign(0, shapeIdx, 0, 'ctr');
  assert(!alignSvg.startsWith('ERROR:'), `updateParagraphAlign should not error`);
  assert(alignSvg.includes('data-ooxml-para-align="ctr"'), 'SVG should have center alignment');

  // updateTextRunDecoration — underline
  const ulSvg = renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, 'sng', '', -1);
  assert(!ulSvg.startsWith('ERROR:'), `updateTextRunDecoration should not error`);
  assert(ulSvg.includes('text-decoration') || ulSvg.includes('data-ooxml-underline="sng"'),
    'SVG should reflect underline');

  // updateTextRunDecoration — strikethrough
  const stSvg = renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, '', 'sngStrike', -1);
  assert(!stSvg.startsWith('ERROR:'), 'strikethrough should not error');

  // updateTextRunDecoration — superscript
  const supSvg = renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, '', '', 30000);
  assert(!supSvg.startsWith('ERROR:'), 'superscript should not error');

  // updateTextRunDecoration — remove underline
  const noUlSvg = renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, 'none', 'none', 0);
  assert(!noUlSvg.startsWith('ERROR:'), 'remove decoration should not error');

  // Error cases
  assert(renderer.updateTextRunFont(0, shapeIdx, 0, 99, 'Arial', '', '').startsWith('ERROR:'), 'invalid run');
  assert(renderer.updateParagraphAlign(0, shapeIdx, 99, 'l').startsWith('ERROR:'), 'invalid para');
  assert(renderer.updateTextRunDecoration(0, shapeIdx, 0, 99, 'sng', '', -1).startsWith('ERROR:'), 'invalid run');

  console.log('  OK: Font family, alignment, decoration updates');
}

// --- Test 32: Text editing E2.5 — round-trip export ---
console.log('Test 32: Text editing E2.5 — round-trip export');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  // Add shape + styled text
  const addResult = renderer.addShape(0, 'rect', 914400, 914400, 3657600, 1828800, 200, 200, 200);
  const shapeIdx = parseInt(addResult.split(':')[1]);
  renderer.addParagraph(0, shapeIdx, 'Bold Red', 'ctr');
  renderer.updateTextRunStyle(0, shapeIdx, 0, 0, 1, 1);
  renderer.updateTextRunColor(0, shapeIdx, 0, 0, 255, 0, 0);
  renderer.updateTextRunFontSize(0, shapeIdx, 0, 0, 2400);
  renderer.updateTextRunFont(0, shapeIdx, 0, 0, 'Impact', '', '');

  // Export and reload
  const exported = await renderer.exportPptx();
  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  await renderer2.loadPptx(exported);
  const svg = renderer2.renderSlideSvg(0);

  // Text may be split by word wrapping, so check for fragments
  assert(svg.includes('Bold') && svg.includes('Red'), 'exported text should survive round-trip');
  assert(svg.includes('font-weight="bold"'), 'bold should survive round-trip');
  assert(svg.includes('rgb(255,0,0)') || svg.includes('#ff0000') || svg.includes('ff0000'), 'red color should survive round-trip');
  assert(svg.includes('Impact'), 'font should survive round-trip');
  assert(svg.includes('data-ooxml-para-align="ctr"'), 'alignment should survive round-trip');

  console.log('  OK: Text editing round-trip export verified');
}

// --- Test 33: Image API — addImage ---
console.log('Test 33: Image API — addImage');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  // Minimal 1x1 red PNG (valid PNG file)
  const pngData = new Uint8Array([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x36, 0x28, 0x19,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
    0x44, 0xAE, 0x42, 0x60, 0x82,
  ]);

  const result = renderer.addImage(0, pngData, 'image/png',
    914400, 914400, 1828800, 1828800);
  assert(result.startsWith('OK:'), `addImage should return OK, got ${result}`);
  const shapeIdx = parseInt(result.split(':')[1]);
  assert(shapeIdx >= 0, `shape index should be non-negative, got ${shapeIdx}`);

  // Render should include the picture shape
  const svg = renderer.renderSlideSvg(0);
  assert(svg.includes('data-ooxml-shape-type="picture"'), 'SVG should contain picture shape');

  // Export and verify the image file exists
  const exported = await renderer.exportPptx();
  assert(exported.byteLength > pptxAb.byteLength, 'exported PPTX should be larger with image');

  // Reload and verify
  const renderer2 = new PptxRenderer({ logLevel: 'silent' });
  await renderer2.init(wasmBuf);
  await renderer2.loadPptx(exported);
  const svg2 = renderer2.renderSlideSvg(0);
  assert(svg2.includes('data-ooxml-shape-type="picture"'), 'reloaded SVG should contain picture shape');

  console.log('  OK: addImage with round-trip export');
}

// --- Test 34: Image API — replaceImage ---
console.log('Test 34: Image API — replaceImage');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  // Add an image first
  const pngData = new Uint8Array([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x36, 0x28, 0x19,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
  ]);

  const addResult = renderer.addImage(0, pngData, 'image/png',
    914400, 914400, 1828800, 1828800);
  const shapeIdx = parseInt(addResult.split(':')[1]);

  // Replace with a different (same-format) image
  const pngData2 = new Uint8Array([...pngData]); // same structure, different identity
  pngData2[pngData2.length - 5] = 0x01; // slightly different
  const replaceResult = renderer.replaceImage(0, shapeIdx, pngData2, 'image/png');
  assert(replaceResult === 'OK', `replaceImage should return OK, got ${replaceResult}`);

  // Error cases
  assert(renderer.replaceImage(0, 999, pngData, 'image/png').startsWith('ERROR:'), 'invalid shape');
  assert(renderer.addImage(0, pngData, 'image/bla', 0, 0, 100, 100).startsWith('ERROR:'), 'invalid mime');

  console.log('  OK: replaceImage');
}

// --- Test 35: Image API — deleteImage ---
console.log('Test 35: Image API — deleteImage');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  renderer.renderSlideSvg(0);

  // Add then delete
  const pngData = new Uint8Array([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x36, 0x28, 0x19,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
  ]);

  const addResult = renderer.addImage(0, pngData, 'image/png',
    914400, 914400, 1828800, 1828800);
  const shapeIdx = parseInt(addResult.split(':')[1]);

  const svg1 = renderer.renderSlideSvg(0);
  assert(svg1.includes('data-ooxml-shape-type="picture"'), 'should have picture before delete');

  const delResult = renderer.deleteImage(0, shapeIdx);
  assert(delResult === 'OK', `deleteImage should return OK, got ${delResult}`);

  const svg2 = renderer.renderSlideSvg(0);
  assert(!svg2.includes(`data-ooxml-shape-idx="${shapeIdx}"`), 'deleted shape should not be in SVG');

  // Error case
  assert(renderer.deleteImage(0, 999).startsWith('ERROR:'), 'delete invalid shape');

  console.log('  OK: deleteImage');
}

// --- Summary ---
console.log('');
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
} else {
  console.log('All Node.js compatibility tests passed!');
}
