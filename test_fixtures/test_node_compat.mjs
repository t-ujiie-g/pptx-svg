/**
 * Node.js compatibility test — verifies that pptx-svg works on Node.js 22+.
 *
 * Run: node test_fixtures/test_node_compat.mjs
 * Requires: Node.js 22+ (WasmGC support)
 */

import { readFileSync, existsSync } from 'node:fs';
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

// --- Test 36: Edit immediately after loadPptx (no prior renderSlideSvg) ---
// Regression: editing APIs read the g_slides cache, which is populated lazily on
// the first render. Before the ensure_slide_parsed guard, an edit issued before any
// renderSlideSvg() silently no-op'd ("ERROR:shape index out of range") and never
// reached the export. minimal.pptx slide 0 shape 0 is a title with inline text.
console.log('Test 36: Edit immediately after loadPptx (no prior render)');
{
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);

  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);

  // No renderSlideSvg(0) here — edit straight after load.
  const ret = renderer.updateShapeText(0, 0, 0, 0, 'EDITED_NO_PRERENDER');
  assert(ret.startsWith('<g'), `updateShapeText should return a fragment, got: ${ret.slice(0, 60)}`);

  const ooxml = renderer.getSlideOoxml(0);
  assert(ooxml.includes('EDITED_NO_PRERENDER'), 'edit should be reflected in slide OOXML');

  // And it must survive a full export round-trip.
  const out = await renderer.exportPptx();
  assert(out.byteLength > 0, 'exportPptx should produce bytes');
  console.log('  OK: edit without pre-render persists to OOXML and export');
}

// ── Undo / Redo (E6.1) ───────────────────────────────────────────────────────
// Shared helper: a freshly loaded renderer on minimal.pptx.
async function freshRenderer(opts = { logLevel: 'silent' }) {
  const renderer = new PptxRenderer(opts);
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const pptxBuf = readFileSync(join(__dirname, 'minimal.pptx'));
  const pptxAb = pptxBuf.buffer.slice(pptxBuf.byteOffset, pptxBuf.byteOffset + pptxBuf.byteLength);
  await renderer.loadPptx(pptxAb);
  return renderer;
}

// --- Test 37: Undo/redo a transform edit restores exact SVG ---
console.log('Test 37: Undo/redo transform');
{
  const r = await freshRenderer();
  const before = r.renderSlideSvg(0);
  assert(r.canUndo() === false, 'fresh load should have nothing to undo');

  r.updateShapeTransform(0, 0, 1234567, 2345678, 1828800, 914400, 0);
  const after = r.renderSlideSvg(0);
  assert(after !== before, 'transform edit should change the SVG');
  assert(r.canUndo() === true, 'should be able to undo after an edit');

  const u = r.undo();
  assert(!u.startsWith('ERROR'), `undo should succeed, got ${u}`);
  const parsed = JSON.parse(u);
  assert(Array.isArray(parsed.slides), 'undo result should carry a slides array');
  assert(parsed.slides.includes(0), 'undo result should flag slide 0');
  assert(r.renderSlideSvg(0) === before, 'undo should restore the original SVG exactly');
  assert(r.canRedo() === true, 'should be able to redo after undo');

  const re = r.redo();
  assert(!re.startsWith('ERROR'), `redo should succeed, got ${re}`);
  assert(r.renderSlideSvg(0) === after, 'redo should re-apply the edit exactly');
  console.log('  OK: transform undo/redo restores SVG');
}

// --- Test 38: Undo an addShape removes the shape; redo restores it ---
console.log('Test 38: Undo/redo addShape');
{
  const r = await freshRenderer();
  const before = r.renderSlideSvg(0);
  r.addShape(0, 'rect', 914400, 914400, 1828800, 914400, 255, 0, 0);
  const withShape = r.renderSlideSvg(0);
  assert(withShape.includes('ff0000') || withShape.includes('FF0000'), 'added shape should be red');

  r.undo();
  const afterUndo = r.renderSlideSvg(0);
  assert(!(afterUndo.includes('ff0000') || afterUndo.includes('FF0000')), 'undo should remove the red shape');
  assert(afterUndo === before, 'undo should restore the original SVG');

  r.redo();
  const afterRedo = r.renderSlideSvg(0);
  assert(afterRedo.includes('ff0000') || afterRedo.includes('FF0000'), 'redo should bring the red shape back');
  console.log('  OK: addShape undo/redo');
}

// --- Test 39: Undo a deleteShape restores the deleted shape ---
console.log('Test 39: Undo deleteShape');
{
  const r = await freshRenderer();
  const before = r.renderSlideSvg(0);
  const del = r.deleteShape(0, 0);
  assert(del === 'OK', `deleteShape should return OK, got ${del}`);
  const afterDel = r.renderSlideSvg(0);
  assert(afterDel !== before, 'delete should change the SVG');

  r.undo();
  assert(r.renderSlideSvg(0) === before, 'undo should restore the deleted shape');
  console.log('  OK: deleteShape undo');
}

// --- Test 40: Undo/redo addSlide changes slide count ---
console.log('Test 40: Undo/redo addSlide');
{
  const r = await freshRenderer();
  const n = r.getSlideCount();
  await r.addSlide();
  assert(r.getSlideCount() === n + 1, 'addSlide should increase the count');

  const u = JSON.parse(r.undo());
  assert(r.getSlideCount() === n, 'undo addSlide should restore the count');
  assert(u.slideCount === n, `undo result slideCount should be ${n}, got ${u.slideCount}`);

  r.redo();
  assert(r.getSlideCount() === n + 1, 'redo addSlide should re-add the slide');
  console.log('  OK: addSlide undo/redo');
}

// --- Test 41: beginBatch/endBatch collapses multiple edits into one undo ---
console.log('Test 41: Batch = single undo step');
{
  const r = await freshRenderer();
  const before = r.renderSlideSvg(0);

  r.beginBatch();
  r.addShape(0, 'rect', 0, 0, 914400, 914400, 255, 0, 0);
  r.addShape(0, 'ellipse', 914400, 0, 914400, 914400, 0, 255, 0);
  r.updateShapeTransform(0, 0, 100000, 100000, 500000, 500000, 0);
  r.endBatch();

  const after = r.renderSlideSvg(0);
  assert(after !== before, 'batch edits should change the SVG');

  const u = r.undo();
  assert(!u.startsWith('ERROR'), 'single undo should revert the whole batch');
  assert(r.renderSlideSvg(0) === before, 'one undo should revert all batch edits');
  assert(r.canUndo() === false, 'batch should be a single undo entry');
  console.log('  OK: batch collapses to one undo');
}

// --- Test 42: empty history + clearHistory ---
console.log('Test 42: Empty history and clearHistory');
{
  const r = await freshRenderer();
  assert(r.undo() === 'ERROR:nothing to undo', 'undo on empty history should error');
  assert(r.redo() === 'ERROR:nothing to redo', 'redo on empty history should error');

  r.updateShapeFill(0, 0, 10, 20, 30);
  assert(r.canUndo() === true, 'edit should create an undo entry');
  r.clearHistory();
  assert(r.canUndo() === false, 'clearHistory should drop undo entries');
  assert(r.canRedo() === false, 'clearHistory should drop redo entries');
  assert(r.undo() === 'ERROR:nothing to undo', 'undo after clearHistory should error');
  console.log('  OK: empty history and clearHistory');
}

// --- Test 43: maxHistory caps the undo depth; export survives undo/redo ---
console.log('Test 43: maxHistory cap + export after undo');
{
  const r = await freshRenderer({ logLevel: 'silent', maxHistory: 2 });
  r.updateShapeFill(0, 0, 1, 1, 1);
  r.updateShapeFill(0, 0, 2, 2, 2);
  r.updateShapeFill(0, 0, 3, 3, 3);
  // Only the 2 most recent pre-edit states are retained.
  assert(r.undo().startsWith('{'), 'first undo should succeed');
  assert(r.undo().startsWith('{'), 'second undo should succeed');
  assert(r.undo() === 'ERROR:nothing to undo', 'third undo should be capped out');

  // Export still works after undo/redo churn.
  const out = await r.exportPptx();
  assert(out.byteLength > 0, 'exportPptx should produce bytes after undo/redo');
  console.log('  OK: maxHistory cap and export-after-undo');
}

// ── Inline text editing (E6.2) ───────────────────────────────────────────────
// Build a shape with two runs: "Hello " (normal) + "World" (bold), in one paragraph.
async function shapeWithTwoRuns() {
  const r = await freshRenderer();
  const sIdx = parseInt(r.addShape(0, 'rect', 914400, 914400, 5486400, 1828800, -1, -1, -1).split(':')[1]);
  r.addShapeText(0, sIdx, 'Hello ', 1800, 0, 0, 0);   // para 0, run 0
  r.addRun(0, sIdx, 0, 'World');                       // para 0, run 1
  r.updateTextRunStyle(0, sIdx, 0, 1, 1, -1);          // run 1 → bold
  return { r, sIdx };
}

// --- Test 44: getTextLayout returns structured geometry ---
console.log('Test 44: getTextLayout');
{
  const { r, sIdx } = await shapeWithTwoRuns();
  const layout = JSON.parse(r.getTextLayout(0, sIdx));
  assert(layout.box && typeof layout.box.cx === 'number', 'layout should carry a box in EMU');
  assert(Array.isArray(layout.lines) && layout.lines.length >= 1, 'layout should have at least one line');
  const totalRuns = layout.lines.reduce((n, l) => n + l.runs.length, 0);
  assert(totalRuns === 2, `expected 2 run boxes, got ${totalRuns}`);
  const totalChars = layout.lines.reduce((n, l) => n + l.runs.reduce((m, rb) => m + rb.chars.length, 0), 0);
  assert(totalChars === 'Hello World'.length, `expected 11 glyphs, got ${totalChars}`);
  // Char x positions are monotonically increasing within the first run.
  const firstRun = layout.lines[0].runs[0];
  let monotonic = true;
  for (let i = 1; i < firstRun.chars.length; i++) {
    if (firstRun.chars[i].x < firstRun.chars[i - 1].x) monotonic = false;
  }
  assert(monotonic, 'char x positions should be monotonically increasing');
  console.log('  OK: getTextLayout geometry');
}

// --- Test 45: hitTestText maps a point to a caret position ---
console.log('Test 45: hitTestText');
{
  const { r, sIdx } = await shapeWithTwoRuns();
  const layout = JSON.parse(r.getTextLayout(0, sIdx));
  const line = layout.lines[0];
  // Click near the start of the text.
  const hit = JSON.parse(r.hitTestText(0, sIdx, layout.box.x + 100, line.y + line.h / 2));
  assert(hit.paraIdx === 0, `expected paraIdx 0, got ${hit.paraIdx}`);
  assert(typeof hit.charOffset === 'number' && typeof hit.paraOffset === 'number', 'hit should carry offsets');
  // Click far to the right → caret near the paragraph end.
  const hitEnd = JSON.parse(r.hitTestText(0, sIdx, layout.box.x + layout.box.cx, line.y + line.h / 2));
  assert(hitEnd.paraOffset >= hit.paraOffset, 'rightward click should not move caret left');
  console.log('  OK: hitTestText');
}

// --- Test 46: replaceTextRange inserts mid-run, preserving boundary formatting ---
console.log('Test 46: replaceTextRange insert preserves format');
{
  const { r, sIdx } = await shapeWithTwoRuns();
  const before = r.getSlideOoxml(0);
  assert(before.includes('b="1"'), 'precondition: bold run present');
  // Insert "Brave " at paragraph offset 6 (boundary between "Hello " and "World").
  const ret = r.replaceTextRange(0, sIdx, 0, 6, 0, 6, 'Brave ');
  assert(!ret.startsWith('ERROR'), `replaceTextRange should succeed, got ${ret}`);
  const after = r.getSlideOoxml(0);
  assert(after.includes('Brave'), 'inserted text should be present');
  assert(after.includes('Hello') && after.includes('World'), 'original text should be preserved');
  assert(after.includes('b="1"'), 'bold run (World) should survive the insert');
  console.log('  OK: insert preserves format');
}

// --- Test 47: replaceTextRange deletes across runs ---
console.log('Test 47: replaceTextRange delete across runs');
{
  const { r, sIdx } = await shapeWithTwoRuns();
  // "Hello World" → delete offsets 3..8 ("lo Wo") → "Helrld".
  r.replaceTextRange(0, sIdx, 0, 3, 0, 8, '');
  const ooxml = r.getSlideOoxml(0);
  // Inspect run texts directly (slide 0's pre-existing title also contains "Hello").
  const texts = [...ooxml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1]);
  assert(texts.includes('Hel'), `left remainder "Hel" should be a run, got ${JSON.stringify(texts)}`);
  assert(texts.includes('rld'), `right remainder "rld" should be a run, got ${JSON.stringify(texts)}`);
  assert(!texts.includes('Hello '), 'deleted "Hello " run should be gone');
  assert(!texts.includes('World'), 'deleted "World" run should be gone');
  assert(ooxml.includes('b="1"'), 'bold formatting of "rld" remainder should survive');
  console.log('  OK: delete across runs');
}

// --- Test 48: replaceTextRange merges paragraphs / splits on newline ---
console.log('Test 48: replaceTextRange paragraph merge + newline split');
{
  // Two paragraphs "AAA" / "BBB" → delete the boundary → single paragraph "AAABBB".
  const r = await freshRenderer();
  const sIdx = parseInt(r.addShape(0, 'rect', 914400, 914400, 5486400, 1828800, -1, -1, -1).split(':')[1]);
  r.addShapeText(0, sIdx, 'AAA', 1800, 0, 0, 0);  // para 0
  r.addParagraph(0, sIdx, 'BBB', '');             // para 1
  let layout = JSON.parse(r.getTextLayout(0, sIdx));
  assert(layout.lines.length === 2, `precondition: 2 lines, got ${layout.lines.length}`);
  r.replaceTextRange(0, sIdx, 0, 3, 1, 0, '');     // merge
  layout = JSON.parse(r.getTextLayout(0, sIdx));
  assert(layout.lines.length === 1, `merge should yield 1 line, got ${layout.lines.length}`);

  // Newline split: insert "X\nY" inside the merged "AAABBB" → 2 paragraphs again.
  r.replaceTextRange(0, sIdx, 0, 3, 0, 3, 'X\nY');
  layout = JSON.parse(r.getTextLayout(0, sIdx));
  assert(layout.lines.length === 2, `newline should split into 2 lines, got ${layout.lines.length}`);
  console.log('  OK: paragraph merge + newline split');
}

// --- Test 49: replaceTextRange is undoable (E6.1 integration) ---
console.log('Test 49: replaceTextRange undo');
{
  const { r, sIdx } = await shapeWithTwoRuns();
  const before = r.renderSlideSvg(0);
  r.replaceTextRange(0, sIdx, 0, 0, 0, 11, 'Replaced');
  const after = r.renderSlideSvg(0);
  assert(after !== before, 'replace should change the SVG');
  assert(r.canUndo(), 'replaceTextRange should be undoable');
  r.undo();
  assert(r.renderSlideSvg(0) === before, 'undo should restore the original text');
  console.log('  OK: replaceTextRange undo');
}

// ── Z-order (E6.3) ───────────────────────────────────────────────────────────
// Serialized shape order == z-order (later = front). Find each shape by its fill
// color's position in the slide OOXML.
function fillPos(ooxml, hex) {
  return ooxml.toLowerCase().indexOf(hex.toLowerCase());
}
// Add red + green rects; returns their shape indices.
async function twoColoredShapes() {
  const r = await freshRenderer();
  const red = parseInt(r.addShape(0, 'rect', 0, 0, 914400, 914400, 255, 0, 0).split(':')[1]);
  const green = parseInt(r.addShape(0, 'rect', 457200, 457200, 914400, 914400, 0, 255, 0).split(':')[1]);
  return { r, red, green };
}

// --- Test 50: bringToFront moves a shape to the end (front) ---
console.log('Test 50: bringToFront');
{
  const { r, red, green } = await twoColoredShapes();
  let ooxml = r.getSlideOoxml(0);
  assert(fillPos(ooxml, 'ff0000') < fillPos(ooxml, '00ff00'), 'precondition: red is behind green');

  const ret = r.bringToFront(0, red);
  assert(ret.startsWith('OK:'), `bringToFront should return OK:<idx>, got ${ret}`);
  ooxml = r.getSlideOoxml(0);
  assert(fillPos(ooxml, 'ff0000') > fillPos(ooxml, '00ff00'), 'red should now be in front of green');
  console.log('  OK: bringToFront');
}

// --- Test 51: sendToBack moves a shape to index 0 ---
console.log('Test 51: sendToBack');
{
  const { r, red, green } = await twoColoredShapes();
  const ret = r.sendToBack(0, green);
  assert(ret === 'OK:0', `sendToBack should return OK:0, got ${ret}`);
  const ooxml = r.getSlideOoxml(0);
  assert(fillPos(ooxml, '00ff00') < fillPos(ooxml, 'ff0000'), 'green should now be behind red');
  console.log('  OK: sendToBack');
}

// --- Test 52: bringForward / sendBackward swap; no-op at the edge ---
console.log('Test 52: bringForward / sendBackward');
{
  const { r, red, green } = await twoColoredShapes();
  // red is directly behind green → bringForward swaps them.
  const fwd = r.bringForward(0, red);
  assert(fwd === `OK:${red + 1}`, `bringForward should return OK:${red + 1}, got ${fwd}`);
  let ooxml = r.getSlideOoxml(0);
  assert(fillPos(ooxml, 'ff0000') > fillPos(ooxml, '00ff00'), 'red should be in front after bringForward');

  // Move it back down again.
  const back = r.sendBackward(0, red + 1);
  assert(back === `OK:${red}`, `sendBackward should return OK:${red}, got ${back}`);
  ooxml = r.getSlideOoxml(0);
  assert(fillPos(ooxml, 'ff0000') < fillPos(ooxml, '00ff00'), 'red should be behind again');

  // green is front-most → bringForward is a no-op (index unchanged).
  const noop = r.bringForward(0, green);
  assert(noop === `OK:${green}`, `bringForward at front should be a no-op, got ${noop}`);
  console.log('  OK: bringForward / sendBackward + edge no-op');
}

// --- Test 53: z-order change is undoable ---
console.log('Test 53: z-order undo');
{
  const { r, red, green } = await twoColoredShapes();
  const before = r.getSlideOoxml(0);
  r.bringToFront(0, red);
  const after = r.getSlideOoxml(0);
  assert(after !== before, 'bringToFront should change the slide');
  r.undo();
  assert(r.getSlideOoxml(0) === before, 'undo should restore the original z-order');
  console.log('  OK: z-order undo');
}

// ── Multi-shape transform (E6.4) ─────────────────────────────────────────────
function shapeX(r, idx) {
  const m = r.renderShapeSvg(0, idx).match(/data-ooxml-x="(-?\d+)"/);
  return m ? parseInt(m[1]) : NaN;
}

// --- Test 54: updateShapesTransform — atomic batch + single undo ---
console.log('Test 54: updateShapesTransform (atomic + undo)');
{
  const { r, red, green } = await twoColoredShapes();
  const redX0 = shapeX(r, red);
  const greenX0 = shapeX(r, green);

  // Batch move both.
  const ret = r.updateShapesTransform(0, [
    { shapeIdx: red, x: 1000000, y: 1000000, cx: 914400, cy: 914400, rot: 0 },
    { shapeIdx: green, x: 2000000, y: 2000000, cx: 914400, cy: 914400, rot: 0 },
  ]);
  assert(ret === 'OK:2', `updateShapesTransform should return OK:2, got ${ret}`);
  assert(shapeX(r, red) === 1000000, 'red should move to x=1000000');
  assert(shapeX(r, green) === 2000000, 'green should move to x=2000000');

  // Atomicity: one bad index → nothing applied.
  const bad = r.updateShapesTransform(0, [
    { shapeIdx: red, x: 5000000, y: 0, cx: 914400, cy: 914400, rot: 0 },
    { shapeIdx: 9999, x: 0, y: 0, cx: 914400, cy: 914400, rot: 0 },
  ]);
  assert(bad.startsWith('ERROR'), `bad index should error, got ${bad}`);
  assert(shapeX(r, red) === 1000000, 'atomic: red must NOT move when batch fails');

  // Single undo reverts the whole successful batch.
  r.undo();
  assert(shapeX(r, red) === redX0, 'undo should restore red');
  assert(shapeX(r, green) === greenX0, 'undo should restore green');
  console.log('  OK: updateShapesTransform atomic batch + single undo');
}

// ── Copy / paste — cross-slide (E6.5) ────────────────────────────────────────
const MINI_PNG = new Uint8Array([
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

// --- Test 55: copy a plain shape to another slide ---
console.log('Test 55: getShapeSpec / insertShapeSpec (plain shape)');
{
  const r = await freshRenderer();
  const red = parseInt(r.addShape(0, 'rect', 100000, 100000, 914400, 914400, 255, 0, 0).split(':')[1]);
  const spec = r.getShapeSpec(0, red);
  assert(!spec.startsWith('ERROR'), `getShapeSpec should succeed, got ${spec.slice(0, 40)}`);
  const parsed = JSON.parse(spec);
  assert(typeof parsed.xml === 'string' && parsed.xml.includes('<p:sp'), 'spec should carry shape XML');
  assert(Array.isArray(parsed.media) && parsed.media.length === 0, 'plain shape has no media');

  // Paste onto a new slide with an offset.
  const { insertedIdx } = await r.addSlide();
  const ins = r.insertShapeSpec(insertedIdx, spec, 457200, 457200);
  assert(ins.startsWith('OK:'), `insertShapeSpec should return OK:<idx>, got ${ins}`);
  const svg = r.renderSlideSvg(insertedIdx);
  assert(svg.includes('ff0000') || svg.includes('FF0000'), 'pasted red shape should render on the target slide');
  // Offset applied: original x=100000 → pasted x=557200.
  const pastedIdx = parseInt(ins.split(':')[1]);
  const px = r.renderShapeSvg(insertedIdx, pastedIdx).match(/data-ooxml-x="(-?\d+)"/);
  assert(px && parseInt(px[1]) === 557200, `paste offset should apply (x=557200), got ${px && px[1]}`);
  console.log('  OK: plain shape cross-slide copy/paste');
}

// --- Test 56: copy an image shape — media re-link + export ---
console.log('Test 56: copy image shape (media re-link)');
{
  const r = await freshRenderer();
  const addRes = r.addImage(0, MINI_PNG, 'image/png', 100000, 100000, 914400, 914400);
  const imgIdx = parseInt(addRes.split(':')[1]);
  const spec = r.getShapeSpec(0, imgIdx);
  const parsed = JSON.parse(spec);
  assert(parsed.media.length === 1, `image shape should carry 1 media entry, got ${parsed.media.length}`);
  assert(parsed.media[0].b64.length > 0, 'media should be base64-encoded inline');
  assert(parsed.media[0].mime === 'image/png', 'media mime should be image/png');

  const { insertedIdx } = await r.addSlide();
  const ins = r.insertShapeSpec(insertedIdx, spec, 0, 0);
  assert(ins.startsWith('OK:'), `image paste should return OK:<idx>, got ${ins}`);

  // The target slide's .rels must reference a (new) media file, and the shape XML
  // must point at that new rId — verify the slide renders an <image>.
  const svg = r.renderSlideSvg(insertedIdx);
  assert(svg.includes('<image'), 'pasted image should render as <image> on the target slide');

  // Export must include the media binary and rebuild cleanly.
  const out = await r.exportPptx();
  assert(out.byteLength > 0, 'export with pasted image should produce bytes');
  console.log('  OK: image cross-slide copy/paste + export');
}

// --- Test 57: insertShapeSpec is undoable ---
console.log('Test 57: paste undo');
{
  const r = await freshRenderer();
  const red = parseInt(r.addShape(0, 'rect', 0, 0, 914400, 914400, 255, 0, 0).split(':')[1]);
  const spec = r.getShapeSpec(0, red);
  const { insertedIdx } = await r.addSlide();
  const before = r.renderSlideSvg(insertedIdx);
  r.insertShapeSpec(insertedIdx, spec);
  const after = r.renderSlideSvg(insertedIdx);
  assert(after !== before, 'paste should change the target slide');
  r.undo();
  assert(r.renderSlideSvg(insertedIdx) === before, 'undo should remove the pasted shape');
  console.log('  OK: paste undo');
}

// --- Test 58: boundary errors ---
console.log('Test 58: copy/paste boundary errors');
{
  const r = await freshRenderer();
  assert(r.getShapeSpec(0, 9999).startsWith('ERROR'), 'getShapeSpec on bad index should error');
  assert(r.insertShapeSpec(0, 'not json').startsWith('ERROR'), 'insertShapeSpec with bad spec should error');
  assert(r.insertShapeSpec(0, JSON.stringify({ xml: '', media: [] })).startsWith('ERROR'), 'empty xml should error');
  console.log('  OK: boundary errors');
}

// ── Table editing (E6.6) ─────────────────────────────────────────────────────
// test_features.pptx slide 41 (0-indexed) has a plain 3×3 table at shape index 0.
const FEATURES_PATH = join(__dirname, 'test_features.pptx');
const TABLE_SLIDE = 41, TABLE_SHAPE = 0;
async function freshFeatures() {
  const renderer = new PptxRenderer({ logLevel: 'silent' });
  renderer.__hasFeatures = false;
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const buf = readFileSync(FEATURES_PATH);
  await renderer.loadPptx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  return renderer;
}
const trCount = (xml) => (xml.match(/<a:tr[ >]/g) || []).length;
const gridColCount = (xml) => (xml.match(/<a:gridCol/g) || []).length;

if (!existsSync(FEATURES_PATH)) {
  console.log('Tests 59–62: SKIPPED (test_features.pptx not found)');
} else {
  // --- Test 59: updateTableCellText + undo ---
  console.log('Test 59: updateTableCellText');
  {
    const r = await freshFeatures();
    const ret = r.updateTableCellText(TABLE_SLIDE, TABLE_SHAPE, 0, 0, 'EDITED_CELL_X');
    assert(!ret.startsWith('ERROR'), `updateTableCellText should succeed, got ${ret.slice(0, 40)}`);
    const ooxml = r.getSlideOoxml(TABLE_SLIDE);
    assert(ooxml.includes('EDITED_CELL_X'), 'new cell text should appear in the slide OOXML');

    const before = r.renderSlideSvg(TABLE_SLIDE);
    r.updateTableCellText(TABLE_SLIDE, TABLE_SHAPE, 1, 1, 'UNDO_ME');
    assert(r.renderSlideSvg(TABLE_SLIDE) !== before, 'cell edit should change the slide');
    r.undo();
    assert(r.renderSlideSvg(TABLE_SLIDE) === before, 'undo should restore the cell');
    console.log('  OK: updateTableCellText + undo');
  }

  // --- Test 60: add/delete row ---
  console.log('Test 60: add/delete table row');
  {
    const r = await freshFeatures();
    const rows0 = trCount(r.getSlideOoxml(TABLE_SLIDE));
    const add = r.addTableRow(TABLE_SLIDE, TABLE_SHAPE, 0);
    assert(add.startsWith('OK:'), `addTableRow should return OK:<idx>, got ${add}`);
    assert(trCount(r.getSlideOoxml(TABLE_SLIDE)) === rows0 + 1, 'row count should increase by 1');
    const del = r.deleteTableRow(TABLE_SLIDE, TABLE_SHAPE, 0);
    assert(del === 'OK', `deleteTableRow should return OK, got ${del}`);
    assert(trCount(r.getSlideOoxml(TABLE_SLIDE)) === rows0, 'row count should return to original');
    console.log('  OK: add/delete row');
  }

  // --- Test 61: add/delete column ---
  console.log('Test 61: add/delete table column');
  {
    const r = await freshFeatures();
    const cols0 = gridColCount(r.getSlideOoxml(TABLE_SLIDE));
    const add = r.addTableColumn(TABLE_SLIDE, TABLE_SHAPE, 0, 914400);
    assert(add.startsWith('OK:'), `addTableColumn should return OK:<idx>, got ${add}`);
    let ooxml = r.getSlideOoxml(TABLE_SLIDE);
    assert(gridColCount(ooxml) === cols0 + 1, 'gridCol count should increase by 1');
    // Every row must gain a cell (tc count = rows * cols).
    const rows = trCount(ooxml);
    const tcCount = (ooxml.match(/<a:tc[ >]/g) || []).length;
    assert(tcCount === rows * (cols0 + 1), `tc count should be rows*cols (${rows}*${cols0 + 1}), got ${tcCount}`);
    const del = r.deleteTableColumn(TABLE_SLIDE, TABLE_SHAPE, 0);
    assert(del === 'OK', `deleteTableColumn should return OK, got ${del}`);
    assert(gridColCount(r.getSlideOoxml(TABLE_SLIDE)) === cols0, 'gridCol count should return to original');
    console.log('  OK: add/delete column');
  }

  // --- Test 62: boundary errors ---
  console.log('Test 62: table editing boundary errors');
  {
    const r = await freshFeatures();
    // shape 1 on this slide is not a table.
    assert(r.addTableRow(TABLE_SLIDE, 1, -1).startsWith('ERROR'), 'non-table shape should error');
    assert(r.updateTableCellText(TABLE_SLIDE, TABLE_SHAPE, 99, 0, 'x').startsWith('ERROR'), 'row out of range should error');
    assert(r.deleteTableColumn(TABLE_SLIDE, TABLE_SHAPE, 99).startsWith('ERROR'), 'col out of range should error');
    // Delete rows down to the last one → next delete must error.
    const rows = trCount(r.getSlideOoxml(TABLE_SLIDE));
    for (let i = 0; i < rows - 1; i++) r.deleteTableRow(TABLE_SLIDE, TABLE_SHAPE, 0);
    assert(r.deleteTableRow(TABLE_SLIDE, TABLE_SHAPE, 0).startsWith('ERROR'), 'deleting the last row should error');
    console.log('  OK: boundary errors');
  }
}

// ── Header/footer field placeholders (date / footer / slide number) ──────────
// test_features.pptx slide 97 (index 96) carries dt/ftr/sldNum placeholders with
// <a:fld> fields and a STALE cached slide number ("1"), mirroring sample.pptx s20.
{
  console.log('Test 63: header/footer field placeholders');
  const renderer = new PptxRenderer({
    logLevel: 'silent',
    currentDate: '2026-06-15', // fixed string → locale-independent assertion
  });
  const wasmBuf = readFileSync(join(__dirname, '..', 'dist', 'main.wasm'));
  await renderer.init(wasmBuf);
  const buf = readFileSync(FEATURES_PATH);
  await renderer.loadPptx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  const svg = renderer.renderSlideSvg(96); // 0-based → slide 97
  // Slide number reflects the actual index (97), not the cached "1".
  assert(svg.includes('>97</tspan>'), 'slide number should render the actual index 97 (not cached 1)');
  // Date field filled with the host-provided current date (verbatim string).
  assert(svg.includes('2026-06-15'), 'date field should show the provided current date');
  // Footer text preserved.
  assert(svg.includes('moon-pptx'), 'footer text should render');
  console.log('  OK: header/footer field placeholders');
}

// --- Summary ---
console.log('');
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
} else {
  console.log('All Node.js compatibility tests passed!');
}
