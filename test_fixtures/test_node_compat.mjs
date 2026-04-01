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

// --- Summary ---
console.log('');
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
} else {
  console.log('All Node.js compatibility tests passed!');
}
