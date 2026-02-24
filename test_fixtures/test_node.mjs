/**
 * Node.js test script for pptx-render Phase 1
 *
 * Tests the JavaScript host layer (ZIP extraction) independently
 * of the Wasm module.
 *
 * Run: node test_fixtures/test_node.mjs
 */

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── Minimal ZIP extractor (mirrors host.js logic) ─────────────────────────────

async function extractZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const textFiles = new Map();
  const decoder = new TextDecoder('utf-8');

  // Node.js 18+ has DecompressionStream
  async function inflate(compressed) {
    const { DecompressionStream } = await import('node:stream/web');
    const stream = new DecompressionStream('deflate-raw');
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();
    writer.write(compressed);
    writer.close();
    const chunks = [];
    let totalLen = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      totalLen += value.length;
    }
    const result = new Uint8Array(totalLen);
    let pos = 0;
    for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
    return result;
  }

  let offset = 0;
  while (offset < bytes.length - 4) {
    const sig = view.getUint32(offset, true);
    if (sig !== 0x04034b50) break;

    const method          = view.getUint16(offset + 8, true);
    const compressedSize  = view.getUint32(offset + 18, true);
    const fileNameLen     = view.getUint16(offset + 26, true);
    const extraLen        = view.getUint16(offset + 28, true);

    const name = decoder.decode(bytes.slice(offset + 30, offset + 30 + fileNameLen));
    const dataOffset = offset + 30 + fileNameLen + extraLen;
    const compressedData = bytes.slice(dataOffset, dataOffset + compressedSize);

    let decompressed;
    if (method === 0) {
      decompressed = compressedData;
    } else if (method === 8) {
      decompressed = await inflate(compressedData);
    } else {
      console.warn(`Unsupported compression method ${method} for ${name}`);
      offset = dataOffset + compressedSize;
      continue;
    }

    const isText = name.toLowerCase().endsWith('.xml') ||
                   name.toLowerCase().endsWith('.rels');
    if (isText) {
      textFiles.set(name, decoder.decode(decompressed));
    }

    offset = dataOffset + compressedSize;
  }
  return textFiles;
}

// ── Slide count logic (mirrors main.mbt logic) ─────────────────────────────

function countSlideIds(xml) {
  const patterns = ['<p:sldId ', '<p:sldId\t', '<p:sldId\n', '<p:sldId/>'];
  let total = 0;
  for (const pat of patterns) {
    let pos = 0;
    while (true) {
      const idx = xml.indexOf(pat, pos);
      if (idx === -1) break;
      total++;
      pos = idx + pat.length;
    }
  }
  return total;
}

// ── Test runner ───────────────────────────────────────────────────────────────

async function runTests() {
  console.log('=== pptx-render Phase 1 Node.js Tests ===\n');

  const pptxPath = join(__dirname, 'minimal.pptx');
  // Node.js Buffer shares underlying ArrayBuffer with a pool, so we must slice
  // to get a standalone ArrayBuffer with byteOffset=0
  const nodeBuffer = readFileSync(pptxPath);
  const pptxBuffer = nodeBuffer.buffer.slice(
    nodeBuffer.byteOffset,
    nodeBuffer.byteOffset + nodeBuffer.byteLength,
  );

  console.log(`Loading: ${pptxPath}`);
  const files = await extractZip(pptxBuffer);

  console.log(`\nExtracted ${files.size} text entries:`);
  for (const [name] of files) {
    console.log(`  ${name}`);
  }

  // Test 1: presentation.xml exists
  const prsXml = files.get('ppt/presentation.xml');
  console.log('\n[TEST 1] presentation.xml exists:', prsXml ? '✓ PASS' : '✗ FAIL');

  // Test 2: slide count
  const slideCount = countSlideIds(prsXml ?? '');
  console.log(`[TEST 2] Slide count = ${slideCount}:`, slideCount === 2 ? '✓ PASS' : `✗ FAIL (expected 2, got ${slideCount})`);

  // Test 3: slide1.xml exists
  const slide1 = files.get('ppt/slides/slide1.xml');
  console.log('[TEST 3] slide1.xml exists:', slide1 ? '✓ PASS' : '✗ FAIL');

  // Test 4: slide1.xml contains expected text
  const hasExpected = slide1?.includes('Hello from MoonBit') ?? false;
  console.log('[TEST 4] slide1.xml contains title text:', hasExpected ? '✓ PASS' : '✗ FAIL');

  // Test 5: slide2.xml exists
  const slide2 = files.get('ppt/slides/slide2.xml');
  console.log('[TEST 5] slide2.xml exists:', slide2 ? '✓ PASS' : '✗ FAIL');

  console.log('\n=== JS host layer tests DONE ===');
  console.log('\nTo test the full Wasm pipeline, open in Chrome/Firefox:');
  console.log('  http://localhost:8765/web/index.html');
  console.log('Then drag test_fixtures/minimal.pptx onto the page.');
}

runTests().catch(console.error);
