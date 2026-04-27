/**
 * Security regression tests.
 *
 * Covers fixes from PR #30 and PR #31:
 *   - L1: getSlideNotes / getSlideComments decode XML entities so consumers
 *     get plain text (not raw `&lt;`), making the API safe for `.textContent`.
 *   - L2: dynamic `new RegExp(...)` paths must escape PPTX-derived rIds /
 *     targets so regex metachars don't throw `SyntaxError`.
 *   - M1: ZIP extraction caps decompression at 256 MiB per entry / 1 GiB per
 *     archive, defending against decompression bombs.
 */

import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { expect, resetAssertions, finishAssertions, FIXTURES_DIR, DIST_DIR } from './_helpers.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const WASM_PATH = join(DIST_DIR, 'main.wasm');

async function loadRenderer() {
  const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
  const r = new PptxRenderer({ logLevel: 'silent' });
  await r.init(readFileSync(WASM_PATH));
  return r;
}

function readFixtureAb(name) {
  const buf = readFileSync(join(FIXTURES_DIR, name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

// ── ZIP construction helpers (for M1 bomb test) ─────────────────────────────

/** CRC-32 over a Uint8Array (matches lib/utils.ts crc32). */
function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Build a minimal single-entry ZIP. `declaredUncompressedSize` lets the
 * caller fake the central-directory size field (used to test the pre-check
 * cap without actually allocating gigabytes).
 */
function buildSingleEntryZip({ name, content, declaredUncompressedSize }) {
  const enc = new TextEncoder();
  const nameBytes = enc.encode(name);
  const compressed = deflateRawSync(content, { level: 9 });
  const realCrc = crc32(content);
  const declaredSize = declaredUncompressedSize ?? content.length;

  // Local header (30 + name + extra)
  const lh = new Uint8Array(30);
  const dvLh = new DataView(lh.buffer);
  dvLh.setUint32(0, 0x04034b50, true);
  dvLh.setUint16(4, 20, true);
  dvLh.setUint16(6, 0, true);
  dvLh.setUint16(8, 8, true); // method: deflate
  dvLh.setUint16(10, 0, true);
  dvLh.setUint16(12, 0, true);
  dvLh.setUint32(14, realCrc, true);
  dvLh.setUint32(18, compressed.length, true);
  dvLh.setUint32(22, declaredSize, true);
  dvLh.setUint16(26, nameBytes.length, true);
  dvLh.setUint16(28, 0, true);

  // Central dir entry
  const cd = new Uint8Array(46);
  const dvCd = new DataView(cd.buffer);
  dvCd.setUint32(0, 0x02014b50, true);
  dvCd.setUint16(4, 20, true);
  dvCd.setUint16(6, 20, true);
  dvCd.setUint16(8, 0, true);
  dvCd.setUint16(10, 8, true);
  dvCd.setUint32(16, realCrc, true);
  dvCd.setUint32(20, compressed.length, true);
  dvCd.setUint32(24, declaredSize, true);
  dvCd.setUint16(28, nameBytes.length, true);
  dvCd.setUint16(42, 0, true); // local header offset

  // EOCD
  const localSize = 30 + nameBytes.length + compressed.length;
  const cdSize = 46 + nameBytes.length;
  const eocd = new Uint8Array(22);
  const dvE = new DataView(eocd.buffer);
  dvE.setUint32(0, 0x06054b50, true);
  dvE.setUint16(8, 1, true);
  dvE.setUint16(10, 1, true);
  dvE.setUint32(12, cdSize, true);
  dvE.setUint32(16, localSize, true);

  const total = localSize + cdSize + 22;
  const out = new Uint8Array(total);
  let pos = 0;
  out.set(lh, pos); pos += 30;
  out.set(nameBytes, pos); pos += nameBytes.length;
  out.set(compressed, pos); pos += compressed.length;
  out.set(cd, pos); pos += 46;
  out.set(nameBytes, pos); pos += nameBytes.length;
  out.set(eocd, pos);
  return out.buffer;
}

// ── Tests ───────────────────────────────────────────────────────────────────

test('M1: extractZip rejects entries with declared size over the cap', async () => {
  resetAssertions();
  const { extractZip } = await import(join(DIST_DIR, 'zip.js'));

  // Tiny actual content but the central directory claims 500 MiB. The
  // pre-check at zip.ts must skip this entry before inflating.
  const huge = 500 * 1024 * 1024;
  const ab = buildSingleEntryZip({
    name: 'evil.bin',
    content: new Uint8Array(8),
    declaredUncompressedSize: huge,
  });

  let warned = false;
  const log = { debug() {}, info() {}, warn() { warned = true; }, error() {} };
  const { textFiles, binaryFiles } = await extractZip(ab, log);

  expect('oversized entry was skipped', !textFiles.has('evil.bin') && !binaryFiles.has('evil.bin'));
  expect('warning was emitted', warned);

  finishAssertions();
});

test('M1: extractZip aborts streaming inflate past the cap', async () => {
  resetAssertions();
  const { extractZip } = await import(join(DIST_DIR, 'zip.js'));

  // Compressed payload that actually inflates to ~260 MiB of zeros. The
  // streaming check inside inflate() must reject this regardless of what
  // the central directory claims.
  const oversize = 260 * 1024 * 1024;
  const zeros = new Uint8Array(oversize); // 260 MiB of zeros
  const ab = buildSingleEntryZip({
    name: 'bomb.bin',
    content: zeros,
    // Lie: claim the entry is small so the pre-check passes.
    declaredUncompressedSize: 1024,
  });

  let threw = null;
  try {
    await extractZip(ab);
  } catch (e) {
    threw = e;
  }
  expect('streaming inflate cap fires', threw !== null,
    threw ? threw.message : 'no error thrown');
  expect('error message mentions decompressed size cap',
    threw && /Decompressed size exceeded|Archive total decompressed/.test(threw.message),
    threw ? threw.message : '');

  finishAssertions();
});

test('L1 + L2: PPTX with metachar rIds and entity-encoded notes works end-to-end', async () => {
  resetAssertions();
  const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
  const { extractZip, buildZip } = await import(join(DIST_DIR, 'zip.js'));

  // Start from the regular fixture, then inject:
  //   1. A notesSlide containing `&amp;` and `&lt;script&gt;` entities — L1
  //      (getSlideNotes must return decoded text).
  //   2. A relationship whose rId contains regex metachars — L2
  //      (no RegExp SyntaxError; renderSlide / save must not throw).
  const ab = readFixtureAb('test_features.pptx');
  const { textFiles, binaryFiles } = await extractZip(ab);

  // Locate notesSlide1.xml and replace its body text with entity-encoded chars.
  // In test_features.pptx, notesSlide1.xml is referenced from slide75.xml
  // (slide index 74).
  const notesPath = 'ppt/notesSlides/notesSlide1.xml';
  const origNotes = textFiles.get(notesPath);
  expect('notesSlide1.xml present in fixture', !!origNotes);
  const evilText = '&amp; &lt;script&gt;alert(1)&lt;/script&gt; &#65;';
  // Replace the body placeholder text content with our entity-encoded string.
  // Match the first <a:t>...</a:t> inside a body placeholder sp.
  const patchedNotes = origNotes.replace(
    /(<p:sp\b[^]*?<p:ph[^>]+type="body"[^]*?<a:t>)[^<]*(<\/a:t>)/,
    `$1${evilText}$2`,
  );
  expect('notes were patched', patchedNotes !== origNotes);

  // Add a relationship whose rId contains regex metachars: `rId$(.*)+`.
  // Append to slide1's rels — rendering shouldn't try to use it,
  // but resolveRidTarget / extractRelTarget must not throw.
  const slide1RelsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const origRels = textFiles.get(slide1RelsPath) ?? '';
  const evilRid = 'rId$(.*)+';
  const evilEntry = `<Relationship Id="${evilRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>`;
  const patchedRels = origRels.replace('</Relationships>', evilEntry + '</Relationships>');

  // Rebuild the PPTX with our patches.
  const mods = new Map([
    [notesPath, patchedNotes],
    [slide1RelsPath, patchedRels],
  ]);
  const newAb = await buildZip(ab, mods, undefined, undefined);

  // Load through the renderer and exercise L1 + L2 paths.
  const r = await loadRenderer();
  const { slideCount } = await r.loadPptx(newAb);
  expect('slides loaded', slideCount > 0);

  // L1: getSlideNotes for the slide that owns notesSlide1.xml.
  const NOTES_SLIDE_IDX = 74; // slide75.xml in this fixture
  const notes = r.getSlideNotes(NOTES_SLIDE_IDX);
  const joined = notes.join(' ');
  expect('& was decoded', joined.includes('&'));
  expect('< was decoded', joined.includes('<script>'));
  expect('numeric ref decoded', joined.includes('A'));
  expect('no raw &amp; remains', !joined.includes('&amp;'),
    `got: ${joined}`);
  expect('no raw &lt; remains', !joined.includes('&lt;'));

  // L2: rendering slide 1 must not throw despite the metachar rId.
  let renderErr = null;
  try {
    const svg = r.renderSlideSvg(0);
    expect('render returned a string', typeof svg === 'string');
  } catch (e) {
    renderErr = e;
  }
  expect('renderSlideSvg did not throw on metachar rId', renderErr === null,
    renderErr ? renderErr.message : '');

  // Export round-trips without throwing (exercises the rels regex paths).
  let exportErr = null;
  try {
    const out = await r.exportPptx();
    expect('export returned bytes', out.byteLength > 0);
  } catch (e) {
    exportErr = e;
  }
  expect('exportPptx did not throw on metachar rId', exportErr === null,
    exportErr ? exportErr.message : '');

  finishAssertions();
});
