import { test } from 'node:test';
import { expect, assert, section, resetAssertions, finishAssertions } from './_helpers.mjs';

test('WMF converter', async () => {
  resetAssertions();
  const { wmfToSvg } = await import('../../dist/wmf-converter.js');

  section('WMF converter — invalid input');
  assert('empty data returns ""', wmfToSvg(new Uint8Array(0)) === '');
  assert('too small returns ""', wmfToSvg(new Uint8Array(10)) === '');
  assert('non-WMF data returns ""', wmfToSvg(new Uint8Array(100)) === '');

  section('WMF converter — minimal standard WMF (no drawings)');
  // Build: METAHEADER(18) + EOF record(6)
  const minWmf = new Uint8Array(24);
  const dv = new DataView(minWmf.buffer);
  dv.setUint16(0, 1, true);     // mtType = memory
  dv.setUint16(2, 9, true);     // mtHeaderSize = 9 words
  dv.setUint16(4, 0x0300, true); // mtVersion
  dv.setUint32(6, 12, true);    // mtSize = 12 words (24 bytes total)
  dv.setUint16(10, 0, true);    // mtNoObjects
  dv.setUint32(12, 3, true);    // mtMaxRecord = 3 words
  dv.setUint16(16, 0, true);    // mtNoParameters
  // EOF record at offset 18
  dv.setUint32(18, 3, true);    // size = 3 words
  dv.setUint16(22, 0, true);    // function = META_EOF
  assert('minimal WMF returns "" (no drawings)', wmfToSvg(minWmf) === '');

  section('WMF converter — standard WMF with rectangle');
  const rectWmf = buildTestWmfWithRect();
  const rectResult = wmfToSvg(rectWmf);
  assert('rect WMF returns non-empty SVG', rectResult.length > 0);
  assert('SVG starts with <svg', rectResult.startsWith('<svg'));
  assert('SVG ends with </svg>', rectResult.endsWith('</svg>'));
  assert('SVG contains <rect', rectResult.includes('<rect '));
  assert('SVG has viewBox', rectResult.includes('viewBox='));

  section('WMF converter — placeable WMF with polygon');
  const placeableWmf = buildTestPlaceableWmf();
  const placeableResult = wmfToSvg(placeableWmf);
  assert('placeable WMF returns non-empty SVG', placeableResult.length > 0);
  assert('SVG contains <path (polygon)', placeableResult.includes('<path '));
  assert('SVG contains fill color', placeableResult.includes('#ff0000'));
  assert('viewBox uses placeable bbox', placeableResult.includes('viewBox="0 0 200 100"'));
  finishAssertions();
});

// ── Fixture builders ─────────────────────────────────────────────────────────

function buildTestWmfWithRect() {
  // METAHEADER(18) + SetWindowOrg(10) + SetWindowExt(10) +
  // CreateBrushIndirect(14) + SelectObject(8) + Rectangle(14) + EOF(6)
  // Total: 80 bytes
  const buf = new Uint8Array(80);
  const dv = new DataView(buf.buffer);
  let off = 0;

  // METAHEADER (18 bytes)
  dv.setUint16(off, 1, true);      // mtType
  dv.setUint16(off + 2, 9, true);  // mtHeaderSize
  dv.setUint16(off + 4, 0x0300, true); // mtVersion
  dv.setUint32(off + 6, 40, true); // mtSize in words (80 bytes)
  dv.setUint16(off + 10, 1, true); // mtNoObjects
  dv.setUint32(off + 12, 7, true); // mtMaxRecord
  dv.setUint16(off + 16, 0, true);
  off = 18;

  // META_SETWINDOWORG (size=5 words, func=0x020B): Y=0, X=0
  dv.setUint32(off, 5, true);
  dv.setUint16(off + 4, 0x020B, true);
  dv.setInt16(off + 6, 0, true);  // Y
  dv.setInt16(off + 8, 0, true);  // X
  off += 10;

  // META_SETWINDOWEXT (size=5 words, func=0x020C): CY=100, CX=200
  dv.setUint32(off, 5, true);
  dv.setUint16(off + 4, 0x020C, true);
  dv.setInt16(off + 6, 100, true); // CY
  dv.setInt16(off + 8, 200, true); // CX
  off += 10;

  // META_CREATEBRUSHINDIRECT (size=7 words, func=0x02FC): solid green
  dv.setUint32(off, 7, true);
  dv.setUint16(off + 4, 0x02FC, true);
  dv.setUint16(off + 6, 0, true);     // BS_SOLID
  dv.setUint8(off + 8, 0x00);         // R
  dv.setUint8(off + 9, 0xFF);         // G
  dv.setUint8(off + 10, 0x00);        // B
  dv.setUint8(off + 11, 0x00);        // reserved
  dv.setInt16(off + 12, 0, true);     // BrushHatch
  off += 14;

  // META_SELECTOBJECT (size=4 words, func=0x012D): object 0
  dv.setUint32(off, 4, true);
  dv.setUint16(off + 4, 0x012D, true);
  dv.setUint16(off + 6, 0, true);
  off += 8;

  // META_RECTANGLE (size=7 words, func=0x041B): bottom=80, right=150, top=10, left=20
  dv.setUint32(off, 7, true);
  dv.setUint16(off + 4, 0x041B, true);
  dv.setInt16(off + 6, 80, true);  // bottom
  dv.setInt16(off + 8, 150, true); // right
  dv.setInt16(off + 10, 10, true); // top
  dv.setInt16(off + 12, 20, true); // left
  off += 14;

  // META_EOF (size=3 words)
  dv.setUint32(off, 3, true);
  dv.setUint16(off + 4, 0, true);
  off += 6;

  return buf.subarray(0, off);
}

function buildTestPlaceableWmf() {
  // PlaceableHeader(22) + METAHEADER(18) +
  // CreateBrushIndirect(14) + SelectObject(8) + Polygon(14 for 3 pts) + EOF(6)
  const buf = new Uint8Array(120);
  const dv = new DataView(buf.buffer);
  let off = 0;

  // Placeable WMF header (22 bytes)
  dv.setUint32(off, 0x9AC6CDD7, true); // magic key
  dv.setUint16(off + 4, 0, true);       // HWmf
  dv.setInt16(off + 6, 0, true);        // BboxLeft
  dv.setInt16(off + 8, 0, true);        // BboxTop
  dv.setInt16(off + 10, 200, true);     // BboxRight
  dv.setInt16(off + 12, 100, true);     // BboxBottom
  dv.setUint16(off + 14, 96, true);     // Inch
  dv.setUint32(off + 16, 0, true);      // Reserved
  dv.setUint16(off + 20, 0, true);      // Checksum (ignored)
  off = 22;

  // METAHEADER (18 bytes)
  dv.setUint16(off, 1, true);
  dv.setUint16(off + 2, 9, true);
  dv.setUint16(off + 4, 0x0300, true);
  dv.setUint32(off + 6, 50, true);  // total size in words
  dv.setUint16(off + 10, 1, true);
  dv.setUint32(off + 12, 7, true);
  dv.setUint16(off + 16, 0, true);
  off += 18;

  // META_CREATEBRUSHINDIRECT: solid red
  dv.setUint32(off, 7, true);
  dv.setUint16(off + 4, 0x02FC, true);
  dv.setUint16(off + 6, 0, true);    // BS_SOLID
  dv.setUint8(off + 8, 0xFF);        // R
  dv.setUint8(off + 9, 0x00);        // G
  dv.setUint8(off + 10, 0x00);       // B
  dv.setUint8(off + 11, 0x00);
  dv.setInt16(off + 12, 0, true);
  off += 14;

  // META_SELECTOBJECT: object 0
  dv.setUint32(off, 4, true);
  dv.setUint16(off + 4, 0x012D, true);
  dv.setUint16(off + 6, 0, true);
  off += 8;

  // META_POLYGON (size = 4 + nPts*2 words): 3 points (triangle)
  // size = 4 + 3*2 = 10 words
  dv.setUint32(off, 10, true);
  dv.setUint16(off + 4, 0x0324, true);
  dv.setInt16(off + 6, 3, true);   // nPts
  // point 1: (10, 90)
  dv.setInt16(off + 8, 10, true);
  dv.setInt16(off + 10, 90, true);
  // point 2: (100, 10)
  dv.setInt16(off + 12, 100, true);
  dv.setInt16(off + 14, 10, true);
  // point 3: (190, 90)
  dv.setInt16(off + 16, 190, true);
  dv.setInt16(off + 18, 90, true);
  off += 20;

  // META_EOF
  dv.setUint32(off, 3, true);
  dv.setUint16(off + 4, 0, true);
  off += 6;

  return buf.subarray(0, off);
}
