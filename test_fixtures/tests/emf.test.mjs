import { test } from 'node:test';
import { expect, assert, section, resetAssertions, finishAssertions } from './_helpers.mjs';

test('EMF converter', async () => {
  resetAssertions();
  const { emfToSvg } = await import('../../dist/emf-converter.js');

  section('EMF converter — invalid input');
  assert('empty data returns ""', emfToSvg(new Uint8Array(0)) === '');
  assert('too small returns ""', emfToSvg(new Uint8Array(50)) === '');
  assert('non-EMF data returns ""', emfToSvg(new Uint8Array(100)) === '');

  section('EMF converter — minimal valid EMF');
  // Build a minimal EMF: header (108 bytes) + EOF record (20 bytes)
  const minEmf = new Uint8Array(128);
  const dv = new DataView(minEmf.buffer);
  // Header record
  dv.setUint32(0, 1, true);     // type = EMR_HEADER
  dv.setUint32(4, 108, true);   // size
  // Bounds: L=0, T=0, R=100, B=50
  dv.setInt32(8, 0, true);
  dv.setInt32(12, 0, true);
  dv.setInt32(16, 100, true);
  dv.setInt32(20, 50, true);
  // Frame (unused but must be present)
  dv.setInt32(24, 0, true);
  dv.setInt32(28, 0, true);
  dv.setInt32(32, 25400, true);
  dv.setInt32(36, 12700, true);
  // Signature " EMF"
  dv.setUint32(40, 0x464D4520, true);
  // Version
  dv.setUint32(44, 0x10000, true);
  // File size, nRecords
  dv.setUint32(48, 128, true);
  dv.setUint32(52, 2, true);
  // EOF record at offset 108
  dv.setUint32(108, 0x0E, true);  // type = EMR_EOF
  dv.setUint32(112, 20, true);    // size

  const minResult = emfToSvg(minEmf);
  assert('minimal EMF returns "" (no drawing commands)', minResult === '');

  section('EMF converter — EMF with filled path');
  // Build EMF with: header + CREATEBRUSHINDIRECT + SELECTOBJECT +
  //   BEGINPATH + MOVETOEX + POLYLINETO16 (triangle) + CLOSEFIGURE +
  //   ENDPATH + FILLPATH + EOF
  const pathEmf = buildTestEmfWithPath();
  const pathResult = emfToSvg(pathEmf);
  assert('path EMF returns non-empty SVG', pathResult.length > 0);
  assert('SVG starts with <svg', pathResult.startsWith('<svg'));
  assert('SVG ends with </svg>', pathResult.endsWith('</svg>'));
  assert('SVG contains <path', pathResult.includes('<path '));
  assert('SVG contains fill color', pathResult.includes('#ff0000'));
  assert('SVG has viewBox', pathResult.includes('viewBox='));

  section('EMF converter — window extent mapping');
  // The viewBox should use window extent, not device bounds
  const wndEmf = buildTestEmfWithWindowExtent();
  const wndResult = emfToSvg(wndEmf);
  assert('window extent EMF returns SVG', wndResult.length > 0);
  assert('viewBox uses window extent (200 100)', wndResult.includes('viewBox="0 0 200 100"'));
  assert('width/height uses device bounds', wndResult.includes('width="100"') && wndResult.includes('height="50"'));
  finishAssertions();
});

// ── Fixture builders ─────────────────────────────────────────────────────────

function buildTestEmfWithPath() {
  // Records: Header(108) + CreateBrushIndirect(16) + SelectObject(12) +
  //   BeginPath(8) + MoveToEx(16) + PolyLineTo16(28+8=36 -> 3pts=40) +
  //   CloseFigure(8) + EndPath(8) + FillPath(24) + EOF(20)
  // Total: 108 + 16 + 12 + 8 + 16 + 40 + 8 + 8 + 24 + 20 = 260
  const buf = new Uint8Array(260);
  const dv = new DataView(buf.buffer);

  // Header
  dv.setUint32(0, 1, true); dv.setUint32(4, 108, true);
  dv.setInt32(8, 0, true); dv.setInt32(12, 0, true);
  dv.setInt32(16, 100, true); dv.setInt32(20, 50, true);
  dv.setInt32(24, 0, true); dv.setInt32(28, 0, true);
  dv.setInt32(32, 25400, true); dv.setInt32(36, 12700, true);
  dv.setUint32(40, 0x464D4520, true);
  dv.setUint32(44, 0x10000, true);
  dv.setUint32(48, 260, true); dv.setUint32(52, 10, true);

  let off = 108;

  // CreateBrushIndirect (type=0x27=39, size=16)
  dv.setUint32(off, 0x27, true); dv.setUint32(off + 4, 16, true);
  dv.setUint32(off + 8, 0, true);  // ihBrush = 0
  // No style field needed at correct offset... actually:
  // CreateBrushIndirect: +8: ihBrush(4), +12: lbStyle(4), +16: lbColor(4), +20: lbHatch(4)
  // But size=16 means only ihBrush + ... hmm, let me recalculate
  // Actually the record is: type(4) + size(4) + ihBrush(4) + logBrush{style(4)+color(4)+hatch(4)} = 24
  // Let me fix the sizes

  // Rebuild with correct sizes
  const buf2 = new Uint8Array(280);
  const dv2 = new DataView(buf2.buffer);

  // Header (108)
  dv2.setUint32(0, 1, true); dv2.setUint32(4, 108, true);
  dv2.setInt32(8, 0, true); dv2.setInt32(12, 0, true);
  dv2.setInt32(16, 100, true); dv2.setInt32(20, 50, true);
  dv2.setInt32(24, 0, true); dv2.setInt32(28, 0, true);
  dv2.setInt32(32, 25400, true); dv2.setInt32(36, 12700, true);
  dv2.setUint32(40, 0x464D4520, true);
  dv2.setUint32(44, 0x10000, true);
  dv2.setUint32(48, 280, true); dv2.setUint32(52, 10, true);
  off = 108;

  // CreateBrushIndirect (type=0x27, size=24): ihBrush=0, style=0(solid), color=red
  dv2.setUint32(off, 0x27, true); dv2.setUint32(off + 4, 24, true);
  dv2.setUint32(off + 8, 0, true);   // ihBrush
  dv2.setUint32(off + 12, 0, true);  // style = BS_SOLID
  dv2.setUint8(off + 16, 0xFF);      // R
  dv2.setUint8(off + 17, 0x00);      // G
  dv2.setUint8(off + 18, 0x00);      // B
  off += 24;

  // SelectObject (type=0x25, size=12): ihObject=0
  dv2.setUint32(off, 0x25, true); dv2.setUint32(off + 4, 12, true);
  dv2.setUint32(off + 8, 0, true);
  off += 12;

  // BeginPath (type=0x3B, size=8)
  dv2.setUint32(off, 0x3B, true); dv2.setUint32(off + 4, 8, true);
  off += 8;

  // MoveToEx (type=0x1B, size=16): x=10, y=40
  dv2.setUint32(off, 0x1B, true); dv2.setUint32(off + 4, 16, true);
  dv2.setInt32(off + 8, 10, true); dv2.setInt32(off + 12, 40, true);
  off += 16;

  // PolyLineTo16 (type=0x59, size=28+nPts*4): bounds(16) + count(4) + 2 pts (8)
  // bounds: L=50,T=10,R=90,B=40, count=2, pts: (50,10), (90,40)
  const nPts = 2;
  const plySize = 28 + nPts * 4;
  dv2.setUint32(off, 0x59, true); dv2.setUint32(off + 4, plySize, true);
  dv2.setInt32(off + 8, 10, true);   // bounds L
  dv2.setInt32(off + 12, 10, true);  // bounds T
  dv2.setInt32(off + 16, 90, true);  // bounds R
  dv2.setInt32(off + 20, 40, true);  // bounds B
  dv2.setUint32(off + 24, nPts, true);
  dv2.setInt16(off + 28, 50, true); dv2.setInt16(off + 30, 10, true);
  dv2.setInt16(off + 32, 90, true); dv2.setInt16(off + 34, 40, true);
  off += plySize;

  // CloseFigure (type=0x3D, size=8)
  dv2.setUint32(off, 0x3D, true); dv2.setUint32(off + 4, 8, true);
  off += 8;

  // EndPath (type=0x3C, size=8)
  dv2.setUint32(off, 0x3C, true); dv2.setUint32(off + 4, 8, true);
  off += 8;

  // FillPath (type=0x3E, size=24): bounds(16)
  dv2.setUint32(off, 0x3E, true); dv2.setUint32(off + 4, 24, true);
  dv2.setInt32(off + 8, 10, true); dv2.setInt32(off + 12, 10, true);
  dv2.setInt32(off + 16, 90, true); dv2.setInt32(off + 20, 40, true);
  off += 24;

  // EOF (type=0x0E, size=20)
  dv2.setUint32(off, 0x0E, true); dv2.setUint32(off + 4, 20, true);
  off += 20;

  return buf2.subarray(0, off);
}

function buildTestEmfWithWindowExtent() {
  // Header + SetWindowOrgEx + SetWindowExtEx + CreateBrush + SelectObj +
  // BeginPath + MoveToEx + PolyLineTo16(1pt) + CloseFigure + EndPath + FillPath + EOF
  const buf = new Uint8Array(300);
  const dv = new DataView(buf.buffer);

  // Header (108) — device bounds: 100x50
  dv.setUint32(0, 1, true); dv.setUint32(4, 108, true);
  dv.setInt32(8, 0, true); dv.setInt32(12, 0, true);
  dv.setInt32(16, 100, true); dv.setInt32(20, 50, true);
  dv.setInt32(24, 0, true); dv.setInt32(28, 0, true);
  dv.setInt32(32, 25400, true); dv.setInt32(36, 12700, true);
  dv.setUint32(40, 0x464D4520, true);
  dv.setUint32(44, 0x10000, true);
  dv.setUint32(48, 300, true); dv.setUint32(52, 12, true);
  let off = 108;

  // SetWindowOrgEx (type=0x0A, size=16): x=0, y=0
  dv.setUint32(off, 0x0A, true); dv.setUint32(off + 4, 16, true);
  dv.setInt32(off + 8, 0, true); dv.setInt32(off + 12, 0, true);
  off += 16;

  // SetWindowExtEx (type=0x09, size=16): cx=200, cy=100
  dv.setUint32(off, 0x09, true); dv.setUint32(off + 4, 16, true);
  dv.setInt32(off + 8, 200, true); dv.setInt32(off + 12, 100, true);
  off += 16;

  // CreateBrushIndirect (type=0x27, size=24): solid blue
  dv.setUint32(off, 0x27, true); dv.setUint32(off + 4, 24, true);
  dv.setUint32(off + 8, 0, true);
  dv.setUint32(off + 12, 0, true); // BS_SOLID
  dv.setUint8(off + 16, 0x00); dv.setUint8(off + 17, 0x00); dv.setUint8(off + 18, 0xFF);
  off += 24;

  // SelectObject (type=0x25, size=12)
  dv.setUint32(off, 0x25, true); dv.setUint32(off + 4, 12, true);
  dv.setUint32(off + 8, 0, true);
  off += 12;

  // BeginPath
  dv.setUint32(off, 0x3B, true); dv.setUint32(off + 4, 8, true);
  off += 8;

  // MoveToEx: (10, 10)
  dv.setUint32(off, 0x1B, true); dv.setUint32(off + 4, 16, true);
  dv.setInt32(off + 8, 10, true); dv.setInt32(off + 12, 10, true);
  off += 16;

  // PolyLineTo16: 1 point (190, 90)
  dv.setUint32(off, 0x59, true); dv.setUint32(off + 4, 32, true);
  dv.setInt32(off + 8, 10, true); dv.setInt32(off + 12, 10, true);
  dv.setInt32(off + 16, 190, true); dv.setInt32(off + 20, 90, true);
  dv.setUint32(off + 24, 1, true);
  dv.setInt16(off + 28, 190, true); dv.setInt16(off + 30, 90, true);
  off += 32;

  // CloseFigure
  dv.setUint32(off, 0x3D, true); dv.setUint32(off + 4, 8, true);
  off += 8;

  // EndPath
  dv.setUint32(off, 0x3C, true); dv.setUint32(off + 4, 8, true);
  off += 8;

  // FillPath (24)
  dv.setUint32(off, 0x3E, true); dv.setUint32(off + 4, 24, true);
  off += 24;

  // EOF (20)
  dv.setUint32(off, 0x0E, true); dv.setUint32(off + 4, 20, true);
  off += 20;

  return buf.subarray(0, off);
}
