/**
 * Lightweight WMF (Windows Metafile) to SVG converter.
 *
 * Handles the subset of WMF records commonly used by PowerPoint:
 * vector paths (lines, polygons, polylines), brushes, pens, text output,
 * and embedded bitmaps (StretchDIBits).
 *
 * Supports both Standard and Placeable WMF formats.
 *
 * Falls back gracefully — returns empty string if parsing fails.
 */

// ── WMF Record Types (from MS-WMF specification) ────────────────────────────

const META_EOF                  = 0x0000;
const META_SETBKCOLOR           = 0x0201;
const META_SETBKMODE            = 0x0102;
const META_SETMAPMODE           = 0x0103;
const META_SETPOLYFILLMODE      = 0x0106;
const META_SETTEXTCOLOR         = 0x0209;
const META_SETWINDOWORG         = 0x020B;
const META_SETWINDOWEXT         = 0x020C;
const META_MOVETO               = 0x0214;
const META_LINETO               = 0x0213;
const META_RECTANGLE            = 0x041B;
const META_ROUNDRECT            = 0x061C;
const META_ELLIPSE              = 0x0418;
const META_POLYGON              = 0x0324;
const META_POLYLINE             = 0x0325;
const META_POLYPOLYGON          = 0x0538;
const META_ARC                  = 0x0817;
const META_PIE                  = 0x081A;
const META_CHORD                = 0x0830;
const META_CREATEPENINDIRECT    = 0x02FA;
const META_CREATEBRUSHINDIRECT  = 0x02FC;
const META_CREATEFONTINDIRECT   = 0x02FB;
const META_SELECTOBJECT         = 0x012D;
const META_DELETEOBJECT         = 0x01F0;
const META_SAVEDC               = 0x001E;
const META_RESTOREDC            = 0x0127;
const META_EXTTEXTOUT           = 0x0A32;
const META_TEXTOUT              = 0x0521;
const META_STRETCHDIB           = 0x0F43;
const META_DIBSTRETCHBLT        = 0x0B41;
const META_ESCAPE               = 0x0626;

// Placeable WMF magic
const PLACEABLE_KEY = 0x9AC6CDD7;

// ── GDI State (reused types from emf-converter) ─────────────────────────────

interface WmfPen {
  style: number;  // 0=solid, 1=dash, 5=null
  width: number;
  color: string;
}

interface WmfBrush {
  style: number;  // 0=solid, 1=null/hollow
  color: string;
}

interface WmfFont {
  height: number;
  weight: number;
  italic: boolean;
  faceName: string;
}

type WmfObject = WmfPen | WmfBrush | WmfFont;

interface WmfState {
  textColor: string;
  bgColor: string;
  bgMode: number;   // 1=TRANSPARENT, 2=OPAQUE
  pen: WmfPen;
  brush: WmfBrush;
  font: WmfFont;
  curX: number;
  curY: number;
  fillMode: number;  // 1=alternate(evenodd), 2=winding
}

function defaultPen(): WmfPen {
  return { style: 0, width: 1, color: '#000000' };
}

function defaultBrush(): WmfBrush {
  return { style: 0, color: '#ffffff' };
}

function defaultFont(): WmfFont {
  return { height: 16, weight: 400, italic: false, faceName: 'Arial' };
}

function cloneState(s: WmfState): WmfState {
  return {
    textColor: s.textColor,
    bgColor: s.bgColor,
    bgMode: s.bgMode,
    pen: { ...s.pen },
    brush: { ...s.brush },
    font: { ...s.font },
    curX: s.curX,
    curY: s.curY,
    fillMode: s.fillMode,
  };
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/** Read a WMF COLORREF (4 bytes: R, G, B, reserved) */
function colorFromRGB(dv: DataView, offset: number): string {
  const r = dv.getUint8(offset);
  const g = dv.getUint8(offset + 1);
  const b = dv.getUint8(offset + 2);
  return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
}

function hex2(n: number): string {
  return n.toString(16).padStart(2, '0');
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ── Main converter ──────────────────────────────────────────────────────────

/**
 * Convert WMF binary data to an SVG string.
 * Returns empty string if the data cannot be parsed.
 */
export function wmfToSvg(data: Uint8Array): string {
  try {
    return parseWmf(data);
  } catch {
    return '';
  }
}

function parseWmf(data: Uint8Array): string {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.length < 18) return '';

  let offset = 0;

  // ── Detect Placeable WMF ──
  let bboxL = 0, bboxT = 0, bboxR = 0, bboxB = 0;
  let unitsPerInch = 96;
  let isPlaceable = false;

  if (data.length >= 22 && dv.getUint32(0, true) === PLACEABLE_KEY) {
    isPlaceable = true;
    // Placeable WMF header (22 bytes)
    // +4: HWmf(2), +6: BboxLeft(2), +8: BboxTop(2), +10: BboxRight(2),
    // +12: BboxBottom(2), +14: Inch(2), +16: Reserved(4), +20: Checksum(2)
    bboxL = dv.getInt16(6, true);
    bboxT = dv.getInt16(8, true);
    bboxR = dv.getInt16(10, true);
    bboxB = dv.getInt16(12, true);
    unitsPerInch = dv.getUint16(14, true);
    if (unitsPerInch === 0) unitsPerInch = 96;
    offset = 22;
  }

  // ── Standard WMF header (METAHEADER, 18 bytes) ──
  if (offset + 18 > data.length) return '';
  const mtType = dv.getUint16(offset, true);          // 1=memory, 2=disk
  const mtHeaderSize = dv.getUint16(offset + 2, true); // Size in words (should be 9)
  // const mtVersion = dv.getUint16(offset + 4, true);
  // const mtSize = dv.getUint32(offset + 6, true);       // Total size in words
  const mtNoObjects = dv.getUint16(offset + 10, true);

  if ((mtType !== 1 && mtType !== 2) || mtHeaderSize < 9) return '';

  const headerBytes = mtHeaderSize * 2;
  offset += headerBytes;

  // ── GDI state ──
  const objects = new Map<number, WmfObject>();
  // Object table: next available slot
  let nextObjIdx = 0;
  const maxObjects = mtNoObjects || 256;

  const state: WmfState = {
    textColor: '#000000',
    bgColor: '#ffffff',
    bgMode: 2,
    pen: defaultPen(),
    brush: defaultBrush(),
    font: defaultFont(),
    curX: 0,
    curY: 0,
    fillMode: 1,
  };
  const stateStack: WmfState[] = [];

  // Window/viewport coordinates
  let winOrgX = isPlaceable ? bboxL : 0;
  let winOrgY = isPlaceable ? bboxT : 0;
  let winExtCx = isPlaceable ? (bboxR - bboxL) : 0;
  let winExtCy = isPlaceable ? (bboxB - bboxT) : 0;

  const svgParts: string[] = [];

  // ── Record loop ──
  while (offset + 6 <= data.length) {
    const recSizeWords = dv.getUint32(offset, true);
    const recFunc = dv.getUint16(offset + 4, true);
    const recSizeBytes = recSizeWords * 2;

    if (recSizeBytes < 6 || offset + recSizeBytes > data.length) break;
    if (recFunc === META_EOF) break;

    // Parameters start at offset + 6
    const p = offset + 6;

    switch (recFunc) {
      case META_SETWINDOWORG:
        // Parameters: Y(int16), X(int16) — NOTE: Y before X in WMF
        winOrgY = dv.getInt16(p, true);
        winOrgX = dv.getInt16(p + 2, true);
        break;

      case META_SETWINDOWEXT:
        winExtCy = dv.getInt16(p, true);
        winExtCx = dv.getInt16(p + 2, true);
        break;

      case META_SETMAPMODE:
        // We only handle MM_ANISOTROPIC/MM_ISOTROPIC via window org/ext
        // For other modes, the window org/ext values are used as-is
        break;

      case META_SETTEXTCOLOR:
        state.textColor = colorFromRGB(dv, p);
        break;

      case META_SETBKCOLOR:
        state.bgColor = colorFromRGB(dv, p);
        break;

      case META_SETBKMODE:
        state.bgMode = dv.getInt16(p, true);
        break;

      case META_SETPOLYFILLMODE:
        state.fillMode = dv.getInt16(p, true);
        break;

      case META_MOVETO:
        // Parameters: Y(int16), X(int16)
        state.curY = dv.getInt16(p, true);
        state.curX = dv.getInt16(p + 2, true);
        break;

      case META_LINETO: {
        const ly = dv.getInt16(p, true);
        const lx = dv.getInt16(p + 2, true);
        if (state.pen.style !== 5) {
          svgParts.push(
            `<line x1="${state.curX}" y1="${state.curY}" x2="${lx}" y2="${ly}" ` +
            `stroke="${state.pen.color}" stroke-width="${state.pen.width}" />`
          );
        }
        state.curX = lx;
        state.curY = ly;
        break;
      }

      case META_RECTANGLE: {
        // Parameters: Bottom(int16), Right(int16), Top(int16), Left(int16)
        const rb = dv.getInt16(p, true);
        const rr = dv.getInt16(p + 2, true);
        const rt = dv.getInt16(p + 4, true);
        const rl = dv.getInt16(p + 6, true);
        const fill = state.brush.style === 1 ? 'none' : state.brush.color;
        const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
        svgParts.push(
          `<rect x="${rl}" y="${rt}" width="${rr - rl}" height="${rb - rt}" ` +
          `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
        );
        break;
      }

      case META_ROUNDRECT: {
        // Parameters: ellipseH(int16), ellipseW(int16), Bottom, Right, Top, Left
        const ey = dv.getInt16(p, true);
        const ex = dv.getInt16(p + 2, true);
        const rb = dv.getInt16(p + 4, true);
        const rr = dv.getInt16(p + 6, true);
        const rt = dv.getInt16(p + 8, true);
        const rl = dv.getInt16(p + 10, true);
        const rx = Math.abs(ex) / 2;
        const ry = Math.abs(ey) / 2;
        const fill = state.brush.style === 1 ? 'none' : state.brush.color;
        const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
        svgParts.push(
          `<rect x="${rl}" y="${rt}" width="${rr - rl}" height="${rb - rt}" ` +
          `rx="${rx}" ry="${ry}" ` +
          `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
        );
        break;
      }

      case META_ELLIPSE: {
        // Parameters: Bottom(int16), Right(int16), Top(int16), Left(int16)
        const eb = dv.getInt16(p, true);
        const er = dv.getInt16(p + 2, true);
        const et = dv.getInt16(p + 4, true);
        const el = dv.getInt16(p + 6, true);
        const ecx = (el + er) / 2;
        const ecy = (et + eb) / 2;
        const erx = (er - el) / 2;
        const ery = (eb - et) / 2;
        const fill = state.brush.style === 1 ? 'none' : state.brush.color;
        const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
        svgParts.push(
          `<ellipse cx="${ecx}" cy="${ecy}" rx="${erx}" ry="${ery}" ` +
          `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
        );
        break;
      }

      case META_POLYGON: {
        // Parameters: NumberOfPoints(int16), Points[](int16 pairs)
        const declared = dv.getInt16(p, true);
        const recEnd = offset + recSizeBytes;
        const capacity = Math.max(0, Math.floor((recEnd - (p + 2)) / 4));
        const nPts = Math.min(Math.max(0, declared), capacity);
        if (nPts > 0) {
          let d = '';
          for (let i = 0; i < nPts; i++) {
            const px = dv.getInt16(p + 2 + i * 4, true);
            const py = dv.getInt16(p + 4 + i * 4, true);
            d += `${i === 0 ? 'M' : 'L'}${px} ${py} `;
          }
          d += 'Z';
          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          const fillRule = state.fillMode === 1 ? ' fill-rule="evenodd"' : '';
          svgParts.push(
            `<path d="${d.trim()}" fill="${fill}"${fillRule} stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        break;
      }

      case META_POLYLINE: {
        const declared = dv.getInt16(p, true);
        const recEnd = offset + recSizeBytes;
        const capacity = Math.max(0, Math.floor((recEnd - (p + 2)) / 4));
        const nPts = Math.min(Math.max(0, declared), capacity);
        if (nPts > 0) {
          let d = '';
          for (let i = 0; i < nPts; i++) {
            const px = dv.getInt16(p + 2 + i * 4, true);
            const py = dv.getInt16(p + 4 + i * 4, true);
            d += `${i === 0 ? 'M' : 'L'}${px} ${py} `;
          }
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          svgParts.push(
            `<path d="${d.trim()}" fill="none" stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        break;
      }

      case META_POLYPOLYGON: {
        // Parameters: NumberOfPolygons(uint16), PointCounts[](uint16), Points[]
        const recEnd = offset + recSizeBytes;
        const declaredPolys = dv.getUint16(p, true);
        const polyHeaderEnd = p + 2 + declaredPolys * 2;
        // Reject if the count table itself runs past the record.
        if (polyHeaderEnd > recEnd) break;
        const counts: number[] = [];
        let off = p + 2;
        for (let i = 0; i < declaredPolys; i++) {
          counts.push(dv.getUint16(off, true));
          off += 2;
        }
        // Compute total declared points and clamp to remaining record bytes / 4.
        let totalDeclared = 0;
        for (const c of counts) totalDeclared += c;
        const capacityPoints = Math.max(0, Math.floor((recEnd - off) / 4));
        if (totalDeclared > capacityPoints) break; // malformed/malicious — skip
        let d = '';
        for (let poly = 0; poly < declaredPolys; poly++) {
          const cnt = counts[poly];
          for (let i = 0; i < cnt; i++) {
            const px = dv.getInt16(off, true);
            const py = dv.getInt16(off + 2, true);
            d += `${i === 0 ? 'M' : 'L'}${px} ${py} `;
            off += 4;
          }
          d += 'Z ';
        }
        if (d) {
          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          const fillRule = state.fillMode === 1 ? ' fill-rule="evenodd"' : '';
          svgParts.push(
            `<path d="${d.trim()}" fill="${fill}"${fillRule} stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        break;
      }

      case META_ARC:
      case META_PIE:
      case META_CHORD: {
        // Parameters: yEnd(16), xEnd(16), yStart(16), xStart(16),
        //             Bottom(16), Right(16), Top(16), Left(16)
        const ye = dv.getInt16(p, true);
        const xe = dv.getInt16(p + 2, true);
        const ys = dv.getInt16(p + 4, true);
        const xs = dv.getInt16(p + 6, true);
        const ab = dv.getInt16(p + 8, true);
        const ar = dv.getInt16(p + 10, true);
        const at = dv.getInt16(p + 12, true);
        const al = dv.getInt16(p + 14, true);
        const acx = (al + ar) / 2;
        const acy = (at + ab) / 2;
        const arx = (ar - al) / 2;
        const ary = (ab - at) / 2;
        if (arx > 0 && ary > 0) {
          // Compute start/end angles from reference points
          const startAngle = Math.atan2((ys - acy) / ary, (xs - acx) / arx);
          const endAngle = Math.atan2((ye - acy) / ary, (xe - acx) / arx);
          const sx = acx + arx * Math.cos(startAngle);
          const sy = acy + ary * Math.sin(startAngle);
          const ex2 = acx + arx * Math.cos(endAngle);
          const ey2 = acy + ary * Math.sin(endAngle);
          // Determine large-arc flag
          let sweep = endAngle - startAngle;
          if (sweep < 0) sweep += Math.PI * 2;
          const largeArc = sweep > Math.PI ? 1 : 0;

          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;

          if (recFunc === META_ARC) {
            svgParts.push(
              `<path d="M${sx.toFixed(1)} ${sy.toFixed(1)} A${arx} ${ary} 0 ${largeArc} 1 ${ex2.toFixed(1)} ${ey2.toFixed(1)}" ` +
              `fill="none" stroke="${stroke}" stroke-width="${state.pen.width}" />`
            );
          } else if (recFunc === META_PIE) {
            svgParts.push(
              `<path d="M${acx} ${acy} L${sx.toFixed(1)} ${sy.toFixed(1)} A${arx} ${ary} 0 ${largeArc} 1 ${ex2.toFixed(1)} ${ey2.toFixed(1)} Z" ` +
              `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
            );
          } else {
            // CHORD
            svgParts.push(
              `<path d="M${sx.toFixed(1)} ${sy.toFixed(1)} A${arx} ${ary} 0 ${largeArc} 1 ${ex2.toFixed(1)} ${ey2.toFixed(1)} Z" ` +
              `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
            );
          }
        }
        break;
      }

      case META_CREATEPENINDIRECT: {
        // Parameters: PenStyle(int16), WidthX(int16), WidthY(int16), ColorRef(4)
        const penStyle = dv.getUint16(p, true);
        const penWidth = dv.getInt16(p + 2, true);
        // WidthY at p+4 ignored
        const penColor = colorFromRGB(dv, p + 6);
        const idx = allocateObject(objects, nextObjIdx, maxObjects);
        nextObjIdx = idx + 1;
        objects.set(idx, { style: penStyle & 0xf, width: Math.max(1, penWidth), color: penColor } as WmfPen);
        break;
      }

      case META_CREATEBRUSHINDIRECT: {
        // Parameters: BrushStyle(uint16), ColorRef(4), BrushHatch(int16)
        const brushStyle = dv.getUint16(p, true);
        const brushColor = colorFromRGB(dv, p + 2);
        const idx = allocateObject(objects, nextObjIdx, maxObjects);
        nextObjIdx = idx + 1;
        objects.set(idx, { style: brushStyle === 1 || brushStyle === 5 ? 1 : 0, color: brushColor } as WmfBrush);
        break;
      }

      case META_CREATEFONTINDIRECT: {
        // Parameters: LOGFONT structure (int16-based fields)
        // Height(16), Width(16), Escapement(16), Orientation(16), Weight(16),
        // Italic(8), Underline(8), StrikeOut(8), CharSet(8),
        // OutPrecision(8), ClipPrecision(8), Quality(8), PitchAndFamily(8),
        // FaceName (char[]) — rest of record
        const fontHeight = Math.abs(dv.getInt16(p, true));
        // Width at p+2 ignored
        // Escapement at p+4, Orientation at p+6 ignored
        const fontWeight = dv.getUint16(p + 8, true);
        const fontItalic = dv.getUint8(p + 10) !== 0;
        // Skip underline(+11), strikeout(+12), charset(+13),
        //   outprecision(+14), clipprecision(+15), quality(+16), pitch(+17)
        let faceName = '';
        const nameStart = p + 18;
        const nameEnd = offset + recSizeBytes;
        for (let i = nameStart; i < nameEnd; i++) {
          const ch = dv.getUint8(i);
          if (ch === 0) break;
          faceName += String.fromCharCode(ch);
        }
        const idx = allocateObject(objects, nextObjIdx, maxObjects);
        nextObjIdx = idx + 1;
        objects.set(idx, {
          height: fontHeight || 16,
          weight: fontWeight,
          italic: fontItalic,
          faceName: faceName || 'Arial',
        } as WmfFont);
        break;
      }

      case META_SELECTOBJECT: {
        const ihObj = dv.getUint16(p, true);
        const obj = objects.get(ihObj);
        if (obj) {
          if ('style' in obj && 'width' in obj) state.pen = obj as WmfPen;
          else if ('style' in obj && !('width' in obj)) state.brush = obj as WmfBrush;
          else if ('faceName' in obj) state.font = obj as WmfFont;
        }
        break;
      }

      case META_DELETEOBJECT: {
        const ihDel = dv.getUint16(p, true);
        objects.delete(ihDel);
        // Allow reuse of this slot
        if (ihDel < nextObjIdx) nextObjIdx = ihDel;
        break;
      }

      case META_SAVEDC:
        stateStack.push(cloneState(state));
        break;

      case META_RESTOREDC: {
        const nSaved = dv.getInt16(p, true);
        if (nSaved < 0) {
          // Relative: pop |nSaved| states, restore last
          const count = Math.min(Math.abs(nSaved), stateStack.length);
          let restored: WmfState | undefined;
          for (let i = 0; i < count; i++) restored = stateStack.pop();
          if (restored) Object.assign(state, restored);
        } else if (stateStack.length > 0) {
          Object.assign(state, stateStack.pop()!);
        }
        break;
      }

      case META_TEXTOUT: {
        // Parameters: StringLength(int16), String(bytes), Y(int16), X(int16)
        const nChars = dv.getInt16(p, true);
        if (nChars > 0) {
          let text = '';
          for (let i = 0; i < nChars; i++) {
            const ch = dv.getUint8(p + 2 + i);
            if (ch === 0) break;
            text += String.fromCharCode(ch);
          }
          // Y and X follow the string (padded to word boundary)
          const strBytes = nChars + (nChars % 2); // pad to even
          const ty = dv.getInt16(p + 2 + strBytes, true);
          const tx = dv.getInt16(p + 4 + strBytes, true);
          if (text.trim()) {
            const fontSize = state.font.height;
            const fontWeight = state.font.weight >= 700 ? 'bold' : 'normal';
            const fontStyle = state.font.italic ? 'italic' : 'normal';
            svgParts.push(
              `<text x="${tx}" y="${ty}" ` +
              `font-family="${escapeXml(state.font.faceName)}" font-size="${fontSize}" ` +
              `font-weight="${fontWeight}" font-style="${fontStyle}" ` +
              `fill="${state.textColor}">${escapeXml(text)}</text>`
            );
          }
        }
        break;
      }

      case META_EXTTEXTOUT: {
        // Parameters: Y(int16), X(int16), StringLength(int16), fwOpts(uint16),
        //             [Rect(8 bytes if ETO_CLIPPED/ETO_OPAQUE)], String, [Dx[]]
        const ety = dv.getInt16(p, true);
        const etx = dv.getInt16(p + 2, true);
        const nChars = dv.getInt16(p + 4, true);
        const fwOpts = dv.getUint16(p + 6, true);
        const hasRect = (fwOpts & 0x06) !== 0; // ETO_CLIPPED=4, ETO_OPAQUE=2
        const strOff = p + 8 + (hasRect ? 8 : 0);

        if (nChars > 0 && strOff + nChars <= offset + recSizeBytes) {
          let text = '';
          for (let i = 0; i < nChars; i++) {
            const ch = dv.getUint8(strOff + i);
            if (ch === 0) break;
            text += String.fromCharCode(ch);
          }
          if (text.trim()) {
            const fontSize = state.font.height;
            const fontWeight = state.font.weight >= 700 ? 'bold' : 'normal';
            const fontStyle = state.font.italic ? 'italic' : 'normal';
            svgParts.push(
              `<text x="${etx}" y="${ety}" ` +
              `font-family="${escapeXml(state.font.faceName)}" font-size="${fontSize}" ` +
              `font-weight="${fontWeight}" font-style="${fontStyle}" ` +
              `fill="${state.textColor}">${escapeXml(text)}</text>`
            );
          }
        }
        break;
      }

      case META_STRETCHDIB:
      case META_DIBSTRETCHBLT: {
        // META_STRETCHDIB:
        //   dwRop(4), ColorUsage(2), SrcH(2), SrcW(2), YSrc(2), XSrc(2),
        //   DestH(2), DestW(2), YDest(2), XDest(2), DeviceIndependentBitmap
        // META_DIBSTRETCHBLT (no color usage):
        //   dwRop(4), SrcH(2), SrcW(2), YSrc(2), XSrc(2),
        //   DestH(2), DestW(2), YDest(2), XDest(2), DeviceIndependentBitmap
        let dibParamOff = p;
        dibParamOff += 4; // skip dwRop
        if (recFunc === META_STRETCHDIB) dibParamOff += 2; // skip ColorUsage
        // SrcH, SrcW, YSrc, XSrc (skip)
        dibParamOff += 8;
        const destH = dv.getInt16(dibParamOff, true);
        const destW = dv.getInt16(dibParamOff + 2, true);
        const yDest = dv.getInt16(dibParamOff + 4, true);
        const xDest = dv.getInt16(dibParamOff + 6, true);
        dibParamOff += 8;

        const dibStart = dibParamOff;
        const dibEnd = offset + recSizeBytes;
        const dibSize = dibEnd - dibStart;

        if (dibSize > 12 && Math.abs(destW) > 0 && Math.abs(destH) > 0) {
          // Read BITMAPINFOHEADER size to split header from bits
          const biSize = dv.getUint32(dibStart, true);
          if (biSize >= 12 && biSize <= dibSize) {
            const biBitCount = dv.getUint16(dibStart + 14, true);
            const biClrUsed = biSize >= 36 ? dv.getUint32(dibStart + 32, true) : 0;
            // Calculate color table size
            let paletteEntries = biClrUsed;
            if (paletteEntries === 0 && biBitCount <= 8) {
              paletteEntries = 1 << biBitCount;
            }
            const paletteSize = paletteEntries * 4;
            const bmiSize = biSize + paletteSize;
            const bitsSize = dibSize - bmiSize;

            if (bitsSize > 0) {
              const bmpFileSize = 14 + dibSize;
              const bmp = new Uint8Array(bmpFileSize);
              const bmpDv = new DataView(bmp.buffer);
              bmp[0] = 0x42; bmp[1] = 0x4D; // "BM"
              bmpDv.setUint32(2, bmpFileSize, true);
              bmpDv.setUint32(10, 14 + bmiSize, true);
              bmp.set(data.subarray(dibStart, dibEnd), 14);

              let binary = '';
              for (let i = 0; i < bmp.length; i++) binary += String.fromCharCode(bmp[i]);
              const b64 = btoa(binary);

              const w = Math.abs(destW);
              const h = Math.abs(destH);
              svgParts.push(
                `<image x="${xDest}" y="${yDest}" width="${w}" height="${h}" ` +
                `href="data:image/bmp;base64,${b64}" />`
              );
            }
          }
        }
        break;
      }

      case META_ESCAPE:
        // Skip escape records (printer-specific)
        break;
    }

    offset += recSizeBytes;
  }

  if (svgParts.length === 0) return '';

  // Compute physical dimensions from placeable header or window extent
  let vbX = winOrgX, vbY = winOrgY;
  let vbW = winExtCx, vbH = winExtCy;
  if (vbW <= 0 || vbH <= 0) {
    // Fallback if window extent was never set
    vbX = 0; vbY = 0; vbW = 1000; vbH = 1000;
  }

  // Physical size in pixels (from placeable header's inch value)
  const pxW = Math.round(vbW * 96 / unitsPerInch);
  const pxH = Math.round(vbH * 96 / unitsPerInch);

  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vbX} ${vbY} ${vbW} ${vbH}" ` +
    `width="${pxW}" height="${pxH}" preserveAspectRatio="xMidYMid meet">` +
    svgParts.join('') +
    `</svg>`
  );
}

/** Find the lowest available object table index. */
function allocateObject(objects: Map<number, WmfObject>, hint: number, max: number): number {
  for (let i = hint; i < max; i++) {
    if (!objects.has(i)) return i;
  }
  // Fallback: use hint anyway (may overwrite)
  return hint;
}
