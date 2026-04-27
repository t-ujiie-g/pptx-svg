/**
 * Lightweight EMF (Enhanced Metafile) to SVG converter.
 *
 * Handles the subset of EMF records commonly used by PowerPoint:
 * vector paths (bezier curves, lines), brushes, pens, clip paths,
 * text output, and embedded bitmaps.
 *
 * Falls back gracefully — returns empty string if parsing fails.
 */

// ── EMF Record Types (from MS-EMF specification) ────────────────────────────

const EMR_HEADER = 0x01;
const EMR_POLYGON = 0x03;
const EMR_POLYLINE = 0x04;
const EMR_POLYBEZIERTO = 0x05;
const EMR_POLYLINETO = 0x06;
const EMR_SETWINDOWEXTEX = 0x09;
const EMR_SETWINDOWORGEX = 0x0A;
const EMR_EOF = 0x0E;
const EMR_SETPOLYFILLMODE = 0x13;
const EMR_SETTEXTCOLOR = 0x18;
const EMR_SETBKCOLOR = 0x19;
const EMR_MOVETOEX = 0x1B;
const EMR_SAVEDC = 0x21;
const EMR_RESTOREDC = 0x22;
const EMR_SELECTOBJECT = 0x25;
const EMR_CREATEPEN = 0x26;
const EMR_CREATEBRUSHINDIRECT = 0x27;
const EMR_DELETEOBJECT = 0x28;
const EMR_ELLIPSE = 0x2A;
const EMR_RECTANGLE = 0x2B;
const EMR_LINETO = 0x36;
const EMR_BEGINPATH = 0x3B;
const EMR_ENDPATH = 0x3C;
const EMR_CLOSEFIGURE = 0x3D;
const EMR_FILLPATH = 0x3E;
const EMR_STROKEANDFILLPATH = 0x3F;
const EMR_STROKEPATH = 0x40;
const EMR_SELECTCLIPPATH = 0x43;
const EMR_GDICOMMENT = 0x46;
const EMR_STRETCHDIBITS = 0x51;
const EMR_EXTCREATEFONTINDIRECTW = 0x52;
const EMR_EXTTEXTOUTW = 0x54;
const EMR_POLYGON16 = 0x56;
const EMR_POLYLINE16 = 0x57;
const EMR_POLYBEZIERTO16 = 0x58;
const EMR_POLYLINETO16 = 0x59;
const EMR_EXTCREATEPEN = 0x5F;

// Stock objects (0x80000000 + index)
const STOCK_OBJECT_FLAG = 0x80000000;
// Stock brush indices
const WHITE_BRUSH = 0;
const NULL_BRUSH = 5;
// Stock pen indices
const NULL_PEN = 8;

// ── GDI State ───────────────────────────────────────────────────────────────

interface GdiPen {
  kind: 'pen';
  width: number;
  color: string;
  style: number; // 0=solid, 5=null
}

interface GdiBrush {
  kind: 'brush';
  style: number; // 0=solid, 1=null/hollow
  color: string;
}

interface GdiFont {
  kind: 'font';
  height: number;
  weight: number;
  italic: boolean;
  faceName: string;
}

type GdiObject = GdiPen | GdiBrush | GdiFont;

interface GdiState {
  textColor: string;
  bgColor: string;
  pen: GdiPen;
  brush: GdiBrush;
  font: GdiFont;
  curX: number;
  curY: number;
  fillMode: number; // 1=alternate(evenodd), 2=winding
}

function defaultPen(): GdiPen {
  return { kind: 'pen', width: 1, color: '#000000', style: 0 };
}

function nullPen(): GdiPen {
  return { kind: 'pen', width: 0, color: 'none', style: 5 };
}

function defaultBrush(): GdiBrush {
  return { kind: 'brush', style: 0, color: '#ffffff' };
}

function nullBrush(): GdiBrush {
  return { kind: 'brush', style: 1, color: 'none' };
}

function defaultFont(): GdiFont {
  return { kind: 'font', height: 16, weight: 400, italic: false, faceName: 'Arial' };
}

function cloneState(s: GdiState): GdiState {
  return {
    textColor: s.textColor,
    bgColor: s.bgColor,
    pen: { ...s.pen },
    brush: { ...s.brush },
    font: { ...s.font },
    curX: s.curX,
    curY: s.curY,
    fillMode: s.fillMode,
  };
}

// ── Helpers ─────────────────────────────────────────────────────────────────

function colorFromBGR(dv: DataView, offset: number): string {
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
 * Convert EMF binary data to an SVG string.
 * Returns empty string if the data cannot be parsed.
 */
export function emfToSvg(data: Uint8Array): string {
  try {
    return parseEmf(data);
  } catch {
    return '';
  }
}

function parseEmf(data: Uint8Array): string {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.length < 88) return '';

  // Verify header
  const recType = dv.getUint32(0, true);
  if (recType !== EMR_HEADER) return '';

  // Bounds (in device units)
  const boundsL = dv.getInt32(8, true);
  const boundsT = dv.getInt32(12, true);
  const boundsR = dv.getInt32(16, true);
  const boundsB = dv.getInt32(20, true);

  const width = boundsR - boundsL;
  const height = boundsB - boundsT;
  if (width <= 0 || height <= 0) return '';

  const headerSize = dv.getUint32(4, true);

  // GDI objects table
  const objects = new Map<number, GdiObject>();
  const state: GdiState = {
    textColor: '#000000',
    bgColor: '#ffffff',
    pen: defaultPen(),
    brush: defaultBrush(),
    font: defaultFont(),
    curX: 0,
    curY: 0,
    fillMode: 1,
  };
  const stateStack: GdiState[] = [];

  // Window (logical coordinate space) — overrides bounds if set
  let winOrgX = boundsL, winOrgY = boundsT;
  let winExtCx = width, winExtCy = height;

  // Path building
  let pathData = '';
  let inPath = false;

  // SVG output fragments
  const svgParts: string[] = [];

  let offset = headerSize;

  while (offset + 8 <= data.length) {
    const type = dv.getUint32(offset, true);
    const size = dv.getUint32(offset + 4, true);
    if (size < 8 || offset + size > data.length) break;
    if (type === EMR_EOF) break;

    switch (type) {
      case EMR_SETWINDOWORGEX:
        winOrgX = dv.getInt32(offset + 8, true);
        winOrgY = dv.getInt32(offset + 12, true);
        break;

      case EMR_SETWINDOWEXTEX:
        winExtCx = dv.getInt32(offset + 8, true);
        winExtCy = dv.getInt32(offset + 12, true);
        break;

      case EMR_SETTEXTCOLOR:
        state.textColor = colorFromBGR(dv, offset + 8);
        break;

      case EMR_SETBKCOLOR:
        state.bgColor = colorFromBGR(dv, offset + 8);
        break;

      case EMR_SETPOLYFILLMODE:
        state.fillMode = dv.getUint32(offset + 8, true);
        break;

      case EMR_MOVETOEX:
        state.curX = dv.getInt32(offset + 8, true);
        state.curY = dv.getInt32(offset + 12, true);
        if (inPath) pathData += `M${state.curX} ${state.curY} `;
        break;

      case EMR_LINETO: {
        const lx = dv.getInt32(offset + 8, true);
        const ly = dv.getInt32(offset + 12, true);
        if (inPath) {
          pathData += `L${lx} ${ly} `;
        } else if (state.pen.style !== 5) {
          svgParts.push(
            `<line x1="${state.curX}" y1="${state.curY}" x2="${lx}" y2="${ly}" ` +
            `stroke="${state.pen.color}" stroke-width="${state.pen.width}" />`
          );
        }
        state.curX = lx;
        state.curY = ly;
        break;
      }

      case EMR_RECTANGLE: {
        const rl = dv.getInt32(offset + 8, true);
        const rt = dv.getInt32(offset + 12, true);
        const rr = dv.getInt32(offset + 16, true);
        const rb = dv.getInt32(offset + 20, true);
        const fill = state.brush.style === 1 ? 'none' : state.brush.color;
        const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
        svgParts.push(
          `<rect x="${rl}" y="${rt}" width="${rr - rl}" height="${rb - rt}" ` +
          `fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
        );
        break;
      }

      case EMR_ELLIPSE: {
        const el = dv.getInt32(offset + 8, true);
        const et = dv.getInt32(offset + 12, true);
        const er = dv.getInt32(offset + 16, true);
        const eb = dv.getInt32(offset + 20, true);
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

      case EMR_POLYGON:
      case EMR_POLYGON16: {
        const pts = readPoints(dv, offset, type === EMR_POLYGON16, offset + size);
        if (pts.length > 0) {
          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          const d = pts.map((p, i) => `${i === 0 ? 'M' : 'L'}${p[0]} ${p[1]}`).join(' ') + ' Z';
          svgParts.push(
            `<path d="${d}" fill="${fill}" stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        break;
      }

      case EMR_POLYLINE:
      case EMR_POLYLINE16: {
        const pts = readPoints(dv, offset, type === EMR_POLYLINE16, offset + size);
        if (pts.length > 0) {
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          const d = pts.map((p, i) => `${i === 0 ? 'M' : 'L'}${p[0]} ${p[1]}`).join(' ');
          svgParts.push(
            `<path d="${d}" fill="none" stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        break;
      }

      case EMR_POLYBEZIERTO:
      case EMR_POLYBEZIERTO16: {
        const pts = readPoints(dv, offset, type === EMR_POLYBEZIERTO16, offset + size);
        if (inPath && pts.length >= 3) {
          for (let i = 0; i + 2 < pts.length; i += 3) {
            pathData += `C${pts[i][0]} ${pts[i][1]} ${pts[i + 1][0]} ${pts[i + 1][1]} ${pts[i + 2][0]} ${pts[i + 2][1]} `;
          }
          if (pts.length > 0) {
            state.curX = pts[pts.length - 1][0];
            state.curY = pts[pts.length - 1][1];
          }
        }
        break;
      }

      case EMR_POLYLINETO:
      case EMR_POLYLINETO16: {
        const pts = readPoints(dv, offset, type === EMR_POLYLINETO16, offset + size);
        if (inPath) {
          for (const p of pts) pathData += `L${p[0]} ${p[1]} `;
          if (pts.length > 0) {
            state.curX = pts[pts.length - 1][0];
            state.curY = pts[pts.length - 1][1];
          }
        }
        break;
      }

      case EMR_BEGINPATH:
        inPath = true;
        pathData = '';
        break;

      case EMR_ENDPATH:
        inPath = false;
        break;

      case EMR_CLOSEFIGURE:
        if (inPath) pathData += 'Z ';
        break;

      case EMR_FILLPATH: {
        if (pathData) {
          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const fillRule = state.fillMode === 1 ? ' fill-rule="evenodd"' : '';
          svgParts.push(`<path d="${pathData.trim()}" fill="${fill}"${fillRule} stroke="none" />`);
        }
        pathData = '';
        break;
      }

      case EMR_STROKEPATH:
        if (pathData) {
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          svgParts.push(
            `<path d="${pathData.trim()}" fill="none" stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        pathData = '';
        break;

      case EMR_STROKEANDFILLPATH: {
        if (pathData) {
          const fill = state.brush.style === 1 ? 'none' : state.brush.color;
          const stroke = state.pen.style === 5 ? 'none' : state.pen.color;
          const fillRule = state.fillMode === 1 ? ' fill-rule="evenodd"' : '';
          svgParts.push(
            `<path d="${pathData.trim()}" fill="${fill}"${fillRule} stroke="${stroke}" stroke-width="${state.pen.width}" />`
          );
        }
        pathData = '';
        break;
      }

      case EMR_SELECTCLIPPATH:
        // Discard clip path data (we skip clipping for simplicity)
        pathData = '';
        break;

      case EMR_CREATEPEN: {
        const ihPen = dv.getUint32(offset + 8, true);
        const penStyle = dv.getUint32(offset + 12, true);
        const penWidth = dv.getInt32(offset + 16, true);
        const penColor = colorFromBGR(dv, offset + 24);
        objects.set(ihPen, { kind: 'pen', width: Math.max(1, penWidth), color: penColor, style: penStyle & 0xf });
        break;
      }

      case EMR_EXTCREATEPEN: {
        const ihPen = dv.getUint32(offset + 8, true);
        // +12: offBmi(4), +16: cbBmi(4), +20: offBits(4), +24: cbBits(4)
        // +28: penStyle(4), +32: width(4), +36: brushStyle(4), +40: color(4)
        const penStyle = dv.getUint32(offset + 28, true);
        const penWidth = dv.getUint32(offset + 32, true);
        const penColor = colorFromBGR(dv, offset + 40);
        objects.set(ihPen, { kind: 'pen', width: Math.max(1, penWidth), color: penColor, style: penStyle & 0xf });
        break;
      }

      case EMR_CREATEBRUSHINDIRECT: {
        const ihBrush = dv.getUint32(offset + 8, true);
        const brushStyle = dv.getUint32(offset + 12, true);
        const brushColor = colorFromBGR(dv, offset + 16);
        objects.set(ihBrush, { kind: 'brush', style: brushStyle, color: brushColor });
        break;
      }

      case EMR_EXTCREATEFONTINDIRECTW: {
        const ihFont = dv.getUint32(offset + 8, true);
        // LOGFONT starts at offset + 12
        const fontHeight = Math.abs(dv.getInt32(offset + 12, true));
        // +16: width(4), +20: escapement(4), +24: orientation(4), +28: weight(4)
        const fontWeight = dv.getUint32(offset + 28, true);
        const fontItalic = dv.getUint8(offset + 32) !== 0;
        // Face name starts at +40 (after lfOutPrecision, lfClipPrecision, lfQuality, lfPitchAndFamily)
        // Actually: lfItalic(1), lfUnderline(1), lfStrikeOut(1), lfCharSet(1) = 4 bytes at +32
        // lfOutPrecision(1), lfClipPrecision(1), lfQuality(1), lfPitchAndFamily(1) = 4 bytes at +36
        // lfFaceName starts at +40, 32 UTF-16LE chars
        let faceName = '';
        const nameStart = offset + 40;
        for (let i = 0; i < 32; i++) {
          if (nameStart + i * 2 + 1 >= data.length) break;
          const ch = dv.getUint16(nameStart + i * 2, true);
          if (ch === 0) break;
          faceName += String.fromCharCode(ch);
        }
        objects.set(ihFont, { kind: 'font', height: fontHeight, weight: fontWeight, italic: fontItalic, faceName: faceName || 'Arial' });
        break;
      }

      case EMR_SELECTOBJECT: {
        const ihObj = dv.getUint32(offset + 8, true);
        if (ihObj & STOCK_OBJECT_FLAG) {
          const stockIdx = ihObj & 0x7fffffff;
          if (stockIdx === WHITE_BRUSH) state.brush = { kind: 'brush', style: 0, color: '#ffffff' };
          else if (stockIdx === NULL_BRUSH) state.brush = nullBrush();
          else if (stockIdx === NULL_PEN) state.pen = nullPen();
          else if (stockIdx <= 4) {
            // BLACK_BRUSH=4, DKGRAY=3, GRAY=2, LTGRAY=1
            const colors = ['#ffffff', '#c0c0c0', '#808080', '#404040', '#000000'];
            state.brush = { kind: 'brush', style: 0, color: colors[stockIdx] };
          } else if (stockIdx === 6) state.pen = { kind: 'pen', width: 1, color: '#ffffff', style: 0 };
          else if (stockIdx === 7) state.pen = defaultPen();
        } else {
          const obj = objects.get(ihObj);
          if (obj) {
            if (obj.kind === 'pen') state.pen = obj;
            else if (obj.kind === 'brush') state.brush = obj;
            else if (obj.kind === 'font') state.font = obj;
          }
        }
        break;
      }

      case EMR_DELETEOBJECT: {
        const ihDel = dv.getUint32(offset + 8, true);
        objects.delete(ihDel);
        break;
      }

      case EMR_SAVEDC:
        stateStack.push(cloneState(state));
        break;

      case EMR_RESTOREDC:
        if (stateStack.length > 0) {
          Object.assign(state, stateStack.pop()!);
        }
        break;

      case EMR_EXTTEXTOUTW: {
        // +8: bounds(16), +24: iGraphicsMode(4), +28: exScale(4), +32: eyScale(4)
        // +36: EmrText: ptlReference(8), +44: nChars(4), +48: offString(4),
        // +52: fOptions(4), +56: rcl(16), +72: offDx(4)
        const refX = dv.getInt32(offset + 36, true);
        const refY = dv.getInt32(offset + 40, true);
        const nChars = dv.getUint32(offset + 44, true);
        const offString = dv.getUint32(offset + 48, true);

        if (nChars > 0 && offString > 0 && offset + offString + nChars * 2 <= data.length) {
          let text = '';
          for (let i = 0; i < nChars; i++) {
            const ch = dv.getUint16(offset + offString + i * 2, true);
            if (ch === 0) break;
            text += String.fromCharCode(ch);
          }
          if (text.trim()) {
            const fontSize = state.font.height;
            const fontWeight = state.font.weight >= 700 ? 'bold' : 'normal';
            const fontStyle = state.font.italic ? 'italic' : 'normal';
            const fontFamily = escapeXml(state.font.faceName);
            svgParts.push(
              `<text x="${refX}" y="${refY}" ` +
              `font-family="${fontFamily}" font-size="${fontSize}" ` +
              `font-weight="${fontWeight}" font-style="${fontStyle}" ` +
              `fill="${state.textColor}">${escapeXml(text)}</text>`
            );
          }
        }
        break;
      }

      case EMR_STRETCHDIBITS: {
        // +8: bounds(16), +24: xDest(4), +28: yDest(4), +32: xSrc(4), +36: ySrc(4)
        // +40: cxSrc(4), +44: cySrc(4), +48: offBmiSrc(4), +52: cbBmiSrc(4)
        // +56: offBitsSrc(4), +60: cbBitsSrc(4), +64: iUsageSrc(4), +68: dwRop(4)
        // +72: cxDest(4), +76: cyDest(4)
        const xDest = dv.getInt32(offset + 24, true);
        const yDest = dv.getInt32(offset + 28, true);
        const cxDest = dv.getInt32(offset + 72, true);
        const cyDest = dv.getInt32(offset + 76, true);
        const offBmi = dv.getUint32(offset + 48, true);
        const cbBmi = dv.getUint32(offset + 52, true);
        const offBits = dv.getUint32(offset + 56, true);
        const cbBits = dv.getUint32(offset + 60, true);

        // Validate offsets/lengths fit within this record (defends against
        // malicious EMF where cbBmi/cbBits sum could allocate gigabytes).
        const recEnd = offset + size;
        const bmiEnd = offset + offBmi + cbBmi;
        const bitsEnd = offset + offBits + cbBits;
        if (
          cbBmi > 0 && cbBits > 0 && cxDest > 0 && cyDest > 0 &&
          offBmi >= 88 && offBits >= 88 &&
          bmiEnd <= recEnd && bitsEnd <= recEnd &&
          bmiEnd >= offset + offBmi && bitsEnd >= offset + offBits   // overflow guard
        ) {
          const bmpSize = 14 + cbBmi + cbBits;
          const bmp = new Uint8Array(bmpSize);
          const bmpDv = new DataView(bmp.buffer);
          bmp[0] = 0x42; bmp[1] = 0x4D; // "BM"
          bmpDv.setUint32(2, bmpSize, true);
          bmpDv.setUint32(10, 14 + cbBmi, true);
          bmp.set(data.subarray(offset + offBmi, bmiEnd), 14);
          bmp.set(data.subarray(offset + offBits, bitsEnd), 14 + cbBmi);
          let binary = '';
          for (let i = 0; i < bmp.length; i++) binary += String.fromCharCode(bmp[i]);
          const b64 = btoa(binary);
          svgParts.push(
            `<image x="${xDest}" y="${yDest}" width="${cxDest}" height="${cyDest}" ` +
            `href="data:image/bmp;base64,${b64}" />`
          );
        }
        break;
      }

      case EMR_GDICOMMENT:
        // Skip GDI comments (may contain EMF+ or embedded PDFs)
        break;
    }

    offset += size;
  }

  if (svgParts.length === 0) return '';

  // Use window extent (logical coords) for viewBox; device bounds for width/height
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${winOrgX} ${winOrgY} ${winExtCx} ${winExtCy}" ` +
    `width="${width}" height="${height}" preserveAspectRatio="xMidYMid meet">` +
    svgParts.join('') +
    `</svg>`
  );
}

// ── Point readers ───────────────────────────────────────────────────────────

/**
 * Hard cap on points per EMF record.
 * Guards against malicious EMF where `count` is set to ~2^32, which would
 * cause a multi-billion-iteration loop and OOM on the points array.
 */
const MAX_POINTS_PER_RECORD = 100_000;

/** Read points from polygon/polyline/polybezier records (with bounds + count header). */
function readPoints(dv: DataView, offset: number, is16bit: boolean, recordEnd: number): number[][] {
  // +8: bounds (16 bytes), +24: count (4 bytes), +28: points
  const declared = dv.getUint32(offset + 24, true);
  const ptSize = is16bit ? 4 : 8;
  // Clamp by declared cap, by record-size capacity, and by absolute maximum.
  const ptOff = offset + 28;
  const capacity = Math.max(0, Math.floor((recordEnd - ptOff) / ptSize));
  const count = Math.min(declared, capacity, MAX_POINTS_PER_RECORD);
  const pts: number[][] = [];
  let p = ptOff;
  for (let i = 0; i < count; i++) {
    if (is16bit) {
      pts.push([dv.getInt16(p, true), dv.getInt16(p + 2, true)]);
      p += 4;
    } else {
      pts.push([dv.getInt32(p, true), dv.getInt32(p + 4, true)]);
      p += 8;
    }
  }
  return pts;
}
