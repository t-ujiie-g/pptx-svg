
/** EMU per point (1 pt = 12700 EMU). */
export const EMU_PER_PT = 12700;

/** EMU per pixel at 96 DPI (914400 / 96 = 9525). */
export const EMU_PER_PX_96DPI = 9525;

/** Convert pixels to EMU. */
export function pxToEmu(px: number, dpi = 96): number {
  return Math.round(px * 914400 / dpi);
}

/** Convert EMU to pixels. */
export function emuToPx(emu: number, dpi = 96): number {
  return Math.round(emu * dpi / 914400);
}

/** Convert points to OOXML hundredths of a point (18 → 1800). */
export function ptToHundredths(pt: number): number {
  return Math.round(pt * 100);
}

/** Convert OOXML hundredths of a point to points (1800 → 18). */
export function hundredthsToPt(val: number): number {
  return val / 100;
}

/** Convert degrees to OOXML angle units, 60000ths of a degree (90 → 5400000). */
export function degreesToOoxml(deg: number): number {
  return Math.round(deg * 60000);
}

/** Convert OOXML angle units to degrees (5400000 → 90). */
export function ooxmlToDegrees(val: number): number {
  return val / 60000;
}

/** Shape transform in EMU (as stored in data-ooxml-* attributes). */
export interface ShapeTransform {
  x: number;
  y: number;
  cx: number;
  cy: number;
  rot: number;
}

/** Find a shape's `<g>` element by its `data-ooxml-shape-idx`. */
export function findShapeElement(svg: SVGSVGElement, shapeIdx: number): SVGGElement | null {
  return svg.querySelector<SVGGElement>(`g[data-ooxml-shape-idx="${shapeIdx}"]`);
}

/** Read the OOXML transform (EMU) from a shape's `<g>` data attributes. */
export function getShapeTransform(g: SVGGElement): ShapeTransform {
  return {
    x: intAttr(g, 'data-ooxml-x'),
    y: intAttr(g, 'data-ooxml-y'),
    cx: intAttr(g, 'data-ooxml-cx'),
    cy: intAttr(g, 'data-ooxml-cy'),
    rot: intAttr(g, 'data-ooxml-rot'),
  };
}

/** All shape `<g>` elements in an SVG (direct children with `data-ooxml-shape-idx`). */
export function getAllShapes(svg: SVGSVGElement): SVGGElement[] {
  return Array.from(svg.querySelectorAll<SVGGElement>('g[data-ooxml-shape-idx]'));
}

/** The EMU-per-pixel scale from the SVG root's `data-ooxml-scale`. */
export function getSlideScale(svg: SVGSVGElement): number {
  return intAttr(svg, 'data-ooxml-scale') || EMU_PER_PX_96DPI;
}

function intAttr(el: Element, name: string): number {
  return parseInt(el.getAttribute(name) ?? '0', 10) || 0;
}
