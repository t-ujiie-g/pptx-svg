/**
 * Editing helpers — unit conversion and SVG DOM utilities for interactive editing.
 *
 * These helpers let frontend developers work with familiar units (px, pt, degrees)
 * instead of OOXML's EMU / hundredths-of-a-point / 60000ths-of-a-degree.
 */

// ── Unit constants ──────────────────────────────────────────────────────────

/** EMU per point (1 pt = 12700 EMU). */
export const EMU_PER_PT = 12700;

/** EMU per pixel at 96 DPI (914400 / 96 = 9525). */
export const EMU_PER_PX_96DPI = 9525;

// ── Unit conversion functions ───────────────────────────────────────────────

/**
 * Convert pixels to EMU.
 * @param px - Pixel value
 * @param dpi - DPI (default 96)
 */
export function pxToEmu(px: number, dpi = 96): number {
  return Math.round(px * 914400 / dpi);
}

/**
 * Convert EMU to pixels.
 * @param emu - EMU value
 * @param dpi - DPI (default 96)
 */
export function emuToPx(emu: number, dpi = 96): number {
  return Math.round(emu * dpi / 914400);
}

/**
 * Convert points to OOXML hundredths-of-a-point.
 * @param pt - Point value (e.g. 18)
 * @returns Hundredths (e.g. 1800)
 */
export function ptToHundredths(pt: number): number {
  return Math.round(pt * 100);
}

/**
 * Convert OOXML hundredths-of-a-point to points.
 * @param val - Hundredths value (e.g. 1800)
 * @returns Points (e.g. 18)
 */
export function hundredthsToPt(val: number): number {
  return val / 100;
}

/**
 * Convert degrees to OOXML angle units (60000ths of a degree).
 * @param deg - Degrees (e.g. 90)
 * @returns OOXML angle (e.g. 5400000)
 */
export function degreesToOoxml(deg: number): number {
  return Math.round(deg * 60000);
}

/**
 * Convert OOXML angle units to degrees.
 * @param val - OOXML angle (e.g. 5400000)
 * @returns Degrees (e.g. 90)
 */
export function ooxmlToDegrees(val: number): number {
  return val / 60000;
}

// ── SVG DOM helpers ─────────────────────────────────────────────────────────

/** Shape transform in EMU (as stored in data-ooxml-* attributes). */
export interface ShapeTransform {
  x: number;
  y: number;
  cx: number;
  cy: number;
  rot: number;
}

/**
 * Find a shape's `<g>` element by its shape index.
 * Looks for `data-ooxml-shape-idx="N"` in the SVG.
 */
export function findShapeElement(svg: SVGSVGElement, shapeIdx: number): SVGGElement | null {
  return svg.querySelector<SVGGElement>(`g[data-ooxml-shape-idx="${shapeIdx}"]`);
}

/**
 * Read the OOXML transform from a shape's `<g>` element data attributes.
 * Returns EMU values.
 */
export function getShapeTransform(g: SVGGElement): ShapeTransform {
  return {
    x: intAttr(g, 'data-ooxml-x'),
    y: intAttr(g, 'data-ooxml-y'),
    cx: intAttr(g, 'data-ooxml-cx'),
    cy: intAttr(g, 'data-ooxml-cy'),
    rot: intAttr(g, 'data-ooxml-rot'),
  };
}

/**
 * Get all shape `<g>` elements in an SVG (direct children with data-ooxml-shape-idx).
 */
export function getAllShapes(svg: SVGSVGElement): SVGGElement[] {
  return Array.from(svg.querySelectorAll<SVGGElement>('g[data-ooxml-shape-idx]'));
}

/**
 * Get the EMU-to-pixel scale factor from the SVG root's data attribute.
 */
export function getSlideScale(svg: SVGSVGElement): number {
  return intAttr(svg, 'data-ooxml-scale') || EMU_PER_PX_96DPI;
}

// ── Internal ────────────────────────────────────────────────────────────────

function intAttr(el: Element, name: string): number {
  return parseInt(el.getAttribute(name) ?? '0', 10) || 0;
}
