/**
 * pptx-svg — PPTX ↔ SVG conversion library
 *
 * @example
 * ```ts
 * import { PptxRenderer } from 'pptx-svg';
 *
 * const renderer = new PptxRenderer();
 * await renderer.init('./main.wasm');
 * await renderer.loadPptx(pptxArrayBuffer);
 *
 * const svgString = renderer.renderSlideSvg(0);
 * ```
 */

export { PptxRenderer } from './pptx-renderer.js';
export type { MeasureTextFn, PptxRendererOptions, LogLevel } from './pptx-renderer.js';
export { DEFAULT_FONT_FALLBACKS } from './font-fallbacks.js';
export type { FontFallbackMap } from './font-fallbacks.js';
export { bytesToBase64, crc32 } from './utils.js';
export { extractZip, buildZip } from './zip.js';
export type { ZipContents } from './zip.js';
export { parseWasmStringConstants, instantiateWasmWithFallback } from './wasm-compat.js';
export { emfToSvg } from './emf-converter.js';
export {
  EMU_PER_PT, EMU_PER_PX_96DPI,
  pxToEmu, emuToPx, ptToHundredths, hundredthsToPt,
  degreesToOoxml, ooxmlToDegrees,
  findShapeElement, getShapeTransform, getAllShapes, getSlideScale,
} from './editing-helpers.js';
export type { ShapeTransform } from './editing-helpers.js';
