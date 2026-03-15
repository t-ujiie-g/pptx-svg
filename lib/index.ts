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
export type { MeasureTextFn, PptxRendererOptions } from './pptx-renderer.js';
export { DEFAULT_FONT_FALLBACKS } from './font-fallbacks.js';
export type { FontFallbackMap } from './font-fallbacks.js';
export { bytesToBase64, crc32 } from './utils.js';
export { extractZip, buildZip } from './zip.js';
export type { ZipContents } from './zip.js';
export { parseWasmStringConstants, instantiateWasmWithFallback } from './wasm-compat.js';
