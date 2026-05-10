/**
 * PptxRenderer — main API class for rendering PPTX files.
 *
 * Handles Wasm lifecycle, PPTX loading, SVG rendering, and export.
 */

import { bytesToBase64 } from './utils.js';
import { emfToSvg } from './emf-converter.js';
import { wmfToSvg } from './wmf-converter.js';
import { instantiateWasmWithFallback } from './wasm-compat.js';
import { extractZip, buildZip } from './zip.js';
import { DEFAULT_FONT_FALLBACKS } from './font-fallbacks.js';
import type { FontFallbackMap } from './font-fallbacks.js';

/** Wasm exports provided by the MoonBit module. */
interface PptxWasmExports {
  initialize_pptx(): string;
  get_slide_count(): number;
  is_slide_hidden(idx: number): number;
  get_slide_xml_raw(idx: number): string;
  get_entry_list(): string;
  render_slide_svg(idx: number): string;
  update_slide_from_svg(idx: number, svg: string): string;
  get_slide_ooxml(idx: number): string;
  get_modified_entries(): string;
  render_shape_svg(slideIdx: number, shapeIdx: number): string;
  update_shape_transform(slideIdx: number, shapeIdx: number,
    x: number, y: number, cx: number, cy: number, rot: number): string;
  update_shape_text(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, text: string): string;
  update_shape_fill(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number): string;
  delete_shape(slideIdx: number, shapeIdx: number): string;
  add_shape(slideIdx: number, geomType: string,
    x: number, y: number, cx: number, cy: number,
    fillR: number, fillG: number, fillB: number): string;
  add_shape_text(slideIdx: number, shapeIdx: number,
    text: string, fontSize: number,
    colorR: number, colorG: number, colorB: number): string;
  duplicate_shape(slideIdx: number, shapeIdx: number,
    dxEmu: number, dyEmu: number): string;
  update_shape_gradient_fill(slideIdx: number, shapeIdx: number,
    angle: number, stopsData: string): string;
  update_shape_stroke(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number, widthEmu: number, dash: string): string;
  add_paragraph(slideIdx: number, shapeIdx: number,
    text: string, align: string): string;
  delete_paragraph(slideIdx: number, shapeIdx: number, paraIdx: number): string;
  add_run(slideIdx: number, shapeIdx: number, paraIdx: number, text: string): string;
  delete_run(slideIdx: number, shapeIdx: number, paraIdx: number, runIdx: number): string;
  update_text_run_style(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, bold: number, italic: number): string;
  update_text_run_font_size(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontSize: number): string;
  update_text_run_color(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, r: number, g: number, b: number): string;
  update_text_run_font(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontFace: string, eaFont: string, csFont: string): string;
  update_paragraph_align(slideIdx: number, shapeIdx: number,
    paraIdx: number, align: string): string;
  update_text_run_decoration(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, underline: string, strike: string, baseline: number): string;
  add_picture_shape(slideIdx: number, rid: string,
    x: number, y: number, cx: number, cy: number): string;
  replace_picture_rid(slideIdx: number, shapeIdx: number, newRid: string): string;
}

/** Options for text measurement callback. Font size is in CSS pixels (px). */
export interface MeasureTextFn {
  (text: string, fontFace: string, fontSizePx: number): number;
}

/** A comment on a slide. */
export interface SlideComment {
  authorId: number;
  date: string;
  index: number;
  text: string;
  x: number;
  y: number;
}

/** A comment author in the presentation. */
export interface CommentAuthor {
  id: number;
  name: string;
  initials: string;
}

/** Log level for controlling console output. */
export type LogLevel = 'silent' | 'error' | 'warn' | 'info' | 'debug';

/** Options for initializing PptxRenderer. */
export interface PptxRendererOptions {
  /** Custom text measurement function. If not provided, uses Canvas 2D (browser only). */
  measureText?: MeasureTextFn;
  /**
   * Custom font fallback mappings. Merged with built-in defaults (lib/font-fallbacks.ts).
   * User entries override built-in entries for the same font name.
   */
  fontFallbacks?: FontFallbackMap;
  /**
   * Log level for console output. Default: `'error'`.
   * - `'silent'`: No console output at all
   * - `'error'`:  Errors only (default)
   * - `'warn'`:   Errors + warnings
   * - `'info'`:   Errors + warnings + info messages
   * - `'debug'`:  All messages including debug details
   */
  logLevel?: LogLevel;
}

const LOG_LEVELS: Record<LogLevel, number> = {
  silent: 0, error: 1, warn: 2, info: 3, debug: 4,
};

/** Internal logger that respects log level. */
export interface Logger {
  debug(...args: unknown[]): void;
  info(...args: unknown[]): void;
  warn(...args: unknown[]): void;
  error(...args: unknown[]): void;
}

function createLogger(level: LogLevel): Logger {
  const threshold = LOG_LEVELS[level];
  return {
    debug: threshold >= 4 ? (...args: unknown[]) => console.log('[pptx]', ...args) : () => {},
    info:  threshold >= 3 ? (...args: unknown[]) => console.log('[pptx]', ...args) : () => {},
    warn:  threshold >= 2 ? (...args: unknown[]) => console.warn('[pptx]', ...args) : () => {},
    error: threshold >= 1 ? (...args: unknown[]) => console.error('[pptx]', ...args) : () => {},
  };
}

/** Default Wasm URL resolved relative to this module. */
const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;

// ── OOXML constants ──────────────────────────────────────────────────────────

const RELS_XMLNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL_TYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const REL_TYPE_SLIDE_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const CONTENT_TYPE_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const NS_DRAWINGML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PRESENTATIONML = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** Default slide size: 10" x 5.625" (standard widescreen 16:9) */
const DEFAULT_SLIDE_CX = 9144000;
const DEFAULT_SLIDE_CY = 5143500;
const REL_TYPE_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/**
 * Escape RegExp metacharacters so untrusted strings (rIds, targets, etc.
 * from PPTX) can be embedded in dynamic patterns without breaking them.
 * Without this, malicious input can throw `SyntaxError` or alter matching.
 */
function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Map MIME type to file extension. Returns null for unsupported types. */
function mimeToExt(mime: string): string | null {
  const map: Record<string, string> = {
    'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg',
    'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/webp': 'webp',
    'image/bmp': 'bmp', 'image/tiff': 'tiff',
  };
  return map[mime] ?? null;
}

/** First slide ID used in presentation.xml sldIdLst (OOXML convention). */
const FIRST_SLIDE_ID = 256;

/** Regex patterns for slide file paths. */
const RE_SLIDE_FILE = /^ppt\/slides\/slide\d+\.xml$/;
const RE_SLIDE_RELS_FILE = /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/;

export class PptxRenderer {
  private wasm: WebAssembly.Instance | null = null;

  /** Decompressed text ZIP entries (path → UTF-8 string) */
  private files = new Map<string, string>();

  /** Raw binary ZIP entries (path → bytes) */
  private rawFiles = new Map<string, Uint8Array>();

  /** Original PPTX bytes for export */
  private originalBuffer: ArrayBuffer | null = null;

  /** Files added after loadPptx (not in original ZIP) */
  private addedFiles = new Map<string, string>();

  /** Files removed after loadPptx */
  private removedFiles = new Set<string>();

  /** Binary files added/replaced after loadPptx (e.g. images) */
  private addedBinaryFiles = new Map<string, Uint8Array>();

  /** Canvas for text measurement (lazily created) */
  private canvas: HTMLCanvasElement | null = null;
  private ctx: CanvasRenderingContext2D | null = null;

  /** Custom text measurement function */
  private measureTextFn: MeasureTextFn | null = null;

  /** Font fallback lookup map (source → comma-separated fallbacks) */
  private fontFallbackCache = new Map<string, string>();

  /** Internal logger */
  private log: Logger;

  constructor(options?: PptxRendererOptions) {
    this.log = createLogger(options?.logLevel ?? 'error');
    if (options?.measureText) {
      this.measureTextFn = options.measureText;
    }
    // Build font fallback cache: merge defaults with user overrides
    const merged: FontFallbackMap = { ...DEFAULT_FONT_FALLBACKS, ...options?.fontFallbacks };
    for (const [font, fallbacks] of Object.entries(merged)) {
      this.fontFallbackCache.set(font, fallbacks.join(', '));
    }
  }

  /** Get typed Wasm exports. */
  private get exports(): PptxWasmExports {
    if (!this.wasm) throw new Error('Wasm not initialized — call init() first.');
    return this.wasm.exports as unknown as PptxWasmExports;
  }

  /**
   * Initialize the renderer by loading the Wasm module.
   *
   * When called without arguments, the bundled Wasm binary is loaded
   * automatically via `import.meta.url` resolution. This works with
   * Vite, webpack, Rollup, and CDN imports (unpkg, jsdelivr).
   *
   * In Node.js, pass an ArrayBuffer (e.g. from `fs.readFileSync(path).buffer`)
   * or a file:// URL / http(s):// URL string.
   *
   * @param wasmSource - Optional URL string, file path (Node.js), or ArrayBuffer of .wasm bytes.
   *                     If omitted, the bundled Wasm is used.
   */
  async init(wasmSource?: string | ArrayBuffer | Uint8Array): Promise<void> {
    let bytes: ArrayBuffer;
    if (wasmSource instanceof ArrayBuffer) {
      bytes = wasmSource;
    } else if (wasmSource instanceof Uint8Array) {
      // Accept Uint8Array (common in Node.js: fs.readFileSync returns Buffer which extends Uint8Array)
      bytes = wasmSource.buffer.slice(wasmSource.byteOffset, wasmSource.byteOffset + wasmSource.byteLength) as ArrayBuffer;
    } else {
      const url = wasmSource ?? DEFAULT_WASM_URL;
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status} fetching ${url}`);
      bytes = await response.arrayBuffer();
    }

    const result = await instantiateWasmWithFallback(bytes, this.buildImportObject(), this.log);
    this.wasm = result.instance;
  }

  /**
   * Load a PPTX file from an ArrayBuffer.
   * @returns Object with slideCount
   */
  async loadPptx(arrayBuffer: ArrayBuffer): Promise<{ slideCount: number }> {
    if (!this.wasm) {
      throw new Error('Wasm not initialized — wait for init() to complete before loading files.');
    }
    this.originalBuffer = arrayBuffer.slice(0); // keep a copy for export
    this.addedFiles.clear();
    this.addedBinaryFiles.clear();
    this.removedFiles.clear();
    this.log.debug('Parsing ZIP archive...');
    const { textFiles, binaryFiles } = await extractZip(arrayBuffer, this.log);
    this.files = textFiles;
    this.rawFiles = binaryFiles;
    this.log.debug(`Extracted ${textFiles.size} text entries, ${binaryFiles.size} binary entries`);

    const result = this.exports.initialize_pptx();
    this.log.debug('initialize_pptx result:', result);

    if (result.startsWith('ERROR:')) throw new Error(result.slice(6));

    const slideCount = this.exports.get_slide_count();
    return { slideCount };
  }

  /** Number of slides in the loaded presentation. */
  getSlideCount(): number {
    return this.exports.get_slide_count();
  }

  /** Check if a slide is hidden (0-indexed). */
  isSlideHidden(slideIdx: number): boolean {
    return this.exports.is_slide_hidden(slideIdx) === 1;
  }

  /** Get the raw XML of a slide (0-indexed). For debugging. */
  getSlideXmlRaw(slideIdx: number): string {
    return this.exports.get_slide_xml_raw(slideIdx);
  }

  /** Get all entry paths in the PPTX archive. For debugging. */
  getEntryList(): string[] {
    return this.exports.get_entry_list().split('\n').filter(Boolean);
  }

  /**
   * Render a slide as an SVG string (0-indexed).
   * @returns SVG markup, or a string starting with "ERROR:" on failure
   */
  renderSlideSvg(slideIdx: number): string {
    return this.exports.render_slide_svg(slideIdx);
  }

  /**
   * Update a slide's internal data from an edited SVG string.
   * Parses the SVG's data-ooxml-* attributes back into SlideData.
   * @returns "OK" on success, "ERROR:..." on failure
   */
  updateSlideFromSvg(slideIdx: number, svgString: string): string {
    return this.exports.update_slide_from_svg(slideIdx, svgString);
  }

  /**
   * Get the OOXML slide XML for a slide (0-indexed).
   * Returns modified XML if the slide was updated, otherwise original.
   */
  getSlideOoxml(slideIdx: number): string {
    return this.exports.get_slide_ooxml(slideIdx);
  }

  /**
   * Export the (possibly modified) presentation as a PPTX ArrayBuffer.
   * Replaces modified slide XML entries in the original ZIP and rebuilds it.
   */
  async exportPptx(): Promise<ArrayBuffer> {
    if (!this.originalBuffer) {
      throw new Error('No PPTX loaded — call loadPptx() first.');
    }

    // Get modified entries from Wasm: "path\tcontent\n..."
    const modifiedStr = this.exports.get_modified_entries();
    const modifications = new Map<string, string>();
    if (modifiedStr) {
      const lines = modifiedStr.split('\n');
      for (const line of lines) {
        if (!line) continue;
        const tabIdx = line.indexOf('\t');
        if (tabIdx < 0) continue;
        const path = line.substring(0, tabIdx);
        const content = line.substring(tabIdx + 1);
        modifications.set(path, content);
      }
    }

    // Merge in files added/modified by slide operations
    for (const [path, content] of this.addedFiles) {
      if (!modifications.has(path)) {
        modifications.set(path, content);
      }
    }

    this.log.debug(`Exporting PPTX with ${modifications.size} modified entries, ${this.addedBinaryFiles.size} binary entries, ${this.removedFiles.size} removals`);
    return buildZip(
      this.originalBuffer, modifications,
      this.removedFiles.size > 0 ? this.removedFiles : undefined,
      this.addedBinaryFiles.size > 0 ? this.addedBinaryFiles : undefined,
    );
  }

  // ── Notes & Comments API ──────────────────────────────────────────────────

  /**
   * Get speaker notes text for a slide (0-indexed).
   * @returns Array of paragraph strings, or empty array if no notes exist.
   */
  getSlideNotes(slideIdx: number): string[] {
    const notesPath = this.resolveRelTarget(slideIdx, 'notesSlide');
    if (!notesPath) return [];
    const xml = this.files.get(notesPath);
    if (!xml) return [];
    return this.extractNotesText(xml);
  }

  /**
   * Get comments for a slide (0-indexed).
   * @returns Array of comment objects, or empty array if no comments exist.
   */
  getSlideComments(slideIdx: number): SlideComment[] {
    const commentsPath = this.resolveRelTarget(slideIdx, 'comments');
    if (!commentsPath) return [];
    const xml = this.files.get(commentsPath);
    if (!xml) return [];
    return this.parseComments(xml);
  }

  /**
   * Get comment authors defined in the presentation.
   * @returns Array of author objects, or empty array if none exist.
   */
  getCommentAuthors(): CommentAuthor[] {
    const xml = this.files.get('ppt/commentAuthors.xml');
    if (!xml) return [];
    return this.parseCommentAuthors(xml);
  }

  // ── Shape-level editing API ────────────────────────────────────────────────

  /**
   * Render a single shape as SVG (0-indexed slide and shape).
   * Returns SVG fragment (`<defs>...` + `<g>...</g>`), or "ERROR:..." on failure.
   */
  renderShapeSvg(slideIdx: number, shapeIdx: number): string {
    return this.exports.render_shape_svg(slideIdx, shapeIdx);
  }

  /**
   * Update a shape's transform (position, size, rotation) and return re-rendered SVG.
   * All values are in EMU. Marks the slide as modified.
   */
  updateShapeTransform(slideIdx: number, shapeIdx: number,
    x: number, y: number, cx: number, cy: number, rot: number): string {
    return this.exports.update_shape_transform(slideIdx, shapeIdx, x, y, cx, cy, rot);
  }

  /**
   * Update a text run's content and return the shape's re-rendered SVG.
   * Marks the slide as modified.
   */
  updateShapeText(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, text: string): string {
    return this.exports.update_shape_text(slideIdx, shapeIdx, paraIdx, runIdx, text);
  }

  /**
   * Update a shape's solid fill color (RGB 0-255) and return re-rendered SVG.
   * Marks the slide as modified.
   */
  updateShapeFill(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number): string {
    return this.exports.update_shape_fill(slideIdx, shapeIdx, r, g, b);
  }

  /**
   * Delete a shape from a slide.
   * @returns "OK" on success, "ERROR:..." on failure.
   */
  deleteShape(slideIdx: number, shapeIdx: number): string {
    return this.exports.delete_shape(slideIdx, shapeIdx);
  }

  /**
   * Add a basic AutoShape to a slide.
   * @param geomType - Preset geometry: "rect", "ellipse", "roundRect", "line", etc.
   * @param x, y, cx, cy - Position and size in EMU.
   * @param fillR, fillG, fillB - Fill color (0-255). Pass -1 for no fill.
   * @returns "OK:<shapeIndex>" on success, "ERROR:..." on failure.
   */
  addShape(slideIdx: number, geomType: string,
    x: number, y: number, cx: number, cy: number,
    fillR = -1, fillG = -1, fillB = -1): string {
    return this.exports.add_shape(slideIdx, geomType, x, y, cx, cy, fillR, fillG, fillB);
  }

  /**
   * Add a text paragraph to a shape. Creates a single run with the given text.
   * @param fontSize - Font size in hundredths of a point (e.g. 1800 = 18pt). 0 = inherit.
   * @param colorR, colorG, colorB - Text color (0-255). Pass -1 for default/inherit.
   * @returns "OK:<paraIndex>" on success, "ERROR:..." on failure.
   */
  addShapeText(slideIdx: number, shapeIdx: number, text: string,
    fontSize = 0, colorR = -1, colorG = -1, colorB = -1): string {
    return this.exports.add_shape_text(slideIdx, shapeIdx, text, fontSize, colorR, colorG, colorB);
  }

  /**
   * Duplicate a shape, offset by (dxEmu, dyEmu) from the original.
   * @returns "OK:<shapeIndex>" on success, "ERROR:..." on failure.
   */
  duplicateShape(slideIdx: number, shapeIdx: number,
    dxEmu = 457200, dyEmu = 457200): string {
    return this.exports.duplicate_shape(slideIdx, shapeIdx, dxEmu, dyEmu);
  }

  /**
   * Update a shape's fill to a linear gradient. Returns re-rendered SVG.
   * @param angle - Gradient angle in 60000ths of a degree (e.g. 5400000 = 90deg).
   * @param stops - Array of { pos, r, g, b } where pos is 0-100000.
   */
  updateShapeGradientFill(slideIdx: number, shapeIdx: number,
    angle: number, stops: Array<{ pos: number; r: number; g: number; b: number }>): string {
    const stopsData = stops.map(s => `${s.pos},${s.r},${s.g},${s.b}`).join(';');
    return this.exports.update_shape_gradient_fill(slideIdx, shapeIdx, angle, stopsData);
  }

  /**
   * Update a shape's stroke (outline). Returns re-rendered SVG.
   * @param r, g, b - Stroke color (0-255). Pass -1 to remove stroke.
   * @param widthEmu - Stroke width in EMU (default 12700 = 1pt).
   * @param dash - Dash preset: "dash", "dot", "dashDot", "lgDash", etc. "" = solid.
   */
  updateShapeStroke(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number, widthEmu = 12700, dash = ''): string {
    return this.exports.update_shape_stroke(slideIdx, shapeIdx, r, g, b, widthEmu, dash);
  }

  // ── Text editing API (E2.5) ─────────────────────────────────────────────────

  /**
   * Add a new paragraph to a shape with a single text run.
   * @param align - "l" (left), "ctr" (center), "r" (right), "just" (justify), "" (inherit).
   * @returns "OK:<paraIndex>" on success, "ERROR:..." on failure.
   */
  addParagraph(slideIdx: number, shapeIdx: number, text: string, align = ''): string {
    return this.exports.add_paragraph(slideIdx, shapeIdx, text, align);
  }

  /**
   * Delete a paragraph from a shape.
   * @returns "OK" on success, "ERROR:..." on failure.
   */
  deleteParagraph(slideIdx: number, shapeIdx: number, paraIdx: number): string {
    return this.exports.delete_paragraph(slideIdx, shapeIdx, paraIdx);
  }

  /**
   * Add a new text run to a paragraph.
   * @returns "OK:<runIndex>" on success, "ERROR:..." on failure.
   */
  addRun(slideIdx: number, shapeIdx: number, paraIdx: number, text: string): string {
    return this.exports.add_run(slideIdx, shapeIdx, paraIdx, text);
  }

  /**
   * Delete a text run from a paragraph.
   * @returns "OK" on success, "ERROR:..." on failure.
   */
  deleteRun(slideIdx: number, shapeIdx: number, paraIdx: number, runIdx: number): string {
    return this.exports.delete_run(slideIdx, shapeIdx, paraIdx, runIdx);
  }

  /**
   * Update a text run's bold/italic style. Returns re-rendered shape SVG.
   * @param bold - 1 = set bold, 0 = unset bold, -1 = no change.
   * @param italic - 1 = set italic, 0 = unset italic, -1 = no change.
   */
  updateTextRunStyle(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, bold = -1, italic = -1): string {
    return this.exports.update_text_run_style(slideIdx, shapeIdx, paraIdx, runIdx, bold, italic);
  }

  /**
   * Update a text run's font size. Returns re-rendered shape SVG.
   * @param fontSize - In hundredths of a point (e.g. 1800 = 18pt). 0 = inherit.
   */
  updateTextRunFontSize(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontSize: number): string {
    return this.exports.update_text_run_font_size(slideIdx, shapeIdx, paraIdx, runIdx, fontSize);
  }

  /**
   * Update a text run's color (RGB 0-255). Returns re-rendered shape SVG.
   * Pass r = -1 to clear (inherit from theme/master).
   */
  updateTextRunColor(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, r: number, g: number, b: number): string {
    return this.exports.update_text_run_color(slideIdx, shapeIdx, paraIdx, runIdx, r, g, b);
  }

  /**
   * Update a text run's font family. Returns re-rendered shape SVG.
   * @param fontFace - Latin font name. Empty string = no change.
   * @param eaFont - East Asian font name. Empty string = no change.
   * @param csFont - Complex Script font name. Empty string = no change.
   */
  updateTextRunFont(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontFace = '', eaFont = '', csFont = ''): string {
    return this.exports.update_text_run_font(slideIdx, shapeIdx, paraIdx, runIdx, fontFace, eaFont, csFont);
  }

  /**
   * Update a paragraph's alignment. Returns re-rendered shape SVG.
   * @param align - "l" (left), "ctr" (center), "r" (right), "just" (justify), "" (inherit).
   */
  updateParagraphAlign(slideIdx: number, shapeIdx: number,
    paraIdx: number, align: string): string {
    return this.exports.update_paragraph_align(slideIdx, shapeIdx, paraIdx, align);
  }

  /**
   * Update a text run's decoration (underline, strikethrough, baseline shift).
   * Returns re-rendered shape SVG.
   * @param underline - "sng" (single), "dbl" (double), "" (no change), "none" (remove).
   * @param strike - "sngStrike", "dblStrike", "" (no change), "none" (remove).
   * @param baseline - 30000 = superscript, -25000 = subscript, 0 = normal, -1 = no change.
   */
  updateTextRunDecoration(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, underline = '', strike = '', baseline = -1): string {
    return this.exports.update_text_run_decoration(slideIdx, shapeIdx, paraIdx, runIdx, underline, strike, baseline);
  }

  // ── Slide management API ────────────────────────────────────────────────────

  /**
   * Add a blank slide at the specified position (0-indexed).
   * If `afterIdx` is omitted, the slide is appended at the end.
   * If `sourceSlideIdx` is provided, the new slide copies that slide's layout.
   *
   * @param afterIdx - Insert after this slide index (0-indexed). Use -1 to insert at the beginning.
   * @param sourceSlideIdx - Copy layout from this slide (0-indexed). Defaults to last slide.
   * @returns Object with the new slide count and the index of the inserted slide.
   */
  async addSlide(afterIdx?: number, sourceSlideIdx?: number): Promise<{ slideCount: number; insertedIdx: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const oldCount = this.exports.get_slide_count();
    if (oldCount === 0) throw new Error('No presentation loaded.');

    // Determine insert position
    const insertIdx = afterIdx === undefined ? oldCount : afterIdx + 1;
    if (insertIdx < 0 || insertIdx > oldCount) {
      throw new Error(`Invalid insert position: ${insertIdx}`);
    }

    // Determine source layout info
    const srcIdx = sourceSlideIdx ?? Math.max(0, oldCount - 1);
    const srcRelsPath = `ppt/slides/_rels/slide${srcIdx + 1}.xml.rels`;
    const srcRelsXml = this.files.get(srcRelsPath) ?? '';
    const layoutTarget = this.extractRelTarget(srcRelsXml, '/slideLayout');

    // Collect all current slide contents (path → xml) for renumbering
    const slideContents: string[] = [];       // slide XML at index i
    const slideRels: string[] = [];           // slide .rels at index i
    for (let i = 0; i < oldCount; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    // Create blank slide XML
    const slideSize = this.extractSlideSize();
    const blankSlideXml = this.createBlankSlideXml(slideSize.cx, slideSize.cy);

    // Create .rels for new slide pointing to its layout
    const blankRels = layoutTarget
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_XMLNS}"><Relationship Id="rId1" Type="${REL_TYPE_SLIDE_LAYOUT}" Target="${layoutTarget}"/></Relationships>`
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_XMLNS}"></Relationships>`;

    // Insert into arrays
    slideContents.splice(insertIdx, 0, blankSlideXml);
    slideRels.splice(insertIdx, 0, blankRels);

    const newCount = oldCount + 1;

    // Write back all slides with correct numbering
    this.rewriteSlideFiles(slideContents, slideRels, newCount);

    // Update presentation.xml
    this.updatePresentationXmlForAdd(insertIdx);

    // Update [Content_Types].xml if needed
    this.ensureContentTypeForSlide(newCount);

    // Re-initialize Wasm
    this.reinitializeWasm();

    return { slideCount: newCount, insertedIdx: insertIdx };
  }

  /**
   * Delete a slide at the specified index (0-indexed).
   * At least one slide must remain.
   *
   * @param slideIdx - The 0-indexed slide to delete.
   * @returns Object with the new slide count.
   */
  async deleteSlide(slideIdx: number): Promise<{ slideCount: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const oldCount = this.exports.get_slide_count();
    if (oldCount <= 1) throw new Error('Cannot delete the last remaining slide.');
    if (slideIdx < 0 || slideIdx >= oldCount) {
      throw new Error(`Slide index out of range: ${slideIdx}`);
    }

    // Collect current slides
    const slideContents: string[] = [];
    const slideRels: string[] = [];
    for (let i = 0; i < oldCount; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    // Remove the slide at slideIdx
    slideContents.splice(slideIdx, 1);
    slideRels.splice(slideIdx, 1);

    const newCount = oldCount - 1;

    // Mark old highest-numbered files for removal (in case they persist)
    const oldPath = `ppt/slides/slide${oldCount}.xml`;
    const oldRelsPath = `ppt/slides/_rels/slide${oldCount}.xml.rels`;
    this.removedFiles.add(oldPath);
    this.removedFiles.add(oldRelsPath);

    // Write back all slides with correct numbering
    this.rewriteSlideFiles(slideContents, slideRels, newCount);

    // Update presentation.xml
    this.updatePresentationXmlForDelete(slideIdx);

    // Re-initialize Wasm
    this.reinitializeWasm();

    return { slideCount: newCount };
  }

  /**
   * Reorder slides according to the given index mapping.
   *
   * @param newOrder - Array where newOrder[i] is the old index of the slide that should appear at position i.
   *                   Must be a permutation of [0, 1, ..., slideCount-1].
   * @returns Object with the slide count (unchanged).
   */
  async reorderSlides(newOrder: number[]): Promise<{ slideCount: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const count = this.exports.get_slide_count();
    if (newOrder.length !== count) {
      throw new Error(`newOrder length (${newOrder.length}) must equal slide count (${count}).`);
    }

    // Validate permutation
    const seen = new Set<number>();
    for (const idx of newOrder) {
      if (idx < 0 || idx >= count || seen.has(idx)) {
        throw new Error(`Invalid permutation: ${JSON.stringify(newOrder)}`);
      }
      seen.add(idx);
    }

    // Collect current slides
    const slideContents: string[] = [];
    const slideRels: string[] = [];
    for (let i = 0; i < count; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    // Reorder
    const reordered: string[] = newOrder.map(i => slideContents[i]);
    const reorderedRels: string[] = newOrder.map(i => slideRels[i]);

    // Write back
    this.rewriteSlideFiles(reordered, reorderedRels, count);

    // Reorder sldId entries in presentation.xml
    this.updatePresentationXmlForReorder(newOrder);

    // Re-initialize Wasm
    this.reinitializeWasm();

    return { slideCount: count };
  }

  // ── Slide management helpers ─────────────────────────────────────────────

  /** Rewrite slide files in the files map with correct 1..N numbering. */
  private rewriteSlideFiles(contents: string[], rels: string[], count: number): void {
    // Remove all old slide files first
    for (const key of [...this.files.keys()]) {
      if (RE_SLIDE_FILE.test(key) || RE_SLIDE_RELS_FILE.test(key)) {
        this.files.delete(key);
      }
    }

    // Write new files
    for (let i = 0; i < count; i++) {
      const num = i + 1;
      this.persistFile(`ppt/slides/slide${num}.xml`, contents[i]);
      this.persistFile(`ppt/slides/_rels/slide${num}.xml.rels`, rels[i]);
    }
  }

  /** Extract a relationship target from .rels XML by type suffix. */
  private extractRelTarget(relsXml: string, typeSuffix: string): string | null {
    const ts = escapeRegex(typeSuffix);
    const re = new RegExp(`<Relationship[^>]+Type="[^"]*${ts}"[^>]+Target="([^"]+)"`);
    const m = relsXml.match(re);
    if (m) return m[1];
    // Try reversed attr order
    const re2 = new RegExp(`<Relationship[^>]+Target="([^"]+)"[^>]+Type="[^"]*${ts}"`);
    const m2 = relsXml.match(re2);
    return m2 ? m2[1] : null;
  }

  /** Extract slide size from presentation.xml. */
  private extractSlideSize(): { cx: number; cy: number } {
    const prsXml = this.files.get('ppt/presentation.xml') ?? '';
    const m = prsXml.match(/<p:sldSz[^>]+cx="(\d+)"[^>]+cy="(\d+)"/);
    if (m) return { cx: parseInt(m[1]), cy: parseInt(m[2]) };
    return { cx: DEFAULT_SLIDE_CX, cy: DEFAULT_SLIDE_CY };
  }

  /** Create minimal blank slide XML. */
  private createBlankSlideXml(cx: number, cy: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld xmlns:a="${NS_DRAWINGML}" xmlns:r="${NS_RELS}" xmlns:p="${NS_PRESENTATIONML}">` +
      `<p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:grpSpPr>` +
      `</p:spTree></p:cSld></p:sld>`;
  }

  /** Find the next available sldId in presentation.xml. */
  private findNextSlideId(prsXml: string): number {
    const ids: number[] = [];
    const re = /id="(\d+)"/g;
    // Only look inside sldIdLst
    const sldListMatch = prsXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (sldListMatch) {
      let m: RegExpExecArray | null;
      while ((m = re.exec(sldListMatch[1])) !== null) {
        ids.push(parseInt(m[1]));
      }
    }
    return ids.length > 0 ? Math.max(...ids) + 1 : FIRST_SLIDE_ID;
  }

  /** Find the next available rId in a .rels XML string. */
  private nextRid(relsXml: string): string {
    let maxId = 0;
    const re = /Id="rId(\d+)"/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(relsXml)) !== null) {
      const id = parseInt(m[1]);
      if (id > maxId) maxId = id;
    }
    return `rId${maxId + 1}`;
  }

  /** Update presentation.xml and .rels for slide addition. */
  private updatePresentationXmlForAdd(insertIdx: number): void {
    let prsXml = this.files.get('ppt/presentation.xml') ?? '';
    let prsRels = this.files.get('ppt/_rels/presentation.xml.rels') ?? '';

    const newSlideNum = insertIdx + 1;
    const newRId = this.nextRid(prsRels);
    const newSldId = this.findNextSlideId(prsXml);

    // Add relationship for new slide
    const newRelEntry = `<Relationship Id="${newRId}" Type="${REL_TYPE_SLIDE}" Target="slides/slide${newSlideNum}.xml"/>`;
    prsRels = prsRels.replace('</Relationships>', newRelEntry + '</Relationships>');

    // Update existing slide relationship targets (renumber slides after insertIdx)
    // First, collect existing slide rId → target mappings
    const slideRIdMap = this.parseSlideRelationships(prsRels);

    // Renumber: slides after insertIdx shift by +1
    for (const [rId, target] of slideRIdMap) {
      const m = target.match(/^slides\/slide(\d+)\.xml$/);
      if (m) {
        const num = parseInt(m[1]);
        if (num >= newSlideNum && rId !== newRId) {
          const newTarget = `slides/slide${num + 1}.xml`;
          const eRId = escapeRegex(rId);
          const eTarget = escapeRegex(target);
          prsRels = prsRels.replace(
            new RegExp(`(<Relationship[^>]+Id="${eRId}"[^>]+Target=")${eTarget}"`),
            `$1${newTarget}"`
          );
          // Try reversed order too
          prsRels = prsRels.replace(
            new RegExp(`(<Relationship[^>]+Target=")${eTarget}"([^>]+Id="${eRId}")`),
            `$1${newTarget}"$2`
          );
        }
      }
    }

    // Add sldId entry to presentation.xml
    const newSldIdEntry = `<p:sldId id="${newSldId}" r:id="${newRId}"/>`;

    if (insertIdx === 0) {
      // Insert at the beginning of sldIdLst
      prsXml = prsXml.replace('<p:sldIdLst>', `<p:sldIdLst>${newSldIdEntry}`);
    } else {
      // Insert after the insertIdx-1 th sldId entry
      const sldIdEntries = this.parseSldIdEntries(prsXml);
      if (insertIdx <= sldIdEntries.length) {
        const afterEntry = sldIdEntries[insertIdx - 1];
        prsXml = prsXml.replace(afterEntry, afterEntry + newSldIdEntry);
      } else {
        // Append at end
        prsXml = prsXml.replace('</p:sldIdLst>', newSldIdEntry + '</p:sldIdLst>');
      }
    }

    this.persistFile('ppt/presentation.xml', prsXml);
    this.persistFile('ppt/_rels/presentation.xml.rels', prsRels);
  }

  /** Update presentation.xml and .rels for slide deletion. */
  private updatePresentationXmlForDelete(deleteIdx: number): void {
    let prsXml = this.files.get('ppt/presentation.xml') ?? '';
    let prsRels = this.files.get('ppt/_rels/presentation.xml.rels') ?? '';

    const deleteNum = deleteIdx + 1;

    // Find the rId for the deleted slide
    const slideRIdMap = this.parseSlideRelationships(prsRels);
    let deleteRId = '';
    for (const [rId, target] of slideRIdMap) {
      if (target === `slides/slide${deleteNum}.xml`) {
        deleteRId = rId;
        break;
      }
    }

    // Remove sldId entry from presentation.xml
    if (deleteRId) {
      const sldIdEntries = this.parseSldIdEntries(prsXml);
      for (const entry of sldIdEntries) {
        if (entry.includes(`r:id="${deleteRId}"`)) {
          prsXml = prsXml.replace(entry, '');
          break;
        }
      }
    }

    // Remove relationship entry
    if (deleteRId) {
      const relRe = new RegExp(`<Relationship[^>]+Id="${escapeRegex(deleteRId)}"[^>]*/?>`, 'g');
      prsRels = prsRels.replace(relRe, '');
    }

    // Renumber remaining slide targets (shift down slides after deleteNum)
    for (const [rId, target] of slideRIdMap) {
      if (rId === deleteRId) continue;
      const m = target.match(/^slides\/slide(\d+)\.xml$/);
      if (m) {
        const num = parseInt(m[1]);
        if (num > deleteNum) {
          const newTarget = `slides/slide${num - 1}.xml`;
          const eRId = escapeRegex(rId);
          const eTarget = escapeRegex(target);
          prsRels = prsRels.replace(
            new RegExp(`(Id="${eRId}"[^>]*Target=")${eTarget}"`),
            `$1${newTarget}"`
          );
          prsRels = prsRels.replace(
            new RegExp(`(Target=")${eTarget}"([^>]*Id="${eRId}")`),
            `$1${newTarget}"$2`
          );
        }
      }
    }

    this.persistFile('ppt/presentation.xml', prsXml);
    this.persistFile('ppt/_rels/presentation.xml.rels', prsRels);
  }

  /** Update presentation.xml sldId order for reorder. */
  private updatePresentationXmlForReorder(newOrder: number[]): void {
    let prsXml = this.files.get('ppt/presentation.xml') ?? '';
    const prsRels = this.files.get('ppt/_rels/presentation.xml.rels') ?? '';

    const sldIdEntries = this.parseSldIdEntries(prsXml);
    if (sldIdEntries.length !== newOrder.length) return;

    // Build reordered sldId list
    const reordered = newOrder.map(i => sldIdEntries[i]);

    // Also need to update r:id → Target mappings for renumbered files
    // Each sldId's r:id now points to a different slide file number
    const slideRIdMap = this.parseSlideRelationships(prsRels);

    // Extract rId from each sldId entry
    const rIds = sldIdEntries.map(entry => {
      const m = entry.match(/r:id="(rId\d+)"/);
      return m ? m[1] : '';
    });

    // Reordered rIds: reordered[i] has rIds[newOrder[i]]
    // We need rIds[newOrder[i]] to point to slides/slide{i+1}.xml
    let updatedRels = prsRels;
    for (let i = 0; i < newOrder.length; i++) {
      const rId = rIds[newOrder[i]];
      if (!rId) continue;
      const oldTarget = slideRIdMap.get(rId);
      const newTarget = `slides/slide${i + 1}.xml`;
      if (oldTarget && oldTarget !== newTarget) {
        // Use a unique placeholder to avoid conflicts during replacement
        const eRId = escapeRegex(rId);
        const eOldTarget = escapeRegex(oldTarget);
        updatedRels = updatedRels.replace(
          new RegExp(`(Id="${eRId}"[^>]*Target=")${eOldTarget}"`),
          `$1${newTarget}"`
        );
        updatedRels = updatedRels.replace(
          new RegExp(`(Target=")${eOldTarget}"([^>]*Id="${eRId}")`),
          `$1${newTarget}"$2`
        );
      }
    }

    // Replace sldIdLst content
    const sldIdListMatch = prsXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (sldIdListMatch) {
      prsXml = prsXml.replace(sldIdListMatch[1], reordered.join(''));
    }

    this.persistFile('ppt/presentation.xml', prsXml);
    this.persistFile('ppt/_rels/presentation.xml.rels', updatedRels);
  }

  /** Parse all <p:sldId .../> entries from presentation.xml as raw strings. */
  private parseSldIdEntries(prsXml: string): string[] {
    const entries: string[] = [];
    const re = /<p:sldId\s[^>]*?\/?>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(prsXml)) !== null) {
      entries.push(m[0]);
    }
    return entries;
  }

  /** Parse slide relationships: rId → target path. */
  private parseSlideRelationships(relsXml: string): Map<string, string> {
    const map = new Map<string, string>();
    const re = /<Relationship\s([^>]*)\/?>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(relsXml)) !== null) {
      const attrs = m[1];
      if (!attrs.includes(REL_TYPE_SLIDE + '"')) continue;
      const idMatch = attrs.match(/Id="(rId\d+)"/);
      const targetMatch = attrs.match(/Target="([^"]+)"/);
      if (idMatch && targetMatch) {
        map.set(idMatch[1], targetMatch[1]);
      }
    }
    return map;
  }

  /** Ensure [Content_Types].xml has an Override for the given slide number. */
  private ensureContentTypeForSlide(slideCount: number): void {
    let ct = this.files.get('[Content_Types].xml');
    if (!ct) return;

    for (let i = 1; i <= slideCount; i++) {
      const partName = `/ppt/slides/slide${i}.xml`;
      if (!ct.includes(partName)) {
        const override = `<Override PartName="${partName}" ContentType="${CONTENT_TYPE_SLIDE}"/>`;
        ct = ct.replace('</Types>', override + '</Types>');
      }
    }

    this.persistFile('[Content_Types].xml', ct);
  }

  // ── Image management API ───────────────────────────────────────────────────

  /**
   * Add an image to a slide as a Picture shape.
   * @param slideIdx - Slide index (0-based).
   * @param imageData - Raw image bytes (PNG, JPEG, GIF, etc.).
   * @param mimeType - MIME type (e.g. "image/png", "image/jpeg").
   * @param x, y, cx, cy - Position and size in EMU.
   * @returns "OK:<shapeIndex>" on success, "ERROR:..." on failure.
   */
  addImage(slideIdx: number, imageData: Uint8Array, mimeType: string,
    x: number, y: number, cx: number, cy: number): string {
    if (!this.wasm) return 'ERROR:not initialized';

    // Determine file extension from MIME type
    const ext = mimeToExt(mimeType);
    if (!ext) return 'ERROR:unsupported image type';

    // Generate unique media filename
    const mediaPath = this.nextMediaPath(ext);

    // Store binary in rawFiles (for Wasm FFI) and addedBinaryFiles (for export)
    this.rawFiles.set(mediaPath, imageData);
    this.addedBinaryFiles.set(mediaPath, imageData);

    // Add relationship to slide's .rels
    const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
    let relsXml = this.files.get(relsPath) ?? '';
    const rid = this.nextRid(relsXml);
    const relEntry = `<Relationship Id="${rid}" Type="${REL_TYPE_IMAGE}" Target="../media/${mediaPath.split('/').pop()}"/>`;
    if (relsXml.includes('</Relationships>')) {
      relsXml = relsXml.replace('</Relationships>', relEntry + '</Relationships>');
    } else {
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntry}</Relationships>`;
    }
    this.persistFile(relsPath, relsXml);

    // Ensure content type for this extension
    this.ensureContentTypeForExtension(ext, mimeType);

    // Add Picture shape via Wasm
    this.renderSlideSvg(slideIdx); // ensure slide is parsed
    return this.exports.add_picture_shape(slideIdx, rid, x, y, cx, cy);
  }

  /**
   * Replace the image data of an existing Picture shape.
   * @param slideIdx - Slide index (0-based).
   * @param shapeIdx - Shape index (composite index for groups).
   * @param imageData - New image bytes.
   * @param mimeType - MIME type of the new image.
   * @returns "OK" on success, "ERROR:..." on failure.
   */
  replaceImage(slideIdx: number, shapeIdx: number,
    imageData: Uint8Array, mimeType: string): string {
    if (!this.wasm) return 'ERROR:not initialized';

    const ext = mimeToExt(mimeType);
    if (!ext) return 'ERROR:unsupported image type';

    // Find the shape's current rid from SVG data attributes
    this.renderSlideSvg(slideIdx); // ensure slide is parsed
    const svg = this.exports.render_shape_svg(slideIdx, shapeIdx);
    if (svg.startsWith('ERROR:')) return svg;

    const ridMatch = svg.match(/data-ooxml-blip-rid="([^"]+)"/);
    if (!ridMatch) return 'ERROR:shape has no image reference';
    const currentRid = ridMatch[1];

    // Resolve current image path from rels
    const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
    const relsXml = this.files.get(relsPath) ?? '';
    const oldTarget = this.resolveRidTarget(relsXml, currentRid);

    if (oldTarget) {
      // Replace existing file at the same path
      const oldPath = this.resolveRelPath('ppt/slides/', oldTarget);

      if (mimeToExt(mimeType) === oldPath.split('.').pop()) {
        // Same extension — replace in place
        this.rawFiles.set(oldPath, imageData);
        this.addedBinaryFiles.set(oldPath, imageData);
        return 'OK';
      }
    }

    // Different extension or couldn't resolve — create new file and update rid
    const mediaPath = this.nextMediaPath(ext);
    this.rawFiles.set(mediaPath, imageData);
    this.addedBinaryFiles.set(mediaPath, imageData);

    const newRid = this.nextRid(relsXml);
    const relEntry = `<Relationship Id="${newRid}" Type="${REL_TYPE_IMAGE}" Target="../media/${mediaPath.split('/').pop()}"/>`;
    const updatedRels = relsXml.replace('</Relationships>', relEntry + '</Relationships>');
    this.persistFile(relsPath, updatedRels);
    this.ensureContentTypeForExtension(ext, mimeType);

    return this.exports.replace_picture_rid(slideIdx, shapeIdx, newRid);
  }

  /**
   * Delete a Picture shape and mark its media file for removal.
   * @param slideIdx - Slide index (0-based).
   * @param shapeIdx - Shape index.
   * @returns "OK" on success, "ERROR:..." on failure.
   */
  deleteImage(slideIdx: number, shapeIdx: number): string {
    if (!this.wasm) return 'ERROR:not initialized';

    // Find the shape's rid before deleting
    this.renderSlideSvg(slideIdx);
    const svg = this.exports.render_shape_svg(slideIdx, shapeIdx);
    let mediaPath: string | null = null;
    if (!svg.startsWith('ERROR:')) {
      const ridMatch = svg.match(/data-ooxml-blip-rid="([^"]+)"/);
      if (ridMatch) {
        const rid = ridMatch[1];
        const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
        const relsXml = this.files.get(relsPath) ?? '';
        const target = this.resolveRidTarget(relsXml, rid);
        if (target) {
          mediaPath = this.resolveRelPath('ppt/slides/', target);
        }
      }
    }

    // Delete the shape
    const result = this.exports.delete_shape(slideIdx, shapeIdx);
    if (result.startsWith('ERROR:')) return result;

    // Mark media file for removal (if not referenced by other slides)
    if (mediaPath && !this.isMediaReferencedElsewhere(mediaPath, slideIdx)) {
      this.removedFiles.add(mediaPath);
      this.rawFiles.delete(mediaPath);
    }

    return 'OK';
  }

  /** Resolve a relationship ID to its Target attribute in a .rels XML string. */
  private resolveRidTarget(relsXml: string, rid: string): string | null {
    const e = escapeRegex(rid);
    const m = relsXml.match(new RegExp(`<Relationship[^>]+Id="${e}"[^>]+Target="([^"]+)"`))
           ?? relsXml.match(new RegExp(`<Relationship[^>]+Target="([^"]+)"[^>]+Id="${e}"`));
    return m ? m[1] : null;
  }

  /** Generate the next available media file path. */
  private nextMediaPath(ext: string): string {
    let n = 1;
    const allPaths = new Set([...this.rawFiles.keys(), ...this.addedBinaryFiles.keys()]);
    while (allPaths.has(`ppt/media/image${n}.${ext}`)) n++;
    return `ppt/media/image${n}.${ext}`;
  }

  /** Ensure [Content_Types].xml has a Default entry for the given extension. */
  private ensureContentTypeForExtension(ext: string, mimeType: string): void {
    let ct = this.files.get('[Content_Types].xml');
    if (!ct) return;
    if (ct.includes(`Extension="${ext}"`)) return;
    const entry = `<Default Extension="${ext}" ContentType="${mimeType}"/>`;
    ct = ct.replace('</Types>', entry + '</Types>');
    this.persistFile('[Content_Types].xml', ct);
  }

  /** Check if a media file is referenced by any slide other than the given one. */
  private isMediaReferencedElsewhere(mediaPath: string, excludeSlideIdx: number): boolean {
    const filename = mediaPath.split('/').pop() ?? '';
    const slideCount = this.exports.get_slide_count();
    for (let i = 0; i < slideCount; i++) {
      if (i === excludeSlideIdx) continue;
      const relsPath = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
      const relsXml = this.files.get(relsPath) ?? '';
      if (relsXml.includes(filename)) return true;
    }
    return false;
  }

  /** Update a file in both the live files map and the export tracking map. */
  private persistFile(path: string, content: string): void {
    this.files.set(path, content);
    this.addedFiles.set(path, content);
  }

  /** Re-initialize the Wasm engine after structural changes to the files map. */
  private reinitializeWasm(): void {
    const result = this.exports.initialize_pptx();
    if (result.startsWith('ERROR:')) {
      throw new Error(`Re-initialization failed: ${result.slice(6)}`);
    }
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /** Build the Wasm import object that satisfies MoonBit's FFI declarations. */
  private buildImportObject(): Record<string, Record<string, unknown>> {
    return {
      'pptx_ffi': {
        get_file:         (path: string) => this.files.get(path) ?? '',
        get_entry_list:   () => [...this.files.keys(), ...this.rawFiles.keys()].join('\n'),
        get_file_base64:  (path: string) => bytesToBase64(this.rawFiles.get(path)),
        char_code_to_str: (n: number) => String.fromCodePoint(n),
        log:   (msg: string) => this.log.debug(msg),
        warn:  (msg: string) => this.log.warn(msg),
        error: (msg: string) => this.log.error(msg),
        measure_text: (text: string, fontFace: string, fontSizePt: number) =>
          this.measureText(text, fontFace, fontSizePt),
        get_font_fallback: (font: string) => this.fontFallbackCache.get(font) ?? '',
        convert_emf: (path: string) => this.convertEmf(path),
        math_sin:   (x: number) => Math.sin(x),
        math_cos:   (x: number) => Math.cos(x),
        math_atan2: (y: number, x: number) => Math.atan2(y, x),
        math_sqrt:  (x: number) => Math.sqrt(x),
      },
      'moonbit:ffi': {
        make_closure: (f: Function, ctx: unknown) => f.bind(null, ctx),
      },
    };
  }

  /** Convert an EMF/WMF file to an SVG data URI. Returns "" if conversion fails. */
  private convertEmf(path: string): string {
    const bytes = this.rawFiles.get(path);
    if (!bytes) return '';
    const svg = path.toLowerCase().endsWith('.wmf') ? wmfToSvg(bytes) : emfToSvg(bytes);
    if (!svg) return '';
    const encoded = encodeURIComponent(svg);
    return `data:image/svg+xml,${encoded}`;
  }

  /** Quote a font name for use in CSS font shorthand (names with spaces need quotes). */
  private static quoteFontName(name: string): string {
    const n = name.trim();
    if (!n || n === 'sans-serif' || n === 'serif' || n === 'monospace') return n;
    if (n.includes(' ') && !n.startsWith("'") && !n.startsWith('"')) return `'${n}'`;
    return n;
  }

  /** Measure the rendered pixel width of text. */
  private measureText(text: string, fontFace: string, fontSizePx: number): number {
    if (this.measureTextFn) {
      return this.measureTextFn(text, fontFace, fontSizePx);
    }
    // Default: Canvas 2D (browser only)
    if (!this.ctx) {
      if (typeof document === 'undefined') {
        // Non-browser environment — return approximate width
        return text.length * fontSizePx * 0.6;
      }
      this.canvas = document.createElement('canvas');
      this.ctx = this.canvas.getContext('2d')!;
    }
    // Use only the primary font (quoted if needed) for measurement.
    // Adding fallback fonts changes Canvas 2D metrics and breaks wrap accuracy.
    const quoted = PptxRenderer.quoteFontName(fontFace) || 'sans-serif';
    this.ctx.font = `${fontSizePx}px ${quoted}`;
    return this.ctx.measureText(text).width;
  }

  // ── Notes & Comments helpers ─────────────────────────────────────────────

  /**
   * Resolve a relationship target path for a slide.
   * @param slideIdx - 0-indexed slide index
   * @param relType - last segment of the relationship type (e.g. 'notesSlide', 'comments')
   */
  private resolveRelTarget(slideIdx: number, relType: string): string | null {
    const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
    const relsXml = this.files.get(relsPath);
    if (!relsXml) return null;

    // Find Relationship with matching Type
    const rt = escapeRegex(relType);
    const re = new RegExp(
      `<Relationship[^>]+Type="[^"]*/${rt}"[^>]+Target="([^"]+)"`,
    );
    const m = relsXml.match(re);
    if (!m) {
      // Try reversed attribute order (Target before Type)
      const re2 = new RegExp(
        `<Relationship[^>]+Target="([^"]+)"[^>]+Type="[^"]*/${rt}"`,
      );
      const m2 = relsXml.match(re2);
      if (!m2) return null;
      return this.resolveRelPath('ppt/slides/', m2[1]);
    }
    return this.resolveRelPath('ppt/slides/', m[1]);
  }

  /** Resolve a relative path like "../notesSlides/notesSlide1.xml" from a base dir. */
  private resolveRelPath(baseDir: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1); // absolute
    const parts = (baseDir + target).split('/');
    const resolved: string[] = [];
    for (const p of parts) {
      if (p === '..') resolved.pop();
      else if (p && p !== '.') resolved.push(p);
    }
    return resolved.join('/');
  }

  /**
   * Decode XML entity references in a raw inner-text string. Returns plain
   * text suitable for `.textContent`. Callers must NOT pass the result to
   * `.innerHTML` — use `.textContent` to avoid HTML re-interpretation.
   */
  private decodeXmlEntities(s: string): string {
    return s.replace(/&(#x[0-9A-Fa-f]+|#\d+|amp|lt|gt|quot|apos);/g, (_, ent: string) => {
      if (ent === 'amp') return '&';
      if (ent === 'lt') return '<';
      if (ent === 'gt') return '>';
      if (ent === 'quot') return '"';
      if (ent === 'apos') return "'";
      if (ent.startsWith('#x') || ent.startsWith('#X')) {
        const code = parseInt(ent.slice(2), 16);
        return Number.isFinite(code) && code > 0 ? String.fromCodePoint(code) : '';
      }
      if (ent.startsWith('#')) {
        const code = parseInt(ent.slice(1), 10);
        return Number.isFinite(code) && code > 0 ? String.fromCodePoint(code) : '';
      }
      return '';
    });
  }

  /** Extract text paragraphs from a notesSlide XML (body placeholder). */
  private extractNotesText(xml: string): string[] {
    const paragraphs: string[] = [];
    const bodyMatch = xml.match(
      /<p:sp\b[^]*?<p:ph[^>]+type="body"[^]*?<\/p:sp>/,
    );
    if (!bodyMatch) return paragraphs;
    const bodyXml = bodyMatch[0];

    const paraRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
    let pm: RegExpExecArray | null;
    while ((pm = paraRegex.exec(bodyXml)) !== null) {
      const paraContent = pm[1];
      const texts: string[] = [];
      const tRegex = /<a:t>([\s\S]*?)<\/a:t>/g;
      let tm: RegExpExecArray | null;
      while ((tm = tRegex.exec(paraContent)) !== null) {
        texts.push(this.decodeXmlEntities(tm[1]));
      }
      if (texts.length > 0) {
        paragraphs.push(texts.join(''));
      }
    }
    return paragraphs;
  }

  /** Parse comments XML into SlideComment array. */
  private parseComments(xml: string): SlideComment[] {
    const comments: SlideComment[] = [];
    const cmRegex = /<p:cm\b([^>]*)>([\s\S]*?)<\/p:cm>/g;
    let cm: RegExpExecArray | null;
    while ((cm = cmRegex.exec(xml)) !== null) {
      const attrs = cm[1];
      const body = cm[2];

      const authorId = parseInt(attrs.match(/authorId="(\d+)"/)?.[1] ?? '0');
      const dt = attrs.match(/dt="([^"]+)"/)?.[1] ?? '';
      const idx = parseInt(attrs.match(/idx="(\d+)"/)?.[1] ?? '0');

      const posMatch = body.match(/<p:pos\s+x="(\d+)"\s+y="(\d+)"/);
      const x = parseInt(posMatch?.[1] ?? '0');
      const y = parseInt(posMatch?.[2] ?? '0');

      const textMatch = body.match(/<p:text>([\s\S]*?)<\/p:text>/);
      const text = this.decodeXmlEntities(textMatch?.[1] ?? '');

      comments.push({ authorId, date: dt, index: idx, text, x, y });
    }
    return comments;
  }

  /** Parse commentAuthors XML into CommentAuthor array. */
  private parseCommentAuthors(xml: string): CommentAuthor[] {
    const authors: CommentAuthor[] = [];
    const authorRegex = /<p:cmAuthor\b([^>]*)\/?>/g;
    let am: RegExpExecArray | null;
    while ((am = authorRegex.exec(xml)) !== null) {
      const attrs = am[1];
      const id = parseInt(attrs.match(/id="(\d+)"/)?.[1] ?? '0');
      const name = this.decodeXmlEntities(attrs.match(/name="([^"]+)"/)?.[1] ?? '');
      const initials = this.decodeXmlEntities(attrs.match(/initials="([^"]+)"/)?.[1] ?? '');
      authors.push({ id, name, initials });
    }
    return authors;
  }
}
