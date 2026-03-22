/**
 * PptxRenderer — main API class for rendering PPTX files.
 *
 * Handles Wasm lifecycle, PPTX loading, SVG rendering, and export.
 */

import { bytesToBase64 } from './utils.js';
import { emfToSvg } from './emf-converter.js';
import { instantiateWasmWithFallback } from './wasm-compat.js';
import { extractZip, buildZip } from './zip.js';
import { DEFAULT_FONT_FALLBACKS } from './font-fallbacks.js';
import type { FontFallbackMap } from './font-fallbacks.js';

/** Wasm exports provided by the MoonBit module. */
interface PptxWasmExports {
  initialize_pptx(): string;
  get_slide_count(): number;
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
}

/** Options for text measurement callback. Font size is in CSS pixels (px). */
export interface MeasureTextFn {
  (text: string, fontFace: string, fontSizePx: number): number;
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

export class PptxRenderer {
  private wasm: WebAssembly.Instance | null = null;

  /** Decompressed text ZIP entries (path → UTF-8 string) */
  private files = new Map<string, string>();

  /** Raw binary ZIP entries (path → bytes) */
  private rawFiles = new Map<string, Uint8Array>();

  /** Original PPTX bytes for export */
  private originalBuffer: ArrayBuffer | null = null;

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
   * @param wasmSource - Optional URL string or ArrayBuffer of .wasm bytes.
   *                     If omitted, the bundled Wasm is used.
   */
  async init(wasmSource?: string | ArrayBuffer): Promise<void> {
    let bytes: ArrayBuffer;
    if (wasmSource instanceof ArrayBuffer) {
      bytes = wasmSource;
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

    this.log.debug(`Exporting PPTX with ${modifications.size} modified entries`);
    return buildZip(this.originalBuffer, modifications);
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

  /** Convert an EMF file to an SVG data URI. Returns "" if conversion fails. */
  private convertEmf(path: string): string {
    const bytes = this.rawFiles.get(path);
    if (!bytes) return '';
    const svg = emfToSvg(bytes);
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

  /** Build a CSS font shorthand-safe font-family string with fallbacks. */
  private buildCssFontFamily(fontFace: string): string {
    const names: string[] = [];
    // Add the primary font
    if (fontFace) names.push(PptxRenderer.quoteFontName(fontFace));
    // Add fallbacks from cache (comma-separated string)
    const fb = this.fontFallbackCache.get(fontFace);
    if (fb) {
      for (const f of fb.split(',')) {
        const q = PptxRenderer.quoteFontName(f);
        if (q) names.push(q);
      }
    }
    names.push('sans-serif');
    return names.join(', ');
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
}
