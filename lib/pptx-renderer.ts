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
}

/** Options for text measurement callback. Font size is in CSS pixels (px). */
export interface MeasureTextFn {
  (text: string, fontFace: string, fontSizePx: number): number;
}

/** Options for initializing PptxRenderer. */
export interface PptxRendererOptions {
  /** Custom text measurement function. If not provided, uses Canvas 2D (browser only). */
  measureText?: MeasureTextFn;
  /**
   * Custom font fallback mappings. Merged with built-in defaults (lib/font-fallbacks.ts).
   * User entries override built-in entries for the same font name.
   */
  fontFallbacks?: FontFallbackMap;
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

  constructor(options?: PptxRendererOptions) {
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

    const result = await instantiateWasmWithFallback(bytes, this.buildImportObject());
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
    console.log('[pptx] Parsing ZIP archive...');
    const { textFiles, binaryFiles } = await extractZip(arrayBuffer);
    this.files = textFiles;
    this.rawFiles = binaryFiles;
    console.log(`[pptx] Extracted ${textFiles.size} text entries, ${binaryFiles.size} binary entries`);

    const result = this.exports.initialize_pptx();
    console.log('[pptx] initialize_pptx result:', result);

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

    console.log(`[pptx] Exporting PPTX with ${modifications.size} modified entries`);
    return buildZip(this.originalBuffer, modifications);
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /** Build the Wasm import object that satisfies MoonBit's FFI declarations. */
  private buildImportObject(): Record<string, Record<string, unknown>> {
    return {
      'pptx_ffi': {
        get_file:         (path: string) => this.files.get(path) ?? '',
        get_entry_list:   () => [...this.files.keys(), ...this.rawFiles.keys()].join('\n'),
        get_file_base64:  (path: string) => bytesToBase64(this.rawFiles.get(path)),
        char_code_to_str: (n: number) => String.fromCharCode(n),
        log:   (msg: string) => console.log('[pptx]', msg),
        warn:  (msg: string) => console.warn('[pptx]', msg),
        error: (msg: string) => console.error('[pptx]', msg),
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
    this.ctx.font = `${fontSizePx}px ${fontFace}`;
    return this.ctx.measureText(text).width;
  }
}
