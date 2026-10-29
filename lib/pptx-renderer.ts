
import { bytesToBase64, base64ToBytes } from './utils.js';
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
  restore_slide_ooxml(slideIdx: number, xml: string): string;
  get_text_layout(slideIdx: number, shapeIdx: number): string;
  hit_test_text(slideIdx: number, shapeIdx: number, xEmu: number, yEmu: number): string;
  replace_text_range(slideIdx: number, shapeIdx: number,
    startPara: number, startChar: number, endPara: number, endChar: number, newText: string): string;
  bring_to_front(slideIdx: number, shapeIdx: number): string;
  send_to_back(slideIdx: number, shapeIdx: number): string;
  bring_forward(slideIdx: number, shapeIdx: number): string;
  send_backward(slideIdx: number, shapeIdx: number): string;
  update_shapes_transform(slideIdx: number, itemsData: string): string;
  get_shape_ooxml(slideIdx: number, shapeIdx: number): string;
  add_shape_from_ooxml(slideIdx: number, shapeXml: string, dxEmu: number, dyEmu: number): string;
  update_table_cell_text(slideIdx: number, shapeIdx: number, row: number, col: number, text: string): string;
  add_table_row(slideIdx: number, shapeIdx: number, afterRow: number): string;
  delete_table_row(slideIdx: number, shapeIdx: number, row: number): string;
  add_table_column(slideIdx: number, shapeIdx: number, afterCol: number, widthEmu: number): string;
  delete_table_column(slideIdx: number, shapeIdx: number, col: number): string;
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
  /** Font fallbacks merged over the built-in defaults (same-name entries win). */
  fontFallbacks?: FontFallbackMap;
  /** Console verbosity, from `'silent'` to `'debug'`. Default: `'error'`. */
  logLevel?: LogLevel;
  /** Undo steps kept (default 50); `0` disables history. */
  maxHistory?: number;
  /**
   * Date for auto-updating date placeholders (`<a:fld type="datetime1">`); fix it for
   * reproducible output. Default: today via `toLocaleDateString()`.
   */
  currentDate?: Date | string | (() => Date | string);
}

/**
 * Result of {@link PptxRenderer.undo} / {@link PptxRenderer.redo} (as JSON): the
 * 0-indexed slides to re-render, and the slide count afterwards.
 */
export interface HistoryResult {
  slides: number[];
  slideCount: number;
}

/** A single glyph box in {@link TextLayout} (EMU, absolute). */
export interface GlyphBox { x: number; w: number }

/** A run fragment box within a {@link TextLine} (EMU, absolute). */
export interface TextRunBox {
  paraIdx: number;
  runIdx: number;
  x: number;
  w: number;
  /** Char offset within the run where this box begins. */
  runCharStart: number;
  /** Char offset within the paragraph where this box begins. */
  paraCharStart: number;
  chars: GlyphBox[];
}

/** A visual line in {@link TextLayout} (EMU). */
export interface TextLine {
  paraIdx: number;
  y: number;
  h: number;
  runs: TextRunBox[];
}

/** Text geometry from {@link PptxRenderer.getTextLayout} (EMU). */
export interface TextLayout {
  box: { x: number; y: number; cx: number; cy: number };
  lines: TextLine[];
}

/**
 * A shape's OOXML plus its images inline (base64), from {@link PptxRenderer.getShapeSpec}
 * for {@link PptxRenderer.insertShapeSpec}; passed as a JSON string.
 */
export interface ShapeSpec {
  xml: string;
  media: Array<{ rid: string; ext: string; mime: string; b64: string }>;
}

/** Caret position returned by {@link PptxRenderer.hitTestText} (parsed from JSON). */
export interface TextHit {
  paraIdx: number;
  runIdx: number;
  /** Char offset within the run. */
  charOffset: number;
  /** Char offset within the paragraph (use with `replaceTextRange`). */
  paraOffset: number;
}

/**
 * Document state diverging from the original ZIP, for undo/redo. Contents are shared
 * by reference: `persistFile` replaces map entries instead of mutating them.
 */
interface DocSnapshot {
  files: Map<string, string>;
  addedFiles: Map<string, string>;
  removedFiles: Set<string>;
  addedBinaryFiles: Map<string, Uint8Array>;
  /** Pending Wasm-side slide edits: [slideIdx, slide OOXML]. */
  modifiedSlides: Array<[number, string]>;
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

/** Escape RegExp metacharacters: rIds and targets come from untrusted PPTX input. */
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

/** Map a media file extension to a MIME type (used for cross-slide paste). */
function extToMime(ext: string): string {
  const map: Record<string, string> = {
    png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif',
    svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp',
    tiff: 'image/tiff', tif: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf',
  };
  return map[ext.toLowerCase()] ?? 'application/octet-stream';
}

/** First slide ID used in presentation.xml sldIdLst (OOXML convention). */
const FIRST_SLIDE_ID = 256;

const RE_SLIDE_FILE = /^ppt\/slides\/slide\d+\.xml$/;
const RE_SLIDE_RELS_FILE = /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/;

/** Default number of undo steps retained (overridable via `maxHistory` option). */
const DEFAULT_MAX_HISTORY = 50;

/** Relationship-ID attribute names found on shapes (image/link refs), as a regex source. */
const RID_ATTR_PATTERN = 'r:(?:embed|link|id)';

export class PptxRenderer {
  private wasm: WebAssembly.Instance | null = null;

  private files = new Map<string, string>();

  private rawFiles = new Map<string, Uint8Array>();

  private originalBuffer: ArrayBuffer | null = null;

  /** Files added after loadPptx (not in original ZIP) */
  private addedFiles = new Map<string, string>();

  /** Files removed after loadPptx */
  private removedFiles = new Set<string>();

  /** Binary files added/replaced after loadPptx (e.g. images) */
  private addedBinaryFiles = new Map<string, Uint8Array>();

  private canvas: HTMLCanvasElement | null = null;
  private ctx: CanvasRenderingContext2D | null = null;

  private measureTextFn: MeasureTextFn | null = null;

  /** Font fallback lookup map (source → comma-separated fallbacks) */
  private fontFallbackCache = new Map<string, string>();

  private log: Logger;

  /** Past document states (top = most recent state before the current one). */
  private undoStack: DocSnapshot[] = [];

  private redoStack: DocSnapshot[] = [];

  private batchDepth = 0;
  /** Whether the current batch has already recorded its pre-edit snapshot. */
  private batchRecorded = false;

  private maxHistory: number;

  /** Source for date-placeholder auto-content (see PptxRendererOptions.currentDate). */
  private currentDateOption?: Date | string | (() => Date | string);

  constructor(options?: PptxRendererOptions) {
    this.log = createLogger(options?.logLevel ?? 'error');
    this.maxHistory = Math.max(0, options?.maxHistory ?? DEFAULT_MAX_HISTORY);
    this.currentDateOption = options?.currentDate;
    if (options?.measureText) {
      this.measureTextFn = options.measureText;
    }
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
   * Load the Wasm module: by default the bundled `main.wasm` (resolved via `import.meta.url`),
   * or a URL / Node.js file path, or the `.wasm` bytes.
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

  /** Load a PPTX file. */
  async loadPptx(arrayBuffer: ArrayBuffer): Promise<{ slideCount: number }> {
    if (!this.wasm) {
      throw new Error('Wasm not initialized — wait for init() to complete before loading files.');
    }
    this.originalBuffer = arrayBuffer.slice(0);
    this.addedFiles.clear();
    this.addedBinaryFiles.clear();
    this.removedFiles.clear();
    this.clearHistory();
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

  /** Render a slide (0-indexed) as SVG markup, or `"ERROR:..."`. */
  renderSlideSvg(slideIdx: number): string {
    return this.exports.render_slide_svg(slideIdx);
  }

  /** Replace a slide's data with an edited SVG (read back from its `data-ooxml-*`). Returns `"OK"` or `"ERROR:..."`. */
  updateSlideFromSvg(slideIdx: number, svgString: string): string {
    return this.exports.update_slide_from_svg(slideIdx, svgString);
  }

  /** A slide's OOXML: the edited version if the slide changed, otherwise the original. */
  getSlideOoxml(slideIdx: number): string {
    return this.exports.get_slide_ooxml(slideIdx);
  }

  /** Export the presentation, with all edits, as PPTX bytes. */
  async exportPptx(): Promise<ArrayBuffer> {
    if (!this.originalBuffer) {
      throw new Error('No PPTX loaded — call loadPptx() first.');
    }

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

  /**
   * Group the mutations until the matching {@link endBatch} into one undo step (e.g. a
   * paste that adds a shape and sets its text). Nested calls are flattened.
   */
  beginBatch(): void {
    if (this.batchDepth === 0) this.batchRecorded = false;
    this.batchDepth++;
  }

  /** End a batch started with {@link beginBatch}. */
  endBatch(): void {
    if (this.batchDepth > 0) this.batchDepth--;
  }

  /** Whether there is at least one step that can be undone. */
  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  /** Whether there is at least one step that can be redone. */
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Discard all undo/redo history (e.g. after loading a new file). */
  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.batchDepth = 0;
    this.batchRecorded = false;
  }

  /** Undo the last edit or batch. Returns a JSON {@link HistoryResult} or `"ERROR:nothing to undo"`. */
  undo(): string {
    if (this.undoStack.length === 0) return 'ERROR:nothing to undo';
    const current = this.captureSnapshot();
    const prev = this.undoStack.pop()!;
    this.redoStack.push(current);
    this.restoreSnapshot(prev);
    return JSON.stringify(this.diffSnapshots(current, prev));
  }

  /** Redo the last undone edit or batch. Returns a JSON {@link HistoryResult} or `"ERROR:nothing to redo"`. */
  redo(): string {
    if (this.redoStack.length === 0) return 'ERROR:nothing to redo';
    const current = this.captureSnapshot();
    const next = this.redoStack.pop()!;
    this.undoStack.push(current);
    this.restoreSnapshot(next);
    return JSON.stringify(this.diffSnapshots(current, next));
  }

  /** Record the pre-edit state; called first by every mutating method (once per batch). */
  private checkpoint(): void {
    if (!this.wasm || this.maxHistory === 0) return;
    this.commitCheckpoint(this.captureSnapshot());
  }

  /**
   * Push a snapshot taken before a fallible op, once the op succeeded, so a failed edit
   * leaves no phantom undo step. Returns whether a step was recorded.
   */
  private commitCheckpoint(snap: DocSnapshot): boolean {
    if (this.maxHistory === 0) return false;
    if (this.batchDepth > 0) {
      if (this.batchRecorded) return false;
      this.batchRecorded = true;
    }
    this.undoStack.push(snap);
    this.redoStack = [];
    while (this.undoStack.length > this.maxHistory) {
      this.undoStack.shift();
    }
    return true;
  }

  /** Run `op` (returning `"OK..."`/SVG or `"ERROR:..."`) and record history only if it succeeded. */
  private runGuarded(op: () => string): string {
    const snap = (this.wasm && this.maxHistory > 0) ? this.captureSnapshot() : null;
    const result = op();
    if (snap && !result.startsWith('ERROR')) this.commitCheckpoint(snap);
    return result;
  }

  /** Capture a shallow snapshot of all state diverging from the original ZIP. */
  private captureSnapshot(): DocSnapshot {
    const modifiedSlides: Array<[number, string]> = [];
    const modifiedStr = this.exports.get_modified_entries();
    if (modifiedStr) {
      for (const line of modifiedStr.split('\n')) {
        if (!line) continue;
        const tabIdx = line.indexOf('\t');
        if (tabIdx < 0) continue;
        const path = line.substring(0, tabIdx);
        const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (!m) continue;
        modifiedSlides.push([parseInt(m[1]) - 1, line.substring(tabIdx + 1)]);
      }
    }
    return {
      files: new Map(this.files),
      addedFiles: new Map(this.addedFiles),
      removedFiles: new Set(this.removedFiles),
      addedBinaryFiles: new Map(this.addedBinaryFiles),
      modifiedSlides,
    };
  }

  /** Restore a previously captured snapshot, rebuilding Wasm state. */
  private restoreSnapshot(s: DocSnapshot): void {
    this.files = new Map(s.files);
    this.addedFiles = new Map(s.addedFiles);
    this.removedFiles = new Set(s.removedFiles);
    this.addedBinaryFiles = new Map(s.addedBinaryFiles);
    this.reinitializeWasm();
    for (const [slideIdx, xml] of s.modifiedSlides) {
      const r = this.exports.restore_slide_ooxml(slideIdx, xml);
      if (r.startsWith('ERROR:')) {
        this.log.warn(`restore_slide_ooxml(${slideIdx}) failed: ${r}`);
      }
    }
  }

  /** Compute which slides differ between two snapshots, for re-render hints. */
  private diffSnapshots(from: DocSnapshot, to: DocSnapshot): HistoryResult {
    const slides = new Set<number>();
    const fromMap = new Map(from.modifiedSlides);
    const toMap = new Map(to.modifiedSlides);
    for (const [idx, xml] of toMap) {
      if (fromMap.get(idx) !== xml) slides.add(idx);
    }
    for (const [idx] of fromMap) {
      if (!toMap.has(idx)) slides.add(idx);
    }
    // Structural change (slide files added/removed) → flag all slides.
    const slideCount = this.getSlideCount();
    if (from.files.size !== to.files.size) {
      for (let i = 0; i < slideCount; i++) slides.add(i);
    }
    return { slides: [...slides].sort((a, b) => a - b), slideCount };
  }

  /** Speaker notes of a slide, one string per paragraph (`[]` if none). */
  getSlideNotes(slideIdx: number): string[] {
    const notesPath = this.resolveRelTarget(slideIdx, 'notesSlide');
    if (!notesPath) return [];
    const xml = this.files.get(notesPath);
    if (!xml) return [];
    return this.extractNotesText(xml);
  }

  /** Comments on a slide (`[]` if none). */
  getSlideComments(slideIdx: number): SlideComment[] {
    const commentsPath = this.resolveRelTarget(slideIdx, 'comments');
    if (!commentsPath) return [];
    const xml = this.files.get(commentsPath);
    if (!xml) return [];
    return this.parseComments(xml);
  }

  /** Comment authors of the presentation (`[]` if none). */
  getCommentAuthors(): CommentAuthor[] {
    const xml = this.files.get('ppt/commentAuthors.xml');
    if (!xml) return [];
    return this.parseCommentAuthors(xml);
  }

  /** Render one shape as an SVG fragment (`<defs>` + `<g>`), or `"ERROR:..."`. */
  renderShapeSvg(slideIdx: number, shapeIdx: number): string {
    return this.exports.render_shape_svg(slideIdx, shapeIdx);
  }

  /** Set a shape's position, size and rotation (EMU); returns the re-rendered shape SVG. */
  updateShapeTransform(slideIdx: number, shapeIdx: number,
    x: number, y: number, cx: number, cy: number, rot: number): string {
    this.checkpoint();
    return this.exports.update_shape_transform(slideIdx, shapeIdx, x, y, cx, cy, rot);
  }

  /**
   * Set several shapes' transforms (EMU, rot in 1/60000°) as one undo step. All indices are
   * validated first, so a bad one changes nothing. Returns `"OK:<count>"` or `"ERROR:..."`.
   */
  updateShapesTransform(slideIdx: number,
    items: Array<{ shapeIdx: number; x: number; y: number; cx: number; cy: number; rot: number }>): string {
    const data = items.map(it => `${it.shapeIdx},${it.x},${it.y},${it.cx},${it.cy},${it.rot}`).join(';');
    return this.runGuarded(() => this.exports.update_shapes_transform(slideIdx, data));
  }

  /** Set a text run's text; returns the re-rendered shape SVG. */
  updateShapeText(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, text: string): string {
    this.checkpoint();
    return this.exports.update_shape_text(slideIdx, shapeIdx, paraIdx, runIdx, text);
  }

  /** Set a shape's solid fill (RGB 0-255); returns the re-rendered shape SVG. */
  updateShapeFill(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number): string {
    this.checkpoint();
    return this.exports.update_shape_fill(slideIdx, shapeIdx, r, g, b);
  }

  /** Delete a shape. Returns `"OK"` or `"ERROR:..."`. */
  deleteShape(slideIdx: number, shapeIdx: number): string {
    this.checkpoint();
    return this.exports.delete_shape(slideIdx, shapeIdx);
  }

  /**
   * Add a preset AutoShape (`"rect"`, `"ellipse"`, `"roundRect"`, `"line"`, …) at an EMU
   * position; a fill component of -1 means no fill. Returns `"OK:<shapeIdx>"` or `"ERROR:..."`.
   */
  addShape(slideIdx: number, geomType: string,
    x: number, y: number, cx: number, cy: number,
    fillR = -1, fillG = -1, fillB = -1): string {
    this.checkpoint();
    return this.exports.add_shape(slideIdx, geomType, x, y, cx, cy, fillR, fillG, fillB);
  }

  /**
   * Add a one-run paragraph to a shape. `fontSize` is in hundredths of a point (0 = inherit);
   * a color component of -1 inherits. Returns `"OK:<paraIdx>"` or `"ERROR:..."`.
   */
  addShapeText(slideIdx: number, shapeIdx: number, text: string,
    fontSize = 0, colorR = -1, colorG = -1, colorB = -1): string {
    this.checkpoint();
    return this.exports.add_shape_text(slideIdx, shapeIdx, text, fontSize, colorR, colorG, colorB);
  }

  /** Duplicate a shape, offset by (dxEmu, dyEmu). Returns `"OK:<shapeIdx>"` or `"ERROR:..."`. */
  duplicateShape(slideIdx: number, shapeIdx: number,
    dxEmu = 457200, dyEmu = 457200): string {
    this.checkpoint();
    return this.exports.duplicate_shape(slideIdx, shapeIdx, dxEmu, dyEmu);
  }

  /**
   * Set a linear gradient fill: `angle` in 60000ths of a degree, stops `{ pos, r, g, b }`
   * with `pos` 0-100000. Returns the re-rendered shape SVG.
   */
  updateShapeGradientFill(slideIdx: number, shapeIdx: number,
    angle: number, stops: Array<{ pos: number; r: number; g: number; b: number }>): string {
    this.checkpoint();
    const stopsData = stops.map(s => `${s.pos},${s.r},${s.g},${s.b}`).join(';');
    return this.exports.update_shape_gradient_fill(slideIdx, shapeIdx, angle, stopsData);
  }

  /**
   * Set a shape's outline: RGB 0-255 (r = -1 removes it), width in EMU, dash preset
   * (`"dash"`, `"dot"`, …; `""` = solid). Returns the re-rendered shape SVG.
   */
  updateShapeStroke(slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number, widthEmu = 12700, dash = ''): string {
    this.checkpoint();
    return this.exports.update_shape_stroke(slideIdx, shapeIdx, r, g, b, widthEmu, dash);
  }

  /** Add a one-run paragraph; `align` is `"l"`/`"ctr"`/`"r"`/`"just"` or `""` (inherit). Returns `"OK:<paraIdx>"`. */
  addParagraph(slideIdx: number, shapeIdx: number, text: string, align = ''): string {
    this.checkpoint();
    return this.exports.add_paragraph(slideIdx, shapeIdx, text, align);
  }

  /** Delete a paragraph. Returns `"OK"` or `"ERROR:..."`. */
  deleteParagraph(slideIdx: number, shapeIdx: number, paraIdx: number): string {
    this.checkpoint();
    return this.exports.delete_paragraph(slideIdx, shapeIdx, paraIdx);
  }

  /** Append a text run to a paragraph. Returns `"OK:<runIdx>"` or `"ERROR:..."`. */
  addRun(slideIdx: number, shapeIdx: number, paraIdx: number, text: string): string {
    this.checkpoint();
    return this.exports.add_run(slideIdx, shapeIdx, paraIdx, text);
  }

  /** Delete a text run. Returns `"OK"` or `"ERROR:..."`. */
  deleteRun(slideIdx: number, shapeIdx: number, paraIdx: number, runIdx: number): string {
    this.checkpoint();
    return this.exports.delete_run(slideIdx, shapeIdx, paraIdx, runIdx);
  }

  /** Set bold / italic (1 = on, 0 = off, -1 = unchanged); returns the re-rendered shape SVG. */
  updateTextRunStyle(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, bold = -1, italic = -1): string {
    this.checkpoint();
    return this.exports.update_text_run_style(slideIdx, shapeIdx, paraIdx, runIdx, bold, italic);
  }

  /** Set a run's size in hundredths of a point (0 = inherit); returns the re-rendered shape SVG. */
  updateTextRunFontSize(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontSize: number): string {
    this.checkpoint();
    return this.exports.update_text_run_font_size(slideIdx, shapeIdx, paraIdx, runIdx, fontSize);
  }

  /** Set a run's color (RGB 0-255; r = -1 inherits); returns the re-rendered shape SVG. */
  updateTextRunColor(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, r: number, g: number, b: number): string {
    this.checkpoint();
    return this.exports.update_text_run_color(slideIdx, shapeIdx, paraIdx, runIdx, r, g, b);
  }

  /** Set a run's Latin / East Asian / complex-script fonts (`""` = unchanged); returns the re-rendered shape SVG. */
  updateTextRunFont(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, fontFace = '', eaFont = '', csFont = ''): string {
    this.checkpoint();
    return this.exports.update_text_run_font(slideIdx, shapeIdx, paraIdx, runIdx, fontFace, eaFont, csFont);
  }

  /** Set alignment: `"l"`/`"ctr"`/`"r"`/`"just"` or `""` (inherit); returns the re-rendered shape SVG. */
  updateParagraphAlign(slideIdx: number, shapeIdx: number,
    paraIdx: number, align: string): string {
    this.checkpoint();
    return this.exports.update_paragraph_align(slideIdx, shapeIdx, paraIdx, align);
  }

  /**
   * Set underline (`"sng"`/`"dbl"`), strike (`"sngStrike"`/`"dblStrike"`) — `""` unchanged,
   * `"none"` removes — and baseline (30000 super, -25000 sub, 0 normal, -1 unchanged).
   */
  updateTextRunDecoration(slideIdx: number, shapeIdx: number,
    paraIdx: number, runIdx: number, underline = '', strike = '', baseline = -1): string {
    this.checkpoint();
    return this.exports.update_text_run_decoration(slideIdx, shapeIdx, paraIdx, runIdx, underline, strike, baseline);
  }

  /**
   * A shape's text geometry for caret/selection drawing: JSON {@link TextLayout} (EMU), or
   * `"ERROR:..."`. Horizontal text only; vertical/warp/math/multi-column give just the box.
   */
  getTextLayout(slideIdx: number, shapeIdx: number): string {
    return this.exports.get_text_layout(slideIdx, shapeIdx);
  }

  /**
   * The caret position at an EMU point: JSON {@link TextHit} or `"ERROR:..."`. Pass its
   * `paraOffset` to {@link replaceTextRange}.
   */
  hitTestText(slideIdx: number, shapeIdx: number, xEmu: number, yEmu: number): string {
    return this.exports.hit_test_text(slideIdx, shapeIdx, xEmu, yEmu);
  }

  /**
   * Replace the paragraph-offset range with `newText`, keeping the boundary runs' formatting.
   * start == end inserts, `""` deletes, `\n` splits paragraphs. Returns the shape SVG.
   */
  replaceTextRange(slideIdx: number, shapeIdx: number,
    startPara: number, startChar: number, endPara: number, endChar: number, newText: string): string {
    this.checkpoint();
    return this.exports.replace_text_range(slideIdx, shapeIdx, startPara, startChar, endPara, endChar, newText);
  }

  /** Move a shape to the front. Returns `"OK:<newShapeIdx>"` (its index changes) or `"ERROR:..."`. */
  bringToFront(slideIdx: number, shapeIdx: number): string {
    this.checkpoint();
    return this.exports.bring_to_front(slideIdx, shapeIdx);
  }

  /** Move a shape to the back. Returns `"OK:<newShapeIdx>"` or `"ERROR:..."`. */
  sendToBack(slideIdx: number, shapeIdx: number): string {
    this.checkpoint();
    return this.exports.send_to_back(slideIdx, shapeIdx);
  }

  /** Move a shape one step forward. Returns `"OK:<newShapeIdx>"` or `"ERROR:..."`. */
  bringForward(slideIdx: number, shapeIdx: number): string {
    this.checkpoint();
    return this.exports.bring_forward(slideIdx, shapeIdx);
  }

  /** Move a shape one step backward. Returns `"OK:<newShapeIdx>"` or `"ERROR:..."`. */
  sendBackward(slideIdx: number, shapeIdx: number): string {
    this.checkpoint();
    return this.exports.send_backward(slideIdx, shapeIdx);
  }

  /**
   * A shape as a portable JSON {@link ShapeSpec} for copy/paste, or `"ERROR:..."`. Charts are
   * refused: a paste can't carry their chart part and embedded workbook.
   */
  getShapeSpec(slideIdx: number, shapeIdx: number): string {
    const xml = this.exports.get_shape_ooxml(slideIdx, shapeIdx);
    if (xml.startsWith('ERROR')) return xml;
    const relsXml = this.files.get(`ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`) ?? '';
    const media: ShapeSpec['media'] = [];
    const seen = new Set<string>();
    const re = new RegExp(`${RID_ATTR_PATTERN}="([^"]+)"`, 'g');
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) {
      const rid = m[1];
      if (seen.has(rid)) continue;
      seen.add(rid);
      const target = this.resolveRidTarget(relsXml, rid);
      if (!target) continue;
      const mediaPath = this.resolveRelPath('ppt/slides/', target);
      const bytes = this.rawFiles.get(mediaPath);
      if (!bytes) continue; // non-media rIds (e.g. hyperlink URLs) are skipped
      const ext = (mediaPath.split('.').pop() ?? '').toLowerCase();
      media.push({ rid, ext, mime: extToMime(ext), b64: bytesToBase64(bytes) });
    }
    return JSON.stringify({ xml, media });
  }

  /**
   * Paste a {@link ShapeSpec} onto a slide, offset by (dxEmu, dyEmu), re-adding its images
   * under fresh rIds. Undoable. Returns `"OK:<shapeIdx>"` or `"ERROR:..."`.
   */
  insertShapeSpec(slideIdx: number, spec: string, dxEmu = 0, dyEmu = 0): string {
    if (!this.wasm) return 'ERROR:not initialized';
    let parsed: ShapeSpec;
    try {
      parsed = JSON.parse(spec) as ShapeSpec;
    } catch {
      return 'ERROR:invalid shape spec';
    }
    if (!parsed || typeof parsed.xml !== 'string') return 'ERROR:invalid shape spec';

    return this.runGuarded(() => {
      let xml = parsed.xml;
      const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
      for (const item of parsed.media ?? []) {
        const ext = item.ext || 'png';
        const bytes = base64ToBytes(item.b64);
        const mediaPath = this.nextMediaPath(ext);
        this.rawFiles.set(mediaPath, bytes);
        this.addedBinaryFiles.set(mediaPath, bytes);
        this.ensureContentTypeForExtension(ext, item.mime || extToMime(ext));

        let relsXml = this.files.get(relsPath)
          ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_XMLNS}"></Relationships>`;
        const newRid = this.nextRid(relsXml);
        const relEntry = `<Relationship Id="${newRid}" Type="${REL_TYPE_IMAGE}" Target="../media/${mediaPath.split('/').pop()}"/>`;
        relsXml = relsXml.includes('</Relationships>')
          ? relsXml.replace('</Relationships>', relEntry + '</Relationships>')
          : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_XMLNS}">${relEntry}</Relationships>`;
        this.persistFile(relsPath, relsXml);

        // Re-link this image's rId in the shape XML.
        const ridRe = new RegExp(`(${RID_ATTR_PATTERN}=")${escapeRegex(item.rid)}(")`, 'g');
        xml = xml.replace(ridRe, `$1${newRid}$2`);
      }
      return this.exports.add_shape_from_ooxml(slideIdx, xml, dxEmu, dyEmu);
    });
  }

  /** Set a cell's text, keeping its first run's formatting. Returns the shape SVG or `"ERROR:..."`. */
  updateTableCellText(slideIdx: number, shapeIdx: number, row: number, col: number, text: string): string {
    this.checkpoint();
    return this.exports.update_table_cell_text(slideIdx, shapeIdx, row, col, text);
  }

  /**
   * Insert an empty row after `afterRow` (-1 = top). Returns `"OK:<newRowIdx>"` or `"ERROR:..."`.
   * Merged cells are not span-adjusted.
   */
  addTableRow(slideIdx: number, shapeIdx: number, afterRow = -1): string {
    this.checkpoint();
    return this.exports.add_table_row(slideIdx, shapeIdx, afterRow);
  }

  /** Delete a table row (at least one must remain). Returns `"OK"` or `"ERROR:..."`. */
  deleteTableRow(slideIdx: number, shapeIdx: number, row: number): string {
    this.checkpoint();
    return this.exports.delete_table_row(slideIdx, shapeIdx, row);
  }

  /**
   * Insert an empty column after `afterCol` (-1 = left); `widthEmu = 0` copies a neighbour's
   * width. Returns `"OK:<newColIdx>"` or `"ERROR:..."`. Merged cells are not span-adjusted.
   */
  addTableColumn(slideIdx: number, shapeIdx: number, afterCol = -1, widthEmu = 0): string {
    this.checkpoint();
    return this.exports.add_table_column(slideIdx, shapeIdx, afterCol, widthEmu);
  }

  /** Delete a table column (at least one must remain). Returns `"OK"` or `"ERROR:..."`. */
  deleteTableColumn(slideIdx: number, shapeIdx: number, col: number): string {
    this.checkpoint();
    return this.exports.delete_table_column(slideIdx, shapeIdx, col);
  }

  /**
   * Insert a blank slide after `afterIdx` (-1 = first; default last), using the layout of
   * `sourceSlideIdx` (default the last slide).
   */
  async addSlide(afterIdx?: number, sourceSlideIdx?: number): Promise<{ slideCount: number; insertedIdx: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const oldCount = this.exports.get_slide_count();
    if (oldCount === 0) throw new Error('No presentation loaded.');

    const insertIdx = afterIdx === undefined ? oldCount : afterIdx + 1;
    if (insertIdx < 0 || insertIdx > oldCount) {
      throw new Error(`Invalid insert position: ${insertIdx}`);
    }

    this.checkpoint();

    const srcIdx = sourceSlideIdx ?? Math.max(0, oldCount - 1);
    const srcRelsPath = `ppt/slides/_rels/slide${srcIdx + 1}.xml.rels`;
    const srcRelsXml = this.files.get(srcRelsPath) ?? '';
    const layoutTarget = this.extractRelTarget(srcRelsXml, '/slideLayout');

    const slideContents: string[] = [];
    const slideRels: string[] = [];
    for (let i = 0; i < oldCount; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    const slideSize = this.extractSlideSize();
    const blankSlideXml = this.createBlankSlideXml(slideSize.cx, slideSize.cy);

    const blankRels = layoutTarget
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_XMLNS}"><Relationship Id="rId1" Type="${REL_TYPE_SLIDE_LAYOUT}" Target="${layoutTarget}"/></Relationships>`
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_XMLNS}"></Relationships>`;

    slideContents.splice(insertIdx, 0, blankSlideXml);
    slideRels.splice(insertIdx, 0, blankRels);

    const newCount = oldCount + 1;

    this.rewriteSlideFiles(slideContents, slideRels, newCount);

    this.updatePresentationXmlForAdd(insertIdx);

    this.ensureContentTypeForSlide(newCount);

    this.reinitializeWasm();

    return { slideCount: newCount, insertedIdx: insertIdx };
  }

  /** Delete a slide; at least one must remain. */
  async deleteSlide(slideIdx: number): Promise<{ slideCount: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const oldCount = this.exports.get_slide_count();
    if (oldCount <= 1) throw new Error('Cannot delete the last remaining slide.');
    if (slideIdx < 0 || slideIdx >= oldCount) {
      throw new Error(`Slide index out of range: ${slideIdx}`);
    }

    this.checkpoint();

    const slideContents: string[] = [];
    const slideRels: string[] = [];
    for (let i = 0; i < oldCount; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    slideContents.splice(slideIdx, 1);
    slideRels.splice(slideIdx, 1);

    const newCount = oldCount - 1;

    // Mark old highest-numbered files for removal (in case they persist)
    const oldPath = `ppt/slides/slide${oldCount}.xml`;
    const oldRelsPath = `ppt/slides/_rels/slide${oldCount}.xml.rels`;
    this.removedFiles.add(oldPath);
    this.removedFiles.add(oldRelsPath);

    this.rewriteSlideFiles(slideContents, slideRels, newCount);

    this.updatePresentationXmlForDelete(slideIdx);

    this.reinitializeWasm();

    return { slideCount: newCount };
  }

  /** Reorder slides: `newOrder[i]` is the old index of the slide placed at `i` (a permutation). */
  async reorderSlides(newOrder: number[]): Promise<{ slideCount: number }> {
    if (!this.wasm) throw new Error('Not initialized — call init() and loadPptx() first.');

    const count = this.exports.get_slide_count();
    if (newOrder.length !== count) {
      throw new Error(`newOrder length (${newOrder.length}) must equal slide count (${count}).`);
    }

    const seen = new Set<number>();
    for (const idx of newOrder) {
      if (idx < 0 || idx >= count || seen.has(idx)) {
        throw new Error(`Invalid permutation: ${JSON.stringify(newOrder)}`);
      }
      seen.add(idx);
    }

    this.checkpoint();

    const slideContents: string[] = [];
    const slideRels: string[] = [];
    for (let i = 0; i < count; i++) {
      slideContents.push(this.files.get(`ppt/slides/slide${i + 1}.xml`) ?? '');
      slideRels.push(this.files.get(`ppt/slides/_rels/slide${i + 1}.xml.rels`) ?? '');
    }

    const reordered: string[] = newOrder.map(i => slideContents[i]);
    const reorderedRels: string[] = newOrder.map(i => slideRels[i]);

    this.rewriteSlideFiles(reordered, reorderedRels, count);

    this.updatePresentationXmlForReorder(newOrder);

    this.reinitializeWasm();

    return { slideCount: count };
  }

  /** Rewrite slide files in the files map with correct 1..N numbering. */
  private rewriteSlideFiles(contents: string[], rels: string[], count: number): void {
    for (const key of [...this.files.keys()]) {
      if (RE_SLIDE_FILE.test(key) || RE_SLIDE_RELS_FILE.test(key)) {
        this.files.delete(key);
      }
    }

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

    const newRelEntry = `<Relationship Id="${newRId}" Type="${REL_TYPE_SLIDE}" Target="slides/slide${newSlideNum}.xml"/>`;
    prsRels = prsRels.replace('</Relationships>', newRelEntry + '</Relationships>');

    // Update existing slide relationship targets (renumber slides after insertIdx)
    // First, collect existing slide rId → target mappings
    const slideRIdMap = this.parseSlideRelationships(prsRels);

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

    const newSldIdEntry = `<p:sldId id="${newSldId}" r:id="${newRId}"/>`;

    if (insertIdx === 0) {
      prsXml = prsXml.replace('<p:sldIdLst>', `<p:sldIdLst>${newSldIdEntry}`);
    } else {
      // Insert after the insertIdx-1 th sldId entry
      const sldIdEntries = this.parseSldIdEntries(prsXml);
      if (insertIdx <= sldIdEntries.length) {
        const afterEntry = sldIdEntries[insertIdx - 1];
        prsXml = prsXml.replace(afterEntry, afterEntry + newSldIdEntry);
      } else {
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

    const slideRIdMap = this.parseSlideRelationships(prsRels);
    let deleteRId = '';
    for (const [rId, target] of slideRIdMap) {
      if (target === `slides/slide${deleteNum}.xml`) {
        deleteRId = rId;
        break;
      }
    }

    if (deleteRId) {
      const sldIdEntries = this.parseSldIdEntries(prsXml);
      for (const entry of sldIdEntries) {
        if (entry.includes(`r:id="${deleteRId}"`)) {
          prsXml = prsXml.replace(entry, '');
          break;
        }
      }
    }

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

    const reordered = newOrder.map(i => sldIdEntries[i]);

    // Also need to update r:id → Target mappings for renumbered files
    // Each sldId's r:id now points to a different slide file number
    const slideRIdMap = this.parseSlideRelationships(prsRels);

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

  /** Add an image (PNG, JPEG, GIF, …) as a picture at an EMU position. Returns `"OK:<shapeIdx>"` or `"ERROR:..."`. */
  addImage(slideIdx: number, imageData: Uint8Array, mimeType: string,
    x: number, y: number, cx: number, cy: number): string {
    if (!this.wasm) return 'ERROR:not initialized';

    const ext = mimeToExt(mimeType);
    if (!ext) return 'ERROR:unsupported image type';

    this.checkpoint();

    const mediaPath = this.nextMediaPath(ext);

    // Store binary in rawFiles (for Wasm FFI) and addedBinaryFiles (for export)
    this.rawFiles.set(mediaPath, imageData);
    this.addedBinaryFiles.set(mediaPath, imageData);

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

    this.ensureContentTypeForExtension(ext, mimeType);

    // Add Picture shape via Wasm (add_picture_shape parses the slide on demand).
    return this.exports.add_picture_shape(slideIdx, rid, x, y, cx, cy);
  }

  /** Replace a picture's image bytes. Returns `"OK"` or `"ERROR:..."`. */
  replaceImage(slideIdx: number, shapeIdx: number,
    imageData: Uint8Array, mimeType: string): string {
    if (!this.wasm) return 'ERROR:not initialized';

    const ext = mimeToExt(mimeType);
    if (!ext) return 'ERROR:unsupported image type';

    const svg = this.exports.render_shape_svg(slideIdx, shapeIdx);
    if (svg.startsWith('ERROR:')) return svg;

    const ridMatch = svg.match(/data-ooxml-blip-rid="([^"]+)"/);
    if (!ridMatch) return 'ERROR:shape has no image reference';
    const currentRid = ridMatch[1];

    this.checkpoint();

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

  /** Delete a picture and, if nothing else uses it, its media file. Returns `"OK"` or `"ERROR:..."`. */
  deleteImage(slideIdx: number, shapeIdx: number): string {
    if (!this.wasm) return 'ERROR:not initialized';

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

    this.checkpoint();

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
        current_date: () => this.currentDate(),
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

  /** Date text for date placeholders, from the `currentDate` option or today. */
  private currentDate(): string {
    let v = this.currentDateOption;
    if (typeof v === 'function') v = v();
    if (v instanceof Date) return v.toLocaleDateString();
    if (typeof v === 'string') return v;
    return new Date().toLocaleDateString();
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

  /** A slide relationship's target, by the last segment of its type (`'notesSlide'`, …). */
  private resolveRelTarget(slideIdx: number, relType: string): string | null {
    const relsPath = `ppt/slides/_rels/slide${slideIdx + 1}.xml.rels`;
    const relsXml = this.files.get(relsPath);
    if (!relsXml) return null;

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

  /** Decode XML entities. The result is plain text: assign it to `.textContent`, never `.innerHTML`. */
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
