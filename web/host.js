/**
 * pptx-render host.js
 *
 * JavaScript host layer for the MoonBit Wasm module.
 *
 * Responsibilities:
 *   1. Parse the ZIP archive from the PPTX ArrayBuffer
 *   2. Decompress entries using the browser's DecompressionStream API
 *   3. Expose decompressed entry content to MoonBit via FFI imports
 *   4. Instantiate the Wasm module with js-string builtins compatibility
 *   5. Drive the MoonBit PPTX processing pipeline
 */

// ── Wasm js-string builtins compatibility ─────────────────────────────────────
//
// MoonBit's wasm-gc target with `use-js-builtin-string: true` generates a Wasm
// binary that imports:
//   • Functions from `wasm:js-string` (length, charCodeAt, equals, concat)
//   • String-constant globals from `_` (one per string literal in MoonBit code)
//
// Browser support timeline:
//   Chrome 111+: WebAssembly GC (wasm-gc)
//   Chrome 115+: importedStringConstants option ('_' module)
//   Chrome 117+: builtins: ['js-string'] (covers both wasm:js-string + '_')
//   Firefox 120+, Safari 17+: builtins: ['js-string']
//
// We try three tiers in order, falling back to wider compat:
//   Tier 1 — builtins: ['js-string']       (Chrome 117+, FF 120+, Safari 17+)
//   Tier 2 — importedStringConstants + manual wasm:js-string  (Chrome 115–116)
//   Tier 3 — fully manual '_' globals + wasm:js-string        (Chrome 111+)
//
// The '_' module string constants are parsed dynamically from the Wasm binary
// at startup, so no manual STRING_CONSTANTS list needs to be maintained.

/** Manual implementations of wasm:js-string builtin functions (Tier 2 & 3). */
const JS_STRING_MODULE = {
  length:     (s) => (s == null ? 0 : s.length) | 0,
  charCodeAt: (s, i) => s == null ? -1 : (s.charCodeAt(i) | 0),
  equals:     (a, b) => (a === b) ? 1 : 0,
  concat:     (a, b) => (a ?? '') + (b ?? ''),
};

/**
 * Parse the Wasm binary's import section to extract all '_' module
 * string-constant field names dynamically.
 *
 * This replaces the hand-maintained STRING_CONSTANTS array and eliminates
 * the need to run gen_string_constants.py after every source change.
 *
 * Format: each import with module="_" is a global whose field name IS the
 * string value (importedStringConstants convention from the js-string spec).
 *
 * @param {ArrayBuffer} buffer - Raw .wasm bytes
 * @returns {string[]}
 */
function parseWasmStringConstants(buffer) {
  const data = new Uint8Array(buffer);
  let pos = 8; // skip 4-byte magic + 4-byte version

  function readLEB128() {
    let result = 0, shift = 0;
    while (pos < data.length) {
      const b = data[pos++];
      result |= (b & 0x7f) << shift;
      shift += 7;
      if ((b & 0x80) === 0) break;
    }
    return result >>> 0;
  }

  function readUtf8() {
    const len = readLEB128();
    const chunk = data.subarray(pos, pos + len);
    pos += len;
    return new TextDecoder('utf-8').decode(chunk);
  }

  const constants = [];
  while (pos < data.length) {
    const sectionId = data[pos++];
    const sectionSize = readLEB128();
    const sectionEnd = pos + sectionSize;

    if (sectionId !== 2) { pos = sectionEnd; continue; } // not Import section

    const count = readLEB128();
    for (let i = 0; i < count; i++) {
      const mod   = readUtf8();
      const field = readUtf8();
      const kind  = data[pos++];
      if (kind === 0) {          // function: skip type index
        readLEB128();
      } else if (kind === 3) {   // global: skip valtype + possible heap-type + mutability
        const vt = data[pos++];
        if (vt === 0x64 || vt === 0x63) pos++; // ref non-null / null has heap-type byte
        pos++;                   // mutability byte
        if (mod === '_') constants.push(field);
      } else {
        break; // table/memory/tag — not expected in this binary
      }
    }
    break; // Import section fully parsed
  }
  return constants;
}

/**
 * Build the '_' import module as WebAssembly.Global(externref) objects.
 * Used in Tier 3 when the engine cannot resolve the module automatically.
 *
 * @param {string[]} constants - String values parsed from the Wasm binary
 */
function makeUnderscoreModule(constants) {
  const mod = {};
  for (const s of constants) {
    mod[s] = new WebAssembly.Global({ value: 'externref', mutable: false }, s);
  }
  return mod;
}

/**
 * Instantiate a Wasm module with three-tier js-string builtins fallback.
 *
 * @param {ArrayBuffer} bytes - Raw .wasm bytes
 * @param {object} importObject - Base import object (pptx_ffi, moonbit:ffi)
 * @returns {Promise<WebAssembly.WebAssemblyInstantiatedSource>}
 */
async function instantiateWasmWithFallback(bytes, importObject) {
  // Parse '_' module string constants from the binary (used in Tier 3).
  const stringConstants = parseWasmStringConstants(bytes);
  console.log(`[pptx] Parsed ${stringConstants.length} string constants from Wasm binary`);

  // Tier 1: modern builtins
  try {
    const r = await WebAssembly.instantiate(bytes, importObject, { builtins: ['js-string'] });
    console.log('[pptx] Wasm init: tier-1 (js-string builtins)');
    return r;
  } catch (e1) {
    console.warn('[pptx] Tier-1 failed:', e1.message, '— trying tier-2');

    // Tier 2: importedStringConstants + manual wasm:js-string
    const imports2 = { ...importObject, 'wasm:js-string': JS_STRING_MODULE };
    try {
      const r = await WebAssembly.instantiate(
        bytes, imports2, { importedStringConstants: '_' },
      );
      console.log('[pptx] Wasm init: tier-2 (importedStringConstants)');
      return r;
    } catch (e2) {
      console.warn('[pptx] Tier-2 failed:', e2.message, '— trying tier-3');

      // Tier 3: fully manual (any wasm-gc browser, Chrome 111+)
      const imports3 = {
        ...importObject,
        'wasm:js-string': JS_STRING_MODULE,
        '_': makeUnderscoreModule(stringConstants),
      };
      try {
        const r = await WebAssembly.instantiate(bytes, imports3);
        console.log('[pptx] Wasm init: tier-3 (full manual)');
        return r;
      } catch (e3) {
        console.error('[pptx] All instantiation tiers failed.');
        console.error('  Tier-1:', e1.message);
        console.error('  Tier-2:', e2.message);
        console.error('  Tier-3:', e3.message);
        throw new Error(
          `Wasm init failed — browser may not support WebAssembly GC (Chrome 111+). ` +
          `Tier-3 error: ${e3.message}`,
        );
      }
    }
  }
}

// ── PptxRenderer class ────────────────────────────────────────────────────────

export class PptxRenderer {
  /** @type {WebAssembly.Instance|null} */
  #wasm = null;

  /** @type {Map<string, string>} Decompressed text ZIP entries (path → UTF-8 string) */
  #files = new Map();

  /** @type {Map<string, Uint8Array>} Raw binary ZIP entries (path → bytes) */
  #rawFiles = new Map();

  /**
   * Initialize the renderer by loading the Wasm module.
   * The drop zone should be disabled until this resolves.
   *
   * @param {string} wasmUrl - URL to the .wasm file
   */
  async init(wasmUrl) {
    const response = await fetch(wasmUrl);
    if (!response.ok) throw new Error(`HTTP ${response.status} fetching ${wasmUrl}`);
    const bytes = await response.arrayBuffer();

    const result = await instantiateWasmWithFallback(bytes, this.#buildImportObject());
    this.#wasm = result.instance;
    console.log('[pptx] Wasm module loaded. Exports:', Object.keys(this.#wasm.exports));
  }

  /**
   * Load a PPTX file from an ArrayBuffer.
   * Parses the ZIP archive and decompresses all entries,
   * then calls MoonBit's initialize_pptx() to count slides.
   *
   * @param {ArrayBuffer} arrayBuffer - Raw PPTX file bytes
   * @returns {Promise<{slideCount: number}>}
   */
  async loadPptx(arrayBuffer) {
    if (!this.#wasm) {
      throw new Error('Wasm not initialized — wait for init() to complete before loading files.');
    }
    console.log('[pptx] Parsing ZIP archive...');
    const { textFiles, binaryFiles } = await this.#extractZip(arrayBuffer);
    this.#files = textFiles;
    this.#rawFiles = binaryFiles;
    console.log(`[pptx] Extracted ${textFiles.size} text entries, ${binaryFiles.size} binary entries`);

    const result = this.#wasm.exports.initialize_pptx();
    console.log('[pptx] initialize_pptx result:', result);

    if (result.startsWith('ERROR:')) throw new Error(result.slice(6));

    const slideCount = this.#wasm.exports.get_slide_count();
    return { slideCount };
  }

  /** @returns {number} Number of slides in the loaded presentation. */
  getSlideCount() {
    return this.#wasm.exports.get_slide_count();
  }

  /**
   * Get the raw XML of a slide (0-indexed). For debugging.
   * @param {number} slideIdx
   * @returns {string}
   */
  getSlideXmlRaw(slideIdx) {
    return this.#wasm.exports.get_slide_xml_raw(slideIdx);
  }

  /**
   * Get all entry paths in the PPTX archive. For debugging.
   * @returns {string[]}
   */
  getEntryList() {
    return this.#wasm.exports.get_entry_list().split('\n').filter(Boolean);
  }

  /**
   * Render a slide as an SVG string (0-indexed).
   * @param {number} slideIdx
   * @returns {string} SVG markup, or a string starting with "ERROR:" on failure
   */
  renderSlideSvg(slideIdx) {
    return this.#wasm.exports.render_slide_svg(slideIdx);
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /**
   * Build the Wasm import object that satisfies MoonBit's FFI declarations.
   * Arrow functions capture `this` so they can access #files and #rawFiles.
   */
  #buildImportObject() {
    return {
      'pptx_ffi': {
        get_file:         (path) => this.#files.get(path) ?? '',
        get_entry_list:   () => [...this.#files.keys(), ...this.#rawFiles.keys()].join('\n'),
        get_file_base64:  (path) => bytesToBase64(this.#rawFiles.get(path)),
        char_code_to_str: (n) => String.fromCharCode(n),
        log:   (msg) => console.log('[pptx]', msg),
        warn:  (msg) => console.warn('[pptx]', msg),
        error: (msg) => console.error('[pptx]', msg),
        measure_text: (text, fontFace, fontSizePt) => this.#measureText(text, fontFace, fontSizePt),
      },
      // make_closure is used by some MoonBit wasm-gc closure patterns.
      'moonbit:ffi': {
        make_closure: (f, ctx) => f.bind(null, ctx),
      },
    };
  }

  // ── ZIP extraction ───────────────────────────────────────────────────────────

  /**
   * Extract all entries from a ZIP archive.
   *
   * Text entries (XML, rels, …) → textFiles Map (path → UTF-8 string)
   * Binary entries (images, …)  → binaryFiles Map (path → Uint8Array)
   *
   * Handles method 0 (stored) and method 8 (DEFLATE).
   *
   * @param {ArrayBuffer} buffer
   * @returns {Promise<{textFiles: Map<string,string>, binaryFiles: Map<string,Uint8Array>}>}
   */
  async #extractZip(buffer) {
    const bytes = new Uint8Array(buffer);
    const view = new DataView(buffer);
    const textFiles = new Map();
    const binaryFiles = new Map();
    const decoder = new TextDecoder('utf-8');

    let offset = 0;
    while (offset < bytes.length - 4) {
      // Local File Header signature: PK\x03\x04
      if (view.getUint32(offset, true) !== 0x04034b50) break;

      const method          = view.getUint16(offset + 8,  true);
      const compressedSize  = view.getUint32(offset + 18, true);
      const uncompressedSize= view.getUint32(offset + 22, true);
      const fileNameLen     = view.getUint16(offset + 26, true);
      const extraLen        = view.getUint16(offset + 28, true);

      const name       = decoder.decode(bytes.slice(offset + 30, offset + 30 + fileNameLen));
      const dataOffset = offset + 30 + fileNameLen + extraLen;
      const compressed = bytes.slice(dataOffset, dataOffset + compressedSize);

      let decompressed;
      if (method === 0) {
        decompressed = compressed;
      } else if (method === 8) {
        decompressed = await this.#inflate(compressed, uncompressedSize);
      } else {
        console.warn(`[pptx] Unsupported compression method ${method} for ${name}, skipping`);
        offset = dataOffset + compressedSize;
        continue;
      }

      if (this.#isTextEntry(name)) {
        textFiles.set(name, decoder.decode(decompressed));
      } else {
        binaryFiles.set(name, decompressed);
      }

      offset = dataOffset + compressedSize;
    }

    return { textFiles, binaryFiles };
  }

  /**
   * Decompress raw DEFLATE bytes using the browser's native DecompressionStream.
   * @param {Uint8Array} compressed
   * @param {number} _hint - expected size (unused; browser handles allocation)
   * @returns {Promise<Uint8Array>}
   */
  async #inflate(compressed, _hint) {
    const stream = new DecompressionStream('deflate-raw');
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();

    writer.write(compressed);
    writer.close();

    const chunks = [];
    let totalLen = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      totalLen += value.length;
    }

    const result = new Uint8Array(totalLen);
    let pos = 0;
    for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
    return result;
  }

  /**
   * Return true if a ZIP entry should be decoded as UTF-8 text.
   * @param {string} name - entry path
   */
  #isTextEntry(name) {
    const lower = name.toLowerCase();
    return lower.endsWith('.xml')  ||
           lower.endsWith('.rels') ||
           lower.endsWith('.txt')  ||
           lower.endsWith('.json') ||
           lower.endsWith('.html') ||
           lower.endsWith('.css')  ||
           lower === '[content_types].xml';
  }

  // ── Text measurement (Phase 2+) ──────────────────────────────────────────────

  /** @type {HTMLCanvasElement|null} */
  #canvas = null;
  /** @type {CanvasRenderingContext2D|null} */
  #ctx = null;

  /**
   * Measure the rendered pixel width of text using Canvas 2D.
   * Lazily creates the canvas on first call.
   *
   * @param {string} text
   * @param {string} fontFace
   * @param {number} fontSizePt
   * @returns {number} width in pixels
   */
  #measureText(text, fontFace, fontSizePt) {
    if (!this.#ctx) {
      this.#canvas = document.createElement('canvas');
      this.#ctx = this.#canvas.getContext('2d');
    }
    this.#ctx.font = `${fontSizePt}pt ${fontFace}`;
    return this.#ctx.measureText(text).width;
  }
}

// ── Utilities ─────────────────────────────────────────────────────────────────

/**
 * Encode a Uint8Array as a base64 string.
 * @param {Uint8Array|undefined} bytes
 * @returns {string}
 */
function bytesToBase64(bytes) {
  if (!bytes) return '';
  let binary = '';
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}
