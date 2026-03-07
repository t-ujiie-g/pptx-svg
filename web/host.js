/**
 * pptx-svg host.js
 *
 * JavaScript host layer for the MoonBit Wasm module.
 *
 * Responsibilities:
 *   1. Parse the ZIP archive from the PPTX ArrayBuffer
 *   2. Decompress entries using the browser's DecompressionStream API
 *   3. Expose decompressed entry content to MoonBit via FFI imports
 *   4. Instantiate the Wasm module with js-string builtins compatibility
 *   5. Drive the MoonBit PPTX processing pipeline
 *   6. Export modified PPTX via round-trip pipeline
 */

// ── Wasm js-string builtins compatibility ─────────────────────────────────────

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

  /** @type {ArrayBuffer|null} Original PPTX bytes for export */
  #originalBuffer = null;

  /**
   * Initialize the renderer by loading the Wasm module.
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
   * @param {ArrayBuffer} arrayBuffer - Raw PPTX file bytes
   * @returns {Promise<{slideCount: number}>}
   */
  async loadPptx(arrayBuffer) {
    if (!this.#wasm) {
      throw new Error('Wasm not initialized — wait for init() to complete before loading files.');
    }
    this.#originalBuffer = arrayBuffer.slice(0); // keep a copy for export
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

  /**
   * Update a slide's internal data from an edited SVG string.
   * Parses the SVG's data-ooxml-* attributes back into SlideData.
   * @param {number} slideIdx
   * @param {string} svgString
   * @returns {string} "OK" on success, "ERROR:..." on failure
   */
  updateSlideFromSvg(slideIdx, svgString) {
    return this.#wasm.exports.update_slide_from_svg(slideIdx, svgString);
  }

  /**
   * Get the OOXML slide XML for a slide (0-indexed).
   * Returns modified XML if the slide was updated, otherwise original.
   * @param {number} slideIdx
   * @returns {string}
   */
  getSlideOoxml(slideIdx) {
    return this.#wasm.exports.get_slide_ooxml(slideIdx);
  }

  /**
   * Export the (possibly modified) presentation as a PPTX ArrayBuffer.
   * Replaces modified slide XML entries in the original ZIP and rebuilds it.
   * @returns {Promise<ArrayBuffer>}
   */
  async exportPptx() {
    if (!this.#originalBuffer) {
      throw new Error('No PPTX loaded — call loadPptx() first.');
    }

    // Get modified entries from Wasm: "path\tcontent\n..."
    const modifiedStr = this.#wasm.exports.get_modified_entries();
    const modifications = new Map();
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

    // Rebuild ZIP with modifications
    return this.#buildZip(this.#originalBuffer, modifications);
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /**
   * Build the Wasm import object that satisfies MoonBit's FFI declarations.
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
        math_sin:   (x) => Math.sin(x),
        math_cos:   (x) => Math.cos(x),
        math_atan2: (y, x) => Math.atan2(y, x),
        math_sqrt:  (x) => Math.sqrt(x),
      },
      'moonbit:ffi': {
        make_closure: (f, ctx) => f.bind(null, ctx),
      },
    };
  }

  // ── ZIP extraction ───────────────────────────────────────────────────────────

  /**
   * Extract all entries from a ZIP archive.
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
   * @param {number} _hint
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
   * Compress bytes using DEFLATE-raw via CompressionStream.
   * @param {Uint8Array} data
   * @returns {Promise<Uint8Array>}
   */
  async #deflate(data) {
    const stream = new CompressionStream('deflate-raw');
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();

    writer.write(data);
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
   * @param {string} name
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

  // ── ZIP building ────────────────────────────────────────────────────────────

  /**
   * Build a new ZIP by iterating the original ZIP entries and replacing
   * modified text entries with new content.
   *
   * @param {ArrayBuffer} originalBuffer - The original PPTX bytes
   * @param {Map<string, string>} modifications - path → new XML content
   * @returns {Promise<ArrayBuffer>}
   */
  async #buildZip(originalBuffer, modifications) {
    const origBytes = new Uint8Array(originalBuffer);
    const origView = new DataView(originalBuffer);
    const decoder = new TextDecoder('utf-8');
    const encoder = new TextEncoder();

    // Collect all local file entries
    const entries = [];
    let offset = 0;
    while (offset < origBytes.length - 4) {
      if (origView.getUint32(offset, true) !== 0x04034b50) break;

      const method           = origView.getUint16(offset + 8, true);
      const compressedSize   = origView.getUint32(offset + 18, true);
      const fileNameLen      = origView.getUint16(offset + 26, true);
      const extraLen         = origView.getUint16(offset + 28, true);
      const name             = decoder.decode(origBytes.slice(offset + 30, offset + 30 + fileNameLen));
      const dataOffset       = offset + 30 + fileNameLen + extraLen;

      // Copy the flags, time, date from the original header
      const flags   = origView.getUint16(offset + 6, true);
      const time    = origView.getUint16(offset + 10, true);
      const date    = origView.getUint16(offset + 12, true);
      const crc32   = origView.getUint32(offset + 14, true);
      const uncompressedSize = origView.getUint32(offset + 22, true);

      const compressedData = origBytes.slice(dataOffset, dataOffset + compressedSize);

      entries.push({
        name, method, flags, time, date, crc32,
        compressedSize, uncompressedSize,
        compressedData, extra: origBytes.slice(offset + 30 + fileNameLen, dataOffset),
      });

      offset = dataOffset + compressedSize;
    }

    // Process modifications: replace entry data
    for (const entry of entries) {
      if (modifications.has(entry.name)) {
        const newContent = encoder.encode(modifications.get(entry.name));
        const compressed = await this.#deflate(newContent);
        entry.method = 8;
        entry.compressedData = compressed;
        entry.compressedSize = compressed.length;
        entry.uncompressedSize = newContent.length;
        entry.crc32 = crc32(newContent);
        entry.extra = new Uint8Array(0);
      }
    }

    // Build the new ZIP
    const parts = [];
    const centralDir = [];
    let localOffset = 0;

    for (const entry of entries) {
      const nameBytes = encoder.encode(entry.name);

      // Local file header (30 bytes + name + extra + data)
      const localHeader = new ArrayBuffer(30);
      const lhView = new DataView(localHeader);
      lhView.setUint32(0, 0x04034b50, true);   // signature
      lhView.setUint16(4, 20, true);            // version needed
      lhView.setUint16(6, entry.flags, true);   // flags
      lhView.setUint16(8, entry.method, true);  // method
      lhView.setUint16(10, entry.time, true);   // time
      lhView.setUint16(12, entry.date, true);   // date
      lhView.setUint32(14, entry.crc32, true);  // crc32
      lhView.setUint32(18, entry.compressedSize, true);
      lhView.setUint32(22, entry.uncompressedSize, true);
      lhView.setUint16(26, nameBytes.length, true);
      lhView.setUint16(28, entry.extra.length, true);

      parts.push(new Uint8Array(localHeader));
      parts.push(nameBytes);
      parts.push(entry.extra);
      parts.push(entry.compressedData);

      // Central directory entry (46 bytes + name)
      const cdHeader = new ArrayBuffer(46);
      const cdView = new DataView(cdHeader);
      cdView.setUint32(0, 0x02014b50, true);    // signature
      cdView.setUint16(4, 20, true);             // version made by
      cdView.setUint16(6, 20, true);             // version needed
      cdView.setUint16(8, entry.flags, true);
      cdView.setUint16(10, entry.method, true);
      cdView.setUint16(12, entry.time, true);
      cdView.setUint16(14, entry.date, true);
      cdView.setUint32(16, entry.crc32, true);
      cdView.setUint32(20, entry.compressedSize, true);
      cdView.setUint32(24, entry.uncompressedSize, true);
      cdView.setUint16(28, nameBytes.length, true);
      cdView.setUint16(30, 0, true);             // extra length
      cdView.setUint16(32, 0, true);             // comment length
      cdView.setUint16(34, 0, true);             // disk number
      cdView.setUint16(36, 0, true);             // internal attrs
      cdView.setUint32(38, 0, true);             // external attrs
      cdView.setUint32(42, localOffset, true);   // local header offset

      centralDir.push(new Uint8Array(cdHeader));
      centralDir.push(nameBytes);

      localOffset += 30 + nameBytes.length + entry.extra.length + entry.compressedData.length;
    }

    const cdOffset = localOffset;
    let cdSize = 0;
    for (const part of centralDir) cdSize += part.length;

    // End of central directory (22 bytes)
    const eocd = new ArrayBuffer(22);
    const eocdView = new DataView(eocd);
    eocdView.setUint32(0, 0x06054b50, true);    // signature
    eocdView.setUint16(4, 0, true);              // disk number
    eocdView.setUint16(6, 0, true);              // cd disk number
    eocdView.setUint16(8, entries.length, true);  // entries on disk
    eocdView.setUint16(10, entries.length, true); // total entries
    eocdView.setUint32(12, cdSize, true);         // cd size
    eocdView.setUint32(16, cdOffset, true);       // cd offset
    eocdView.setUint16(20, 0, true);              // comment length

    // Combine all parts
    let totalSize = 0;
    for (const p of parts) totalSize += p.length;
    totalSize += cdSize + 22;

    const result = new Uint8Array(totalSize);
    let pos = 0;
    for (const p of parts) { result.set(p, pos); pos += p.length; }
    for (const p of centralDir) { result.set(p, pos); pos += p.length; }
    result.set(new Uint8Array(eocd), pos);

    return result.buffer;
  }

  // ── Text measurement ──────────────────────────────────────────────────────────

  /** @type {HTMLCanvasElement|null} */
  #canvas = null;
  /** @type {CanvasRenderingContext2D|null} */
  #ctx = null;

  /**
   * Measure the rendered pixel width of text using Canvas 2D.
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

/**
 * Compute CRC-32 for a Uint8Array.
 * @param {Uint8Array} data
 * @returns {number}
 */
function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
