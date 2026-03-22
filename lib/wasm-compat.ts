/**
 * Wasm js-string builtins compatibility layer.
 *
 * Provides 3-tier fallback for instantiating MoonBit Wasm modules
 * that use `use-js-builtin-string: true`:
 *   - Tier 1: Native `{ builtins: ['js-string'] }` (Chrome 117+, FF 120+, Safari 17+)
 *   - Tier 2: `{ importedStringConstants: '_' }` + manual wasm:js-string (Chrome 115-116)
 *   - Tier 3: Fully manual Global(externref) for '_' module (Chrome 111+)
 */

/** Manual implementations of wasm:js-string builtin functions. */
const JS_STRING_MODULE: Record<string, Function> = {
  length:     (s: string | null) => (s == null ? 0 : s.length) | 0,
  charCodeAt: (s: string | null, i: number) => s == null ? -1 : (s.charCodeAt(i) | 0),
  equals:     (a: string | null, b: string | null) => (a === b) ? 1 : 0,
  concat:     (a: string | null, b: string | null) => (a ?? '') + (b ?? ''),
};

/**
 * Parse the Wasm binary's import section to extract all '_' module
 * string-constant field names dynamically.
 */
export function parseWasmStringConstants(buffer: ArrayBuffer): string[] {
  const data = new Uint8Array(buffer);
  let pos = 8; // skip 4-byte magic + 4-byte version

  function readLEB128(): number {
    let result = 0, shift = 0;
    while (pos < data.length) {
      const b = data[pos++];
      result |= (b & 0x7f) << shift;
      shift += 7;
      if ((b & 0x80) === 0) break;
    }
    return result >>> 0;
  }

  function readUtf8(): string {
    const len = readLEB128();
    const chunk = data.subarray(pos, pos + len);
    pos += len;
    return new TextDecoder('utf-8').decode(chunk);
  }

  const constants: string[] = [];
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
 */
function makeUnderscoreModule(constants: string[]): Record<string, WebAssembly.Global> {
  const mod: Record<string, WebAssembly.Global> = {};
  for (const s of constants) {
    mod[s] = new WebAssembly.Global({ value: 'externref', mutable: false }, s);
  }
  return mod;
}

/**
 * Instantiate a Wasm module with three-tier js-string builtins fallback.
 *
 * @param bytes - Raw .wasm bytes
 * @param importObject - Base import object (pptx_ffi, moonbit:ffi)
 * @returns Instantiated Wasm source
 */
export async function instantiateWasmWithFallback(
  bytes: ArrayBuffer,
  importObject: Record<string, Record<string, unknown>>,
  log?: { debug(...args: unknown[]): void; info(...args: unknown[]): void; warn(...args: unknown[]): void; error(...args: unknown[]): void },
): Promise<WebAssembly.WebAssemblyInstantiatedSource> {
  const noop = () => {};
  const l = log ?? { debug: noop, info: noop, warn: noop, error: noop };
  const stringConstants = parseWasmStringConstants(bytes);
  l.debug(`Parsed ${stringConstants.length} string constants from Wasm binary`);

  // Tier 1: modern builtins
  try {
    const r = await (WebAssembly as any).instantiate(bytes, importObject, { builtins: ['js-string'] });
    l.info('Wasm init: tier-1 (js-string builtins)');
    return r;
  } catch (e1: any) {
    l.info('Tier-1 failed:', e1.message, '— trying tier-2');

    // Tier 2: importedStringConstants + manual wasm:js-string
    const imports2 = { ...importObject, 'wasm:js-string': JS_STRING_MODULE };
    try {
      const r = await (WebAssembly as any).instantiate(
        bytes, imports2, { importedStringConstants: '_' },
      );
      l.info('Wasm init: tier-2 (importedStringConstants)');
      return r;
    } catch (e2: any) {
      l.info('Tier-2 failed:', e2.message, '— trying tier-3');

      // Tier 3: fully manual (any wasm-gc browser, Chrome 111+)
      const imports3 = {
        ...importObject,
        'wasm:js-string': JS_STRING_MODULE,
        '_': makeUnderscoreModule(stringConstants),
      };
      try {
        const r = await WebAssembly.instantiate(bytes, imports3);
        l.info('Wasm init: tier-3 (full manual)');
        return r;
      } catch (e3: any) {
        l.error('All instantiation tiers failed.');
        l.error('  Tier-1:', e1.message);
        l.error('  Tier-2:', e2.message);
        l.error('  Tier-3:', e3.message);
        throw new Error(
          `Wasm init failed — browser may not support WebAssembly GC (Chrome 111+). ` +
          `Tier-3 error: ${e3.message}`,
        );
      }
    }
  }
}
