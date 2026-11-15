/**
 * Instantiation fallbacks for `use-js-builtin-string` Wasm across browsers and Node.js;
 * the tiers are described in CLAUDE.md.
 */

/** Manual implementations of wasm:js-string builtin functions. */
const JS_STRING_MODULE: Record<string, Function> = {
  length:     (s: string | null) => (s == null ? 0 : s.length) | 0,
  charCodeAt: (s: string | null, i: number) => s == null ? -1 : (s.charCodeAt(i) | 0),
  equals:     (a: string | null, b: string | null) => (a === b) ? 1 : 0,
  concat:     (a: string | null, b: string | null) => (a ?? '') + (b ?? ''),
  // Its Wasm GC array argument is opaque to JS, so it can't be implemented. The stdlib only
  // calls it to format a panic message just before trapping, but the import must still exist.
  fromCharCodeArray: () => '',
};

/** The `_` module string-constant names, read from the Wasm import section. */
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
 * The `_` module as plain strings. The constants are imported as non-null `(ref extern)`
 * globals, which a nullable `WebAssembly.Global({ value: 'externref' })` does not match.
 */
function makeUnderscoreModule(constants: string[]): Record<string, string> {
  const mod: Record<string, string> = {};
  for (const s of constants) {
    mod[s] = s;
  }
  return mod;
}

/**
 * Instantiate the module with the engine's js-string builtins (recent Chrome/Firefox, Safari 26.2+,
 * Node.js 22+) or, failing that, with them implemented in JS — any WasmGC engine (Chrome 119+,
 * Firefox 120+, Safari 18.2+).
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
  const imports = { ...importObject, '_': makeUnderscoreModule(stringConstants) };

  try {
    const r = await (WebAssembly as any).instantiate(bytes, imports, { builtins: ['js-string'] });
    l.info('Wasm init: js-string builtins');
    return r;
  } catch (e1: any) {
    l.info('js-string builtins unavailable:', e1.message, '— using the JS implementation');
    try {
      const r = await WebAssembly.instantiate(
        bytes, { ...imports, 'wasm:js-string': JS_STRING_MODULE } as unknown as WebAssembly.Imports,
      );
      l.info('Wasm init: js-string in JS');
      return r;
    } catch (e2: any) {
      l.error('Wasm instantiation failed.');
      l.error('  with js-string builtins:', e1.message);
      l.error('  with js-string in JS:', e2.message);
      throw new Error(
        `Wasm init failed — requires WebAssembly GC support (Chrome 119+, Firefox 120+, ` +
        `Safari 18.2+, Node.js 22+). Error: ${e2.message}`,
      );
    }
  }
}
