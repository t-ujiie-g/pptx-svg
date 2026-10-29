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

/** The `_` module as `WebAssembly.Global(externref)` values (Tier 3, Chrome 111+). */
function makeUnderscoreModule(constants: string[]): Record<string, WebAssembly.Global> {
  const mod: Record<string, WebAssembly.Global> = {};
  for (const s of constants) {
    mod[s] = new WebAssembly.Global({ value: 'externref', mutable: false }, s);
  }
  return mod;
}

/** The `_` module as plain strings (Tier 1b: Node.js 22+ supports builtins but won't resolve `_`). */
function makeUnderscoreModulePlain(constants: string[]): Record<string, string> {
  const mod: Record<string, string> = {};
  for (const s of constants) {
    mod[s] = s;
  }
  return mod;
}

/** Instantiate the module, trying each js-string tier in turn. */
export async function instantiateWasmWithFallback(
  bytes: ArrayBuffer,
  importObject: Record<string, Record<string, unknown>>,
  log?: { debug(...args: unknown[]): void; info(...args: unknown[]): void; warn(...args: unknown[]): void; error(...args: unknown[]): void },
): Promise<WebAssembly.WebAssemblyInstantiatedSource> {
  const noop = () => {};
  const l = log ?? { debug: noop, info: noop, warn: noop, error: noop };
  const stringConstants = parseWasmStringConstants(bytes);
  l.debug(`Parsed ${stringConstants.length} string constants from Wasm binary`);

  // Tier 1: modern builtins (browser with full js-string + importedStringConstants support)
  try {
    const r = await (WebAssembly as any).instantiate(bytes, importObject, { builtins: ['js-string'] });
    l.info('Wasm init: tier-1 (js-string builtins)');
    return r;
  } catch (e1: any) {
    l.info('Tier-1 failed:', e1.message, '— trying tier-1b');

    // Tier 1b (Node.js 22+): builtins, with `_` supplied as plain strings.
    const imports1b = { ...importObject, '_': makeUnderscoreModulePlain(stringConstants) };
    try {
      const r = await (WebAssembly as any).instantiate(bytes, imports1b, { builtins: ['js-string'] });
      l.info('Wasm init: tier-1b (js-string builtins + manual string constants)');
      return r;
    } catch (e1b: any) {
      l.info('Tier-1b failed:', e1b.message, '— trying tier-2');

      // Tier 2: importedStringConstants + manual wasm:js-string (Chrome 115-116)
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
          l.error('  Tier-1b:', e1b.message);
          l.error('  Tier-2:', e2.message);
          l.error('  Tier-3:', e3.message);
          throw new Error(
            `Wasm init failed — requires WebAssembly GC support (Chrome 111+, Node.js 22+). ` +
            `Tier-3 error: ${e3.message}`,
          );
        }
      }
    }
  }
}
