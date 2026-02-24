# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Build Wasm (output: _build/wasm-gc/release/build/main/main.wasm, ~1.6KB)
moon build --target wasm-gc --release

# Run JS-layer tests (no browser needed, tests ZIP extraction + slide counting)
node test_fixtures/test_node.mjs

# Serve for browser testing
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html

# After changing MoonBit string literals — regenerate STRING_CONSTANTS in host.js
moon build --target wasm-gc --release
python3 scripts/gen_string_constants.py --update
```

## Architecture

**Separation of concerns:**
- **JavaScript** (`web/host.js`): ZIP parsing, DEFLATE decompression via `DecompressionStream`, Wasm lifecycle
- **MoonBit** (`src/`): All OOXML/PPTX logic, SVG generation (Phase 2+), edit operations (Phase 5+)

**FFI boundary:**
- JS pre-decompresses all ZIP entries → stores in `Map<path, string>` and `Map<path, Uint8Array>`
- MoonBit calls `ffi_get_file(path)` to pull individual files on demand
- MoonBit exports `initialize_pptx`, `get_slide_count`, `get_slide_xml_raw`, `get_entry_list`; Phase 2 adds `render_slide_svg(idx)`

**Module dependency (no cycles):**
```
main → renderer → ffi
     → editor
     → ooxml → xml
```

## Critical MoonBit constraints

**No integer string interpolation.** `"\{n}"` for integer `n` calls `fromCharCodeArray` internally, which requires `{ builtins: ['js-string'] }` browser support (Chrome 117+). The codebase uses `int_to_str(n)` helper in `main.mbt` instead, which only uses `concat` + string literals and works in all wasm-gc browsers (Chrome 111+).

**String API:** Use `s.get_char(i).unwrap()` (not deprecated `unsafe_char_at`). Avoid `s[i:j]` in non-error functions — it raises `CreatingViewError`.

**No external packages.** `bobzhang/zip` and `ruifeng/XMLParser` are incompatible with the current compiler (Feb 2026). Do not add external deps; implement needed parsers inline.

## Browser compatibility and STRING_CONSTANTS

`use-js-builtin-string: true` in `src/main/moon.pkg.json` generates Wasm that imports:
1. Functions from `wasm:js-string` (length, charCodeAt, equals, concat)
2. String-constant globals from module `_` (one per string literal in MoonBit)

`web/host.js` handles this with a 3-tier fallback:
- **Tier 1** `{ builtins: ['js-string'] }` — Chrome 117+, Firefox 120+, Safari 17+
- **Tier 2** `{ importedStringConstants: '_' }` + manual `wasm:js-string` — Chrome 115–116
- **Tier 3** Manual `WebAssembly.Global(externref)` for `_` + manual `wasm:js-string` — Chrome 111+

The `STRING_CONSTANTS` array in `host.js` must list every string literal in the MoonBit binary. After any MoonBit source change that adds or removes string literals, run `python3 scripts/gen_string_constants.py --update` to regenerate it.

## Key files

| File | Purpose |
|------|---------|
| `src/ffi/ffi.mbt` | All JS→Wasm import declarations |
| `src/main/main.mbt` | Exported Wasm functions + slide parsing; also contains `int_to_str` helper |
| `src/main/moon.pkg.json` | Export list + `use-js-builtin-string: true` |
| `web/host.js` | ZIP extractor, 3-tier Wasm instantiation, `PptxRenderer` class |
| `scripts/gen_string_constants.py` | Parses Wasm binary import section to regenerate `STRING_CONSTANTS` |
| `test_fixtures/minimal.pptx` | 2-slide test fixture (created with Python) |
