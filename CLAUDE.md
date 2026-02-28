# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Build Wasm (output: _build/wasm-gc/release/build/main/main.wasm, ~35KB)
moon build --target wasm-gc --release

# Build TypeScript library (output: dist/)
tsc

# Build everything (Wasm + TypeScript)
npm run build

# Run JS-layer tests (no browser needed, tests ZIP extraction + slide counting)
node test_fixtures/test_node.mjs

# Serve for browser testing
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html
```

## Architecture

**Separation of concerns:**
- **TypeScript library** (`lib/` → `dist/`): ZIP parsing/building, DEFLATE, Wasm lifecycle, CRC-32, `PptxRenderer` class
- **MoonBit** (`src/`): OOXML parsing, SVG generation, SVG→SlideData parsing, OOXML serialization
- **Demo** (`web/`): Browser demo UI (imports from `dist/`)

**FFI boundary:**
- JS pre-decompresses all ZIP entries → stores in `Map<path, string>` and `Map<path, Uint8Array>`
- MoonBit calls `ffi_get_file(path)` to pull individual files on demand
- MoonBit exports: `initialize_pptx`, `get_slide_count`, `get_slide_xml_raw`, `get_entry_list`, `render_slide_svg`, `update_slide_from_svg`, `get_slide_ooxml`, `get_modified_entries`

**Module dependency (no cycles):**
```
main → renderer   → ooxml → xml
     → svg_parser → ooxml → xml
     → serializer → ooxml
     → ffi
```

## Critical MoonBit constraints

**No integer string interpolation.** `"\{n}"` for integer `n` calls `fromCharCodeArray` internally, which requires `{ builtins: ['js-string'] }` browser support (Chrome 117+). The codebase uses `int_to_str(n)` helper instead, which only uses `concat` + string literals and works in all wasm-gc browsers (Chrome 111+).

**String API:** Use `s.get_char(i).unwrap()` (not deprecated `unsafe_char_at`). Avoid `s[i:j]` in non-error functions — it raises `CreatingViewError`.

**No external packages.** `bobzhang/zip` and `ruifeng/XMLParser` are incompatible with the current compiler (Feb 2026). Do not add external deps; implement needed parsers inline.

**pub(all) for cross-package construction.** Structs and enums in `ooxml` that need to be constructed from other packages (svg_parser, serializer, main) use `pub(all)` visibility. `pub struct` fields are read-only from other packages.

## Browser compatibility and string constants

`use-js-builtin-string: true` in `src/main/moon.pkg.json` generates Wasm that imports:
1. Functions from `wasm:js-string` (length, charCodeAt, equals, concat)
2. String-constant globals from module `_` (one per string literal in MoonBit)

`lib/wasm-compat.ts` handles this with a 3-tier fallback:
- **Tier 1** `{ builtins: ['js-string'] }` — Chrome 117+, Firefox 120+, Safari 17+
- **Tier 2** `{ importedStringConstants: '_' }` + manual `wasm:js-string` — Chrome 115–116
- **Tier 3** Manual `WebAssembly.Global(externref)` for `_` + manual `wasm:js-string` — Chrome 111+

`wasm-compat.ts` parses the Wasm binary at startup to extract `_` module string constants dynamically — no manual list to maintain.

**Critical**: Never use `StringBuilder` in MoonBit. `StringBuilder::to_string()` calls `wasm:js-string "fromCharCodeArray"` which cannot be polyfilled in JS. Build strings with `+` (concat) instead. For Char→String use `@ffi.ffi_char_code_to_str(Char::to_int(c))` (→ `String.fromCharCode`).

## Data model (ooxml.mbt)

```
SlideData { slide_size: SlideSize, background: Color, shapes: Array[Shape] }
Shape { kind: ShapeKind, transform: ShapeTransform, fill, stroke, stroke_w, paragraphs }

ShapeKind = AutoShape(ShapeGeom) | Picture(String) | TableShape(TableData) | GroupShape | Other
ShapeGeom = Rect | Ellipse | RoundRect | Line | Other(String)
ShapeTransform { x, y, cx, cy, rot, flip_h, flip_v }  // all EMU

TextParagraph { runs: Array[TextRun], align: String, level: Int }
TextRun { text, bold, italic, font_size, color, font_face }

TableData { col_widths: Array[Int], rows: Array[TableRow] }
TableRow { height: Int, cells: Array[TableCell] }
TableCell { paragraphs: Array[TextParagraph], fill: Color }

Color { r, g, b }  // -1 = none (sentinel)
ThemeData { dk1..fol_hlink: Color, major_font, minor_font: String }
```

## Key files

| File | Purpose |
|------|---------|
| `src/ffi/ffi.mbt` | All JS→Wasm import declarations |
| `src/xml/xml.mbt` | Generic XML parser (DOM tree) |
| `src/ooxml/ooxml.mbt` | OOXML types (`SlideData`, `Shape`, etc.) + PPTX slide XML parser |
| `src/renderer/renderer.mbt` | SlideData → SVG with `data-ooxml-*` attributes |
| `src/svg_parser/svg_parser.mbt` | SVG (with `data-ooxml-*`) → SlideData |
| `src/serializer/serializer.mbt` | SlideData → OOXML slide XML |
| `src/main/main.mbt` | Wasm exports, slide cache (`g_slides`), round-trip orchestration |
| `src/main/moon.pkg.json` | Export list + `use-js-builtin-string: true` |
| `lib/index.ts` | Library public API re-exports |
| `lib/pptx-renderer.ts` | `PptxRenderer` class (core API) |
| `lib/wasm-compat.ts` | 3-tier Wasm js-string builtins fallback |
| `lib/zip.ts` | ZIP extraction and building |
| `lib/utils.ts` | bytesToBase64, crc32 utilities |
| `web/host.js` | Legacy JS host (kept for reference; demo uses `dist/`) |
| `web/index.html` | Browser demo UI |
| `test_fixtures/minimal.pptx` | 2-slide test fixture |
