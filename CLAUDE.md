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
main → renderer   → xml, ooxml, ffi
     → svg_parser → xml, ooxml, ffi
     → serializer → xml, ooxml, ffi
     → ffi
xml (shared: int_to_str, parse_int, XML parser)
ooxml → xml (types, PPTX parser, parse_hex_color)
```

## Critical MoonBit constraints

**No integer string interpolation.** `"\{n}"` for integer `n` calls `fromCharCodeArray` internally, which requires `{ builtins: ['js-string'] }` browser support (Chrome 117+). The codebase uses `@xml.int_to_str(n)` (defined in `xml.mbt`, aliased locally as `fn int_to_str(n) -> String { @xml.int_to_str(n) }`) which only uses `concat` + string literals and works in all wasm-gc browsers (Chrome 111+).

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
SlideData { slide_size: SlideSize, background: Color, bg_grad: GradientFill, shapes: Array[Shape] }
Shape { kind: ShapeKind, transform: ShapeTransform,
  fill: Color, grad_fill: GradientFill, blip_fill: BlipFill, patt_fill: PatternFill,
  stroke: Color, stroke_w: Int,
  stroke_dash: String, stroke_cap: String, stroke_join: String, stroke_miter_limit: Int,
  stroke_head_type: String, stroke_head_w: String, stroke_head_len: String,
  stroke_tail_type: String, stroke_tail_w: String, stroke_tail_len: String,
  stroke_cmpd: String, stroke_no_fill: Bool,
  paragraphs: Array[TextParagraph], body_props: BodyProps, ph_type: String, ph_idx: Int }

ShapeKind = AutoShape(ShapeGeom) | Picture(String) | TableShape(TableData) | GroupShape(GroupShapeData) | Other
ShapeGeom = Rect | Ellipse | RoundRect | Line | Connector(String, Array[Int]) | Other(String, Array[Int])
GroupShapeData { ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy: Int, children: Array[Shape] }
ShapeTransform { x, y, cx, cy, rot, flip_h, flip_v }  // all EMU

GradientStop { pos: Int, color: Color }  // pos: 0-100000
GradientFill { stops, angle, path_type, rot_with_shape, fill_to_l/t/r/b, tile_flip }
BlipFill { rid, stretch, src_l/t/r/b, tile_tx/ty/sx/sy, tile_flip, tile_algn }
PatternFill { prst, fg_color: Color, bg_color: Color }

TextParagraph { runs, align, level, spc_before, spc_after, mar_l, indent, line_spacing, bullet, bullet_auto, bullet_none, bullet_font, bullet_size, bullet_color, bullet_img_rid, tab_stops, rtl }
TextRun { text, bold, italic, font_size, color, font_face, ea_font, cs_font, sym_font, underline, strike, baseline, char_spacing, kern, cap, hlink_rid, hlink_mouse_over_rid }
BodyProps { anchor, l_ins, t_ins, r_ins, b_ins, auto_fit, font_scale, ln_spc_reduction, wrap, rot, vert, num_cols, col_spacing }

TableData { col_widths: Array[Int], rows: Array[TableRow], style_id: String, first_row/last_row/first_col/last_col/band_row/band_col: Bool }
TableStyleCell { fill, grad_fill, bdr_l/r/t/b_w, bdr_l/r/t/b_color, bold, italic, font_color }
TableStyleDef { id, whole_tbl, band1_h, band2_h, band1_v, band2_v, first_row, last_row, first_col, last_col: TableStyleCell }
TableRow { height: Int, cells: Array[TableCell] }
TableCell { paragraphs, fill: Color, grad_fill: GradientFill, grid_span, row_span: Int, v_merge: Bool, bdr_l/r/t/b_w: Int, bdr_l/r/t/b_color: Color, bdr_tl_br_w/color, bdr_bl_tr_w/color, mar_l/r/t/b: Int, anchor: String }

Color { r, g, b, alpha }  // r=-1 = none (sentinel), alpha: 0-255
ThemeData { dk1..fol_hlink: Color, major_font, minor_font, major_ea_font, minor_ea_font: String }
```

## Key files

| File | Purpose |
|------|---------|
| `src/ffi/ffi.mbt` | All JS→Wasm import declarations |
| `src/xml/xml.mbt` | Generic XML parser (DOM tree) |
| `src/ooxml/ooxml.mbt` | OOXML types (`SlideData`, `Shape`, etc.) + Color/HSL/modifier utilities |
| `src/ooxml/ooxml_theme.mbt` | Theme parser + ColorMap + master/layout parsers |
| `src/ooxml/ooxml_text.mbt` | Text body parsing (paragraphs, runs, bodyPr) |
| `src/ooxml/ooxml_parse.mbt` | Shape/Slide/Fill parsing + rels + slide size |
| `src/renderer/renderer.mbt` | Constants + helpers + Shape/Table rendering + public API |
| `src/renderer/renderer_text.mbt` | Text rendering (bullets, wrapping, tabs, height) |
| `src/renderer/renderer_fill.mbt` | Gradient/pattern fill SVG rendering |
| `src/renderer/renderer_geom.mbt` | Preset geometry evaluator (guide formulas → SVG path) |
| `src/svg_parser/svg_parser.mbt` | SVG (with `data-ooxml-*`) → SlideData |
| `src/serializer/serializer.mbt` | SlideData → OOXML slide XML |
| `src/main/main.mbt` | Wasm exports, slide cache (`g_slides`), global state |
| `src/main/main_inherit.mbt` | Placeholder inheritance + text style defaults |
| `src/main/moon.pkg.json` | Export list + `use-js-builtin-string: true` |
| `lib/index.ts` | Library public API re-exports |
| `lib/pptx-renderer.ts` | `PptxRenderer` class (core API) |
| `lib/wasm-compat.ts` | 3-tier Wasm js-string builtins fallback |
| `lib/zip.ts` | ZIP extraction and building |
| `lib/utils.ts` | bytesToBase64, crc32 utilities |
| `web/host.js` | Legacy JS host (kept for reference; demo uses `dist/`) |
| `web/index.html` | Browser demo UI |
| `test_fixtures/minimal.pptx` | 2-slide test fixture |
| `test_fixtures/test_features.pptx` | Feature regression test fixture (generated) |
| `test_fixtures/gen_test_features.py` | Python script to regenerate test_features.pptx |
| `test_fixtures/test_node.mjs` | Node.js test suite (ZIP + XML structure assertions) |

## Adding new OOXML features — required workflow

When implementing a new OOXML feature (e.g. gradient fill, shadow, connector), **always** update all three layers and add tests:

### 1. Implementation (MoonBit)
Follow the round-trip pipeline — update each relevant file:
- `src/ooxml/ooxml.mbt`: Data model (struct/field definitions)
- `src/ooxml/ooxml_parse.mbt`: XML parser for shapes, fills, transforms
- `src/ooxml/ooxml_text.mbt`: Text body/paragraph/run parsing (if text-related)
- `src/ooxml/ooxml_theme.mbt`: Theme/master/layout parsing (if theme-related)
- `src/renderer/renderer.mbt`: Shape/table SVG rendering + `data-ooxml-*` attributes
- `src/renderer/renderer_text.mbt`: Text SVG rendering (if text-related)
- `src/renderer/renderer_fill.mbt`: Gradient/pattern/blip fill rendering (if fill-related)
- `src/svg_parser/svg_parser.mbt`: `data-ooxml-*` → SlideData round-trip parsing
- `src/serializer/serializer.mbt`: SlideData → OOXML XML serialization
- `src/main/main.mbt`: Wasm exports, global state
- `src/main/main_inherit.mbt`: Placeholder inheritance + text style defaults

### 2. Test fixture (`gen_test_features.py`)
- Add new slide(s) to `gen_test_features.py` exercising the feature
- Update the docstring at the top of the file with the new slide number/description
- Run `python3 test_fixtures/gen_test_features.py` to regenerate `test_features.pptx`
- The `set_gradient_fill()` helper shows how to inject raw XML into shapes via lxml

### 3. Test assertions (`test_node.mjs`)
- Update `slide count = N` assertion to match new total
- Update iteration bounds (`for (let i = 1; i <= N; ...)`) for slide existence and .rels checks
- Add a new test section verifying the XML structure of the new slides
- Run `node test_fixtures/test_node.mjs` to confirm all tests pass

### 4. Verification checklist
```bash
python3 test_fixtures/gen_test_features.py  # Regenerate PPTX
moon build --target wasm-gc --release       # Wasm build (0 errors)
npm run build                               # Full build (Wasm + TypeScript)
node test_fixtures/test_node.mjs            # All tests pass
# Browser: http://localhost:8765/web/index.html  # Visual check
```
