# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Typecheck (fast — run after every meaningful edit)
moon check --target wasm-gc --deny-warn

# Build Wasm (output: _build/wasm-gc/release/build/main/main.wasm, ~400KB)
moon build --target wasm-gc --release

# MoonBit format
moon fmt

# Build TypeScript library (output: dist/)
tsc

# Build everything (Wasm + TypeScript)
npm run build

# Run all tests (MoonBit unit + Node.js integration)
npm test

# Run MoonBit unit tests only (xml, ooxml, renderer, svg_parser, serializer)
npm run test:moon

# Run Node.js integration tests only
npm run test:node

# Serve for browser testing
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html
```

## Architecture

**Separation of concerns:**
- **TypeScript library** (`lib/` → `dist/`): ZIP parsing/building, DEFLATE, Wasm lifecycle, CRC-32, EMF→SVG conversion, `PptxRenderer` class
- **MoonBit** (`src/`): OOXML parsing, SVG generation, SVG→SlideData parsing, OOXML serialization
- **Demo** (`web/`): Browser demo UI (imports from `dist/`)

**FFI boundary:**
- JS pre-decompresses all ZIP entries → stores in `Map<path, string>` and `Map<path, Uint8Array>`
- MoonBit calls `ffi_get_file(path)` to pull individual files on demand
- MoonBit exports (read-only): `initialize_pptx`, `get_slide_count`, `is_slide_hidden`, `get_slide_xml_raw`, `get_entry_list`, `render_slide_svg`, `update_slide_from_svg`, `get_slide_ooxml`, `get_modified_entries`
- MoonBit exports (editing): `render_shape_svg`, `update_shape_transform`, `update_shape_text`, `update_shape_fill`, `delete_shape`, `add_shape`, `add_shape_text`, `duplicate_shape`, `update_shape_gradient_fill`, `update_shape_stroke`, `add_paragraph`, `delete_paragraph`, `add_run`, `delete_run`, `update_text_run_style`, `update_text_run_font_size`, `update_text_run_color`, `update_text_run_font`, `update_paragraph_align`, `update_text_run_decoration`, `add_picture_shape`, `replace_picture_rid`
- MoonBit exports (history): `restore_slide_ooxml`
- MoonBit exports (inline text editing): `get_text_layout`, `hit_test_text`, `replace_text_range`
- MoonBit exports (z-order): `bring_to_front`, `send_to_back`, `bring_forward`, `send_backward`
- MoonBit exports (multi-transform): `update_shapes_transform`
- MoonBit exports (copy/paste): `get_shape_ooxml`, `add_shape_from_ooxml`
- MoonBit exports (table editing): `update_table_cell_text`, `add_table_row`, `delete_table_row`, `add_table_column`, `delete_table_column`
- Full export list: see `src/main/moon.pkg`

**Lazy slide parse (`g_slides` cache).** `initialize_pptx` fills `g_slides` with empty placeholders (`shapes: []`); the real parse + placeholder inheritance happens on the first `render_slide_svg(idx)`, which caches the resolved `SlideData` and sets `g_parsed[idx]`. Editing exports read `g_slides` directly, so any new editing export that reads the cache **must** route through `with_shape`/`with_run` (which call `ensure_slide_parsed`) or call `ensure_slide_parsed(slide_idx)` itself — otherwise it silently no-ops on a slide that was never rendered (this was the 0.5.10 bug). Don't add per-method `renderSlideSvg()` "ensure parsed" calls in the TS layer; the Wasm boundary is the single source of truth.

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

**Watch Int32 overflow in geometry math.** `Int` is 32-bit. `(x2 - x1) * adj / 100000` overflows when both operands are large in opposite signs (e.g. Google Slides writes connector adj1 = -39687500, paired with multi-million-EMU spans). Use `to_int64()` for the multiplication and divide back, or split as `span / 100 * adj / 1000`. See `bend_offset` in `renderer.mbt`.

**Watch integer truncation in `px(emu, scale)`.** `px` is integer division: small EMU values relative to `scale` (≈12700 for a 960px wide 16:9 slide) round down to 0. This matters for group children whose coordinates live in `chExt` space — if `chExt` is small (e.g. 10000 EMU mapped onto a multi-million-EMU group), `px(child_coord, outer_scale)` truncates every dimension to 0px and `render_shape`'s `cx_p <= 0` guard drops the shape, producing an empty `<g transform="..."></g>`. `render_group` works around this by computing a finer-grained `child_scale = min(scale·chExt_x/cx, scale·chExt_y/cy)` (Int64 internally), capped at the outer scale, and compensating with an `(cx·child_scale)/(chExt·scale)` factor inside the SVG `scale()` transform.

## MoonBit idiom (current compiler)

The sources track current MoonBit idiom; match it in new code.

- **Iterate collections with `for x in xs`**, not a manual index. When the index itself is needed, use a range loop: `for i in 0..<n`, `0..<=n`, `n>..0`, `n>=..0`.
- **A range loop evaluates its bounds once**, unlike a re-checked `while` condition. If the body pushes to / removes from the collection it is iterating, keep the `while`.
- **`while` is still correct** for loops with a body-local step (`i = i + char_step(c)` — the step is unavailable in a `for` header), for a `continue` that must *skip* the increment, and when the index is read after the loop.
- **`if opt is Some(x) { … }`** instead of `match opt { Some(x) => … None => () }`.
- **Dispatch on a string with `match`, not an `if s == "a" … else if` chain**: string-literal or-patterns (`"a" | "b" => …`), and array patterns to bind operands at the same time (`match tokens { ["cos", x, ang, ..] => … }`, `match tokens[i:] { ["M", x, y, ..] => … }`). A slice pattern also replaces a manual `i + n < len` arity check.
- **`xs.is_empty()` / `!xs.is_empty()`**, never `xs.length() == 0` / `> 0`.
- Method-call syntax for stdlib conversions: `c.to_int()`, not `Char::to_int(c)`.
- **Top-level declarations are separated by a `///|` block marker.** Each block is independent and order-irrelevant, so a refactor can move blocks one at a time. New top-level items get one too; `moon fmt` keeps them in place.

**Don't guess an API — verify it.** The compiler moves faster than any snapshot of its docs, and a plausible-looking signature that no longer exists costs a build cycle. Check with `moon ide doc <name>` (or `moon ide outline`) before using an unfamiliar stdlib function, and reach for the `moonbit-orientation` / `moonbit-agent-guide` skills for language-level questions. Deprecated APIs already found this way: `unsafe_char_at` → `get_char`, `assert_eq!` → `assert_eq`.

**`moon check` does not surface deprecation warnings** — they appear under `moon build` / `moon test`. A green `moon check --deny-warn` is not proof the tree is warning-free; the release build in the verification checklist is.

## MoonBit unit tests

Tests are in `src/*/..._test.mbt` files and run via `moon test --target js` with FFI stubs.

**Why JS target?** MoonBit FFI functions (`pptx_ffi.*`) are unresolved in the wasm-gc test runner. The JS target compiles to Node.js and allows injecting stubs via `NODE_OPTIONS='--require ./test_fixtures/ffi_stub.js'`.

**Why not remove FFI from xml?** `Char::to_string()` and `String::make(1, c)` both use `wasm:js-string "fromCharCodeArray"` which breaks Tier-2/3 browser polyfill compatibility. The `ffi_char_code_to_str` FFI (→ `String.fromCharCode`) only uses the polyfillable `concat` path.

**Test-only imports:** Use `import { ... } for "test"` in `moon.pkg` to add dependencies needed only by test files (e.g. svg_parser in renderer tests).

**Adding tests:** Place test files in the same package directory as `<name>_test.mbt`. These are black-box tests: refer to the package under test by its qualified name (`@renderer.render_slide_svg`, `@ooxml.Color::white()`), exactly as for any other imported package — an unqualified reference triggers warning 0025 `test_unqualified_package`, which only `moon test` reports. Use `assert_eq(actual, expected)` (not `assert_eq!` which is deprecated). For snapshot testing use `inspect!(value, content="expected")`.

## Browser compatibility and string constants

`use-js-builtin-string: true` in `src/main/moon.pkg` generates Wasm that imports:
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
SlideData { slide_size: SlideSize, background: Color, bg_grad: GradientFill, bg_blip_fill: BlipFill, bg_patt_fill: PatternFill, shapes: Array[Shape], transition_xml: String, timing_xml: String, hidden: Bool }
Shape { kind: ShapeKind, transform: ShapeTransform,
  fill: Color, grad_fill: GradientFill, blip_fill: BlipFill, patt_fill: PatternFill,
  stroke: Color, stroke_w: Int,
  stroke_dash: String, stroke_cap: String, stroke_join: String, stroke_miter_limit: Int,
  stroke_head_type: String, stroke_head_w: String, stroke_head_len: String,
  stroke_tail_type: String, stroke_tail_w: String, stroke_tail_len: String,
  stroke_cmpd: String, stroke_no_fill: Bool,
  stroke_grad_fill: GradientFill, stroke_patt_fill: PatternFill,
  paragraphs: Array[TextParagraph], body_props: BodyProps, ph_type: String, ph_idx: Int,
  st_cxn_id: Int, st_cxn_idx: Int, end_cxn_id: Int, end_cxn_idx: Int,
  sh_link_rid: String, sh_link_hover_rid: String,
  mc_choice_xml: String, ole_xml: String,
  effects: EffectList, scene_3d: Scene3d, sp_3d: Shape3d,
  lst_style: TextStyleGroup, font_ref_color: Color }
  // lst_style: a:lstStyle from p:txBody, used as a placeholder-inheritance source.
  // font_ref_color: resolved color from <p:style>/<a:fontRef>; applied to text runs
  //   that still lack an explicit color after all other inheritance (last-resort fallback).

ShapeKind = AutoShape(ShapeGeom) | Picture(String) | TableShape(TableData) | GroupShape(GroupShapeData) | ChartShape(ChartData) | Other
ShapeGeom = Rect | Ellipse | RoundRect(Int) | Line | Connector(String, Array[Int]) | Other(String, Array[Int]) | Custom(CustomGeomData)
  // RoundRect carries adj (0-100000); default 16667 (= 16.667%) per ECMA-376.
GroupShapeData { ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy: Int, children: Array[Shape] }
CustomGeomData { gdlst: String, paths: String, path_w: Int, path_h: Int, rect_l, rect_t, rect_r, rect_b: String, cxn_lst: String }
ShapeTransform { x, y, cx, cy, rot, flip_h, flip_v }  // all EMU
StrokeProps { color, width, dash, cap, join, miter_limit, head_type/w/len, tail_type/w/len, cmpd, no_fill, grad_fill, patt_fill }

GradientStop { pos: Int, color: Color }  // pos: 0-100000
GradientFill { stops, angle, path_type, rot_with_shape, fill_to_l/t/r/b, tile_flip }
BlipFill { rid, stretch, src_l/t/r/b, tile_tx/ty/sx/sy, tile_flip, tile_algn, alpha, svg_rid, bright, contrast, duotone_1/2: Color, clr_from/to: Color }
PatternFill { prst, fg_color: Color, bg_color: Color }

EffectList { outer_shadow: OuterShadow, inner_shadow: InnerShadow, glow: Glow, soft_edge: SoftEdge, reflection: Reflection, blur: Blur, prst_shadow: PresetShadow, fill_overlay: FillOverlay }
OuterShadow { blur_rad, dist, dir: Int, color: Color, sx, sy: Int, algn: String, rot_with_shape: Bool }
InnerShadow { blur_rad, dist, dir: Int, color: Color }
Glow { rad: Int, color: Color }
SoftEdge { rad: Int }
Blur { rad: Int }
PresetShadow { prst: String, dist, dir: Int, color: Color }
FillOverlay { blend: String, color: Color }
Reflection { blur_rad, dist, dir, st_alpha, end_alpha, fade_dir, sx, sy: Int, algn: String, rot_with_shape: Bool }

Bevel { w, h: Int, prst: String }
Shape3d { bevel_t, bevel_b: Bevel, extrusion_h, contour_w: Int, extrusion_clr, contour_clr: Color, prst_material: String, z: Int }
Scene3d { camera_prst, light_rig, light_dir: String }

TextParagraph { runs, align, level, spc_before, spc_after, mar_l, indent, line_spacing, bullet, bullet_auto, bullet_none, bullet_font, bullet_size, bullet_color, bullet_img_rid, tab_stops, rtl, epr_font_size, epr_color, epr_font_face, epr_ea_font }  // epr_* = endParaRPr fallbacks, applied as last-resort in main_inherit.mbt after layout/master inheritance
TextRun { text, bold, bold_explicit, italic, font_size, color, font_face, ea_font, cs_font, sym_font, underline, underline_color: Color, strike, baseline, char_spacing, kern, cap, hlink_rid, hlink_mouse_over_rid, effects: EffectList, outline_color: Color, outline_w: Int, text_grad_fill: GradientFill, text_patt_fill: PatternFill, math_xml: String }
  // underline_color: a:uFill/a:uLn color (Color::none() = follow text color); rendered as CSS text-decoration-color.
BodyProps { anchor, l_ins, t_ins, r_ins, b_ins, auto_fit, font_scale, ln_spc_reduction, wrap, rot, vert, num_cols, col_spacing, warp_prst: String, warp_av1, warp_av2: Int }

TableData { col_widths: Array[Int], rows: Array[TableRow], style_id: String, first_row/last_row/first_col/last_col/band_row/band_col: Bool }
TableStyleCell { fill, grad_fill, bdr_l/r/t/b_w, bdr_l/r/t/b_color, bold, italic, font_color }
TableStyleDef { id, whole_tbl, band1_h, band2_h, band1_v, band2_v, first_row, last_row, first_col, last_col: TableStyleCell }
TableRow { height: Int, cells: Array[TableCell] }
TableCell { paragraphs, fill: Color, grad_fill: GradientFill, grid_span, row_span: Int, v_merge, h_merge: Bool, bdr_l/r/t/b_w: Int, bdr_l/r/t/b_color: Color, bdr_tl_br_w/color, bdr_bl_tr_w/color, mar_l/r/t/b: Int, anchor: String }

Color { r, g, b, alpha }  // r=-1 = none (sentinel), alpha: 0-255
ThemeData { dk1..fol_hlink: Color, major_font, minor_font, major_ea_font, minor_ea_font: String, fill_style_xmls, ln_style_xmls: Array[String], clr_map: ColorMap }  // fmtScheme entries (raw XML with phClr placeholders) — resolved by parse_sp when a shape's <p:style>/<a:fillRef>/<a:lnRef> has no explicit fill/line. `clr_map` aliases logical names (bg1/tx1/bg2/tx2, accents, hlink) onto physical slots in `resolve_scheme_color`; physical references (dk1/lt1/...) bypass the cmap.
LevelTextDefaults { font_size, bold, italic, color, font_face, ea_font, align, mar_l, indent, bullet, bullet_auto, bullet_none, bullet_font, bullet_size, bullet_color, line_spacing, spc_before, spc_after }

ChartData { groups: Array[ChartGroup], axes: Array[ChartAxis], title: String, legend: ChartLegend, style: Int, chart_xml: String, view_3d: ChartView3D }
ChartKind = BarChart | LineChart | PieChart | DoughnutChart | ScatterChart | AreaChart | RadarChart | BubbleChart | StockChart | SurfaceChart | OfPieChart | WaterfallChart | TreemapChart | SunburstChart | HistogramChart | BoxWhiskerChart | FunnelChart
ChartGroup { chart_type: ChartKind, series: Array[ChartSeries], bar_dir, grouping: String, gap_width, overlap: Int, vary_colors: Bool, hole_size: Int, scatter_style: String, ax_ids: Array[Int], data_labels: ChartDataLabels, of_pie_type: String, split_pos: Int, wireframe: Bool, subtotals: Array[Int] }
ChartSeries { idx, order: Int, title: String, sp_pr: ChartSpPr, cat, val, x_val, y_val, bubble_size: AxisDataSource, smooth: Bool, explosion: Int, data_points: Array[ChartDataPoint], trendlines: Array[ChartTrendline], err_bars: ChartErrBars, data_labels: ChartDataLabels }
ChartSpPr { fill: Color, grad_fill: GradientFill, patt_fill: PatternFill, stroke: Color, stroke_w: Int, no_fill: Bool }
ChartView3D { rot_x, rot_y, depth_percent: Int, r_ang_ax: Bool, perspective: Int }
ChartDataLabels { show_val, show_cat_name, show_ser_name, show_percent, show_leader_lines: Bool, separator: String }
ChartDataPoint { idx: Int, sp_pr: ChartSpPr }
ChartTrendline { trendline_type, name: String, order, period, forward, backward: Int, sp_pr: ChartSpPr }
ChartErrBars { err_dir, err_bar_type, err_val_type: String, val: Int, sp_pr: ChartSpPr }
ChartLegend { position: String, overlay, show: Bool }
AxisDataSource = NumSource(String, NumData) | StrSource(String, StrData) | NoData
NumData { format_code: String, points: Array[ChartPoint] }
StrData { points: Array[ChartPoint] }
ChartPoint { idx: Int, value: String }
ChartAxis { ax_id, cross_ax: Int, ax_pos: String, delete, is_val, major_gridlines, minor_gridlines: Bool, title, orientation, min_val, max_val, major_unit, num_fmt, tick_lbl_pos, cross_between: String, sp_pr: ChartSpPr }
```

## Key files

| File | Purpose |
|------|---------|
| `src/ffi/ffi.mbt` | All JS→Wasm import declarations |
| `src/xml/xml.mbt` | Generic XML parser (DOM tree) |
| `src/ooxml/ooxml.mbt` | OOXML types (`SlideData`, `Shape`, etc.) + Color/HSL/modifier utilities |
| `src/ooxml/ooxml_theme.mbt` | Theme parser + ColorMap + master/layout parsers |
| `src/ooxml/ooxml_text.mbt` | Text body parsing (paragraphs, runs, bodyPr) |
| `src/ooxml/ooxml_parse.mbt` | Shape/Slide/Fill parsing + rels + slide size + SmartArt cached-drawing parsing (`parse_diagram_drawing`/`parse_diagram_text_color`/`apply_diagram_text_color`) |
| `src/ooxml/ooxml_chart.mbt` | ChartML parser (c:chartSpace → ChartData) |
| `src/renderer/renderer.mbt` | Constants + helpers + Shape rendering + public API |
| `src/renderer/renderer_table.mbt` | Table SVG rendering (cell borders, merging, conditional formatting) |
| `src/renderer/renderer_text.mbt` | Text rendering (bullets, wrapping, tabs, height) + shared helpers (`solve_text_autofit`, `run_fs_px`, `run_ff`, `wrap_paragraph`, `measure_line_width`) |
| `src/renderer/renderer_text_layout.mbt` | Text geometry for inline editing: `build_text_layout` / `text_layout_to_json` / `hit_test_layout` (EMU per-char boxes) |
| `src/renderer/renderer_warp.mbt` | Text warp rendering (SVG `<textPath>` + transforms for prstTxWarp presets) |
| `src/renderer/renderer_math.mbt` | OMML math rendering (fractions, radicals, integrals, matrices → SVG) |
| `src/renderer/renderer_fill.mbt` | Gradient/pattern/blip fill + effect filter SVG rendering |
| `src/renderer/renderer_geom.mbt` | Guide-formula evaluator + path rendering (ECMA-376 §20.1.9). Formula and path strings are parsed once into typed enums (`Formula`, `PathCmd`, operands `Operand::Lit`/`Var`) and evaluated by exhaustive `match`; preset definitions are parsed on first use and cached (`preset_geom`). `sin`/`cos`/`tan x ang` = x·f(ang), `arcTo` angles are visual (not parametric), sweeps > 180° are split into two SVG arcs. A compact sub-path may start with `W pw ph` (its own `a:path w/h`), `F mode` (`a:path fill`) and `S false` (not stroked); `render_styled_paths` honours them — shade modes are black/white overlays, stroke-only sub-paths are drawn after the fills |
| `src/renderer/renderer_geom_presets.mbt` | **Generated — do not edit.** All 182 ECMA-376 presets (`get_preset_def`: avLst / gdLst / paths / text rect as compact strings), from `scripts/gen_preset_geom.py` |
| `scripts/gen_preset_geom.py` | Regenerates `renderer_geom_presets.mbt` from `presetShapeDefinitions.xml` (ECMA-376 Part 1 attachment; path passed as the argument, the XML is not vendored), then run `moon fmt`. Documents the compact string formats |
| `src/renderer/renderer_geom_wbtest.mbt` | White-box tests for the guide evaluator and arc conversion |
| `src/renderer/renderer_chart.mbt` | Chart SVG rendering (bar/line/pie/donut/scatter/area/radar/bubble/stock/surface/ofPie) |
| `src/svg_parser/svg_parser.mbt` | SVG (with `data-ooxml-*`) → SlideData |
| `src/serializer/serializer.mbt` | SlideData → OOXML slide XML |
| `src/main/main.mbt` | Wasm exports (read-only APIs), slide cache (`g_slides`), global state; post-parse resolution of charts (`resolve_chart_shapes`) and SmartArt diagrams (`resolve_diagram_shapes`) |
| `src/main/main_edit.mbt` | Shape/text/image editing API exports (CRUD, fill, stroke, text formatting, picture shapes); history (`restore_slide_ooxml`); z-order; multi-transform; copy/paste. Shared helpers: `with_shape`/`with_run`, `make_default_run`/`paragraph`, `array_remove_at`/`array_insert_at` |
| `src/main/main_text_edit.mbt` | Inline text editing exports: `get_text_layout`, `hit_test_text`, `replace_text_range` + `split_para_runs` |
| `src/main/main_table_edit.mbt` | Table editing exports: `update_table_cell_text`, `add_table_row`/`column`, `delete_table_row`/`column` + `with_table` |
| `src/main/main_inherit.mbt` | Placeholder inheritance + text style defaults (transforms, text styles, auto-content) |
| `src/main/moon.pkg` | Export list + `use-js-builtin-string: true` |
| `lib/index.ts` | Library public API re-exports |
| `lib/pptx-renderer.ts` | `PptxRenderer` class (core API) |
| `lib/wasm-compat.ts` | 3-tier Wasm js-string builtins fallback |
| `lib/zip.ts` | ZIP extraction and building |
| `lib/utils.ts` | bytesToBase64, crc32 utilities |
| `lib/font-fallbacks.ts` | Font fallback mappings (customizable via `PptxRendererOptions`) |
| `lib/emf-converter.ts` | Lightweight EMF→SVG converter (vector paths, text, bitmaps) |
| `lib/wmf-converter.ts` | Lightweight WMF→SVG converter (vector paths, text, bitmaps) |
| `docs/svg-specification.md` | SVG output format specification (`data-ooxml-*` attributes) |
| `web/index.html` | Browser demo UI |
| `src/xml/xml_test.mbt` | XML parser unit tests |
| `src/ooxml/ooxml_test.mbt` | OOXML types/parsing unit tests |
| `src/renderer/renderer_test.mbt` | Renderer + round-trip unit tests |
| `src/svg_parser/svg_parser_test.mbt` | SVG parser unit tests |
| `src/serializer/serializer_test.mbt` | Serializer unit tests |
| `test_fixtures/ffi_stub.js` | FFI stubs for MoonBit JS-target tests |
| `test_fixtures/minimal.pptx` | 2-slide test fixture |
| `test_fixtures/test_features.pptx` | Feature regression test fixture (generated) |
| `test_fixtures/gen_test_features.py` | Thin orchestrator — imports each `fixtures/slides_NN_*.py` category module in order, saves the presentation, then runs `_postprocess` |
| `test_fixtures/fixtures/_ctx.py` | Shared `prs`/`blank` state + cross-module helpers (`nsmap`, `set_fill_xml`) |
| `test_fixtures/fixtures/slides_NN_*.py` | Per-category slide builders (text basics, fills, shapes, tables, images, charts, misc, chartex, regressions) |
| `test_fixtures/fixtures/_postprocess.py` | Post-save ZIP patching (OMML, ChartEx, media, WMF) |
| `test_fixtures/tests/*.test.mjs` | Node `--test` suite split by feature category (structure, text, fills, shapes, tables, images, charts, misc, chartex, emf, wmf, utils, regressions) |
| `test_fixtures/tests/_helpers.mjs` | Shared test helpers (`expect`, `loadFeatures`, `hasTag`, `findRelTarget`, …) |
| `test_fixtures/test_node_compat.mjs` | Node.js editing API test suite (shape/text/image CRUD, round-trip export) |

## Adding new OOXML features — required workflow

When implementing a new OOXML feature (e.g. gradient fill, shadow, connector), **always** update all three layers and add tests:

### 1. Implementation (MoonBit)
Follow the round-trip pipeline — update each relevant file:
- `src/ooxml/ooxml.mbt`: Data model (struct/field definitions)
- `src/ooxml/ooxml_parse.mbt`: XML parser for shapes, fills, transforms
- `src/ooxml/ooxml_chart.mbt`: ChartML parser (if chart-related)
- `src/ooxml/ooxml_text.mbt`: Text body/paragraph/run parsing (if text-related)
- `src/ooxml/ooxml_theme.mbt`: Theme/master/layout parsing (if theme-related)
- `src/renderer/renderer.mbt`: Shape/table SVG rendering + `data-ooxml-*` attributes
- `src/renderer/renderer_text.mbt`: Text SVG rendering (if text-related)
- `src/renderer/renderer_fill.mbt`: Gradient/pattern/blip fill rendering (if fill-related)
- `src/renderer/renderer_chart.mbt`: Chart SVG rendering (if chart-related)
- `src/svg_parser/svg_parser.mbt`: `data-ooxml-*` → SlideData round-trip parsing
- `src/serializer/serializer.mbt`: SlideData → OOXML XML serialization
- `src/main/main.mbt`: Wasm exports (read-only), global state
- `src/main/main_edit.mbt`: Shape/text/image editing API exports
- `src/main/main_inherit.mbt`: Placeholder inheritance + text style defaults

### 2. MoonBit unit tests
- Add tests in the relevant `*_test.mbt` file (e.g. `src/ooxml/ooxml_test.mbt`, `src/renderer/renderer_test.mbt`)
- Test pure functions (color parsing, geometry, serialization) and round-trip (render → parse → compare)
- Run `npm run test:moon` to confirm all MoonBit tests pass

### 3. Test fixture (`test_fixtures/fixtures/slides_NN_*.py`)
- Append the new slide to the matching category module under `test_fixtures/fixtures/` (e.g. `slides_03_fills.py` for a new fill feature). Each module imports `prs`/`blank` from `_ctx` and mutates the shared presentation at import time.
- If the new slide needs a helper that crosses category modules, put it in `_ctx.py` (alongside `set_fill_xml`) and import it from `_ctx`.
- Run `python3 test_fixtures/gen_test_features.py` to regenerate `test_features.pptx` (the orchestrator imports each category module in order, then runs `_postprocess`).
- The `set_gradient_fill()` helper in `slides_03_fills.py` shows how to inject raw XML into shapes via lxml.

### 4. Test assertions (`test_fixtures/tests/*.test.mjs`)
- Update the slide-count assertion in `structure.test.mjs` to match the new total.
- Update iteration bounds (`for (let i = 1; i <= N; ...)`) in `structure.test.mjs` for slide existence and `.rels` checks.
- Add (or extend) a category file — e.g. a fill change goes in `fills.test.mjs`, a new chart in `charts.test.mjs`. Each file is a single top-level `test(...)` block that calls `resetAssertions()`/`finishAssertions()` from `_helpers.mjs`.
- Run `node --test test_fixtures/tests/*.test.mjs` to confirm all tests pass.

### 5. Verification checklist
```bash
python3 test_fixtures/gen_test_features.py  # Regenerate PPTX
npm run test:moon                           # MoonBit unit tests pass
moon build --target wasm-gc --release       # Wasm build (0 errors)
npm run build                               # Full build (Wasm + TypeScript)
npm run test:node                           # Node.js integration tests pass
# Browser: http://localhost:8765/web/index.html  # Visual check
```

## Refactoring checklist

When asked to "refactor" (リファクタリング), review against these six criteria and apply only low-risk, high-value changes — don't restructure working code for its own sake. If nothing needs doing, say so rather than forcing changes.

1. **Constant management / no magic numbers.** Numeric literals with domain meaning (EMU sizes, font fallbacks, default depths, sentinels) get a named `let` in the constants block (`renderer.mbt` for renderer, module top for others). Re-used literals must be a single shared constant, not copies.
2. **No duplication / dead code.** Identical logic in 2+ places → extract one shared helper (e.g. per-run font resolution → `run_fs_px`/`run_ff` in `renderer_text.mbt`, shared by `wrap_paragraph`/`measure_line_width`/`build_text_layout`). Remove unused params (don't paper over with `ignore(x)`), unreachable branches, and commented-out code.
3. **File splitting.** Keep files cohesive and roughly under ~1500 lines where practical. Same-package MoonBit files (`src/main/*.mbt`, `src/renderer/*.mbt`) can be split purely for organization with zero API risk — prefer one file per concern (e.g. `main_text_edit.mbt` for inline-text exports). Don't split the giant pre-existing renderers (`renderer_chart.mbt`, `render_text`) without a strong reason — high risk.
4. **Docs up to date.** After any change, sync `CLAUDE.md` (Key files table, FFI export list, data model), `docs/editing-guide.md`, `README.md`/`README.ja.md`, `CHANGELOG.md` (`## Unreleased`), and `TODO.md`. Version bumps in `package.json` are deferred to release (keep current; CHANGELOG accumulates under `## Unreleased`).
5. **Test coverage.** Every behavior has a MoonBit unit test (pure functions, round-trip) and/or a Node integration test (`test_node_compat.mjs`). After refactoring, the full suite (`npm test`) must stay green with unchanged counts unless tests were intentionally added.
6. **Comment hygiene.** See below.

### Comment hygiene

A comment must make sense to a reader who has never opened `TODO.md` and was not there when the code landed.

- **No roadmap / phase codes in code comments.** `E6.1`, `E2.5`, `P3`, `Phase 2`, `see TODO R1` mean nothing to an outside reader and go stale as soon as `TODO.md` is reorganized. Name the *thing* instead — "History support", "cross-slide copy/paste", "run-box layout". Stable, findable references are fine and should stay: ECMA-376 § numbers, GitHub issue `#N`, browser-version constraints. `TODO.md` and `CHANGELOG.md` keep the codes; the code does not.
- **Delete comments the code already says.** A comment restating the next line, narrating what was changed, or arguing the change to a reviewer is noise once merged. That history lives in git and `CHANGELOG.md`.
- **Keep comments that carry a constraint the code cannot show.** Spec rules, sentinel meanings (`-1 = unset`), units (EMU, 1/60000 deg, hundredths-of-pt), the browser/compiler limitations in this file, and deliberate deviations ("browsers ignore `text-decoration-color` on SVG `<text>`") all stay. `///` doc comments on Wasm exports and public helpers stay mandatory.
- **Keep comments short and non-repeating.** When several fields share one mechanism, explain it once above the group instead of repeating it per field. Trailing per-field comments in `ooxml.mbt` structs are the established convention here and stay — they carry units and sentinels — but they must be one short line, never a section marker or a paragraph wedged between fields.
- **Section markers** (`// ── Title ────…`) label a block of related declarations; keep the title descriptive and the rule padded to the same column as its neighbours.

After refactoring, run `npm test` and confirm rendering output is unchanged (the renderer/Node suites guard this).
