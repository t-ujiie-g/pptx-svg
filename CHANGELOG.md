# Changelog

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
