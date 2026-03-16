# Changelog

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
