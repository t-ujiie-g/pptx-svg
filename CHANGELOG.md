# Changelog

## 0.1.0 (Unreleased)

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
- Charts: 13 chart types with data labels, trendlines, error bars
- Theme resolution: 12 colors, font scheme, all color modifiers
- Master/Layout placeholder inheritance
- 3-tier Wasm js-string builtins compatibility (Chrome 111+)
- Zero external dependencies
