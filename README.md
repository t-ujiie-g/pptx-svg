# pptx-svg

A browser-based PPTX viewer/editor built with MoonBit (wasm-gc).
ZIP extraction, OOXML processing, SVG rendering, and PPTX export — all client-side, no server required.

[Japanese / 日本語](README.ja.md)

## Features

- **Round-trip conversion**: PPTX → SVG → edit → PPTX export
- **No server required**: ZIP extraction, OOXML parsing, SVG generation, ZIP rebuilding all run in the browser
- **Lossless round-trip**: `data-ooxml-*` attributes embedded in SVG preserve OOXML metadata
- **Lightweight**: ~24KB Wasm binary, zero external dependencies

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Logic | [MoonBit](https://moonbitlang.com/) → WebAssembly GC (wasm-gc) |
| Rendering | SVG (with `data-ooxml-*` attributes) |
| ZIP handling | Browser-native `DecompressionStream` / `CompressionStream` API (JS) |
| String FFI | `use-js-builtin-string: true` (MoonBit String = JS String, zero-cost) |
| Host layer | Plain JavaScript ES Modules |

## Architecture

```
[Browser]
  ┌─────────────────────────────────────────────────┐
  │  web/index.html                                 │
  │  lib/ → dist/  ← PptxRenderer class            │
  │    │                                            │
  │    ├─ ZIP extraction (DecompressionStream)       │
  │    ├─ ZIP building (CompressionStream + CRC-32)  │
  │    │                                            │
  │    └─ FFI ──────────────────────────────┐      │
  │                                          │      │
  │  [WebAssembly GC]                        │      │
  │  _build/.../main.wasm                    │      │
  │    src/ffi/         ← FFI declarations   │      │
  │    src/xml/         ← Generic XML parser │      │
  │    src/ooxml/       ← OOXML types+parser │      │
  │    src/renderer/    ← SlideData → SVG    │      │
  │    src/svg_parser/  ← SVG → SlideData    │      │
  │    src/serializer/  ← SlideData → XML    │      │
  │    src/main/        ← Public API ────────┘      │
  └─────────────────────────────────────────────────┘
```

**Data flow (Round-trip):**
1. User drops a .pptx file
2. JS parses and decompresses the ZIP, storing entries in a Map
3. `render_slide_svg(idx)` → SVG with `data-ooxml-*` attributes
4. (Edit SVG in browser)
5. `update_slide_from_svg(idx, svg)` → update cached SlideData
6. `exportPptx()` → rebuild ZIP with modified slide XML → download .pptx

## Quick Start

### Prerequisites

- [MoonBit toolchain](https://moonbitlang.com/download/) (`moon` command)
- Node.js 18+ (for building TypeScript and running tests)
- Chrome 111+ / Firefox 120+ / Safari 17+

### Build

```bash
# Build Wasm
moon build --target wasm-gc --release
# → _build/wasm-gc/release/build/main/main.wasm (~24KB)

# Build TypeScript library
tsc
# → dist/

# Build everything (Wasm + TypeScript)
npm run build
```

### Development Server

```bash
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html
```

### Tests

```bash
node test_fixtures/test_node.mjs
```

## Project Structure

```
pptx-svg/
├── moon.mod.json                  # MoonBit project config (no external deps)
├── package.json                   # npm package definition
├── src/                           # MoonBit (Wasm-GC)
│   ├── ffi/ffi.mbt               # JS host FFI declarations
│   ├── xml/xml.mbt               # Generic XML parser (DOM tree)
│   ├── ooxml/
│   │   ├── ooxml.mbt             # OOXML types + Color/HSL utilities
│   │   ├── ooxml_theme.mbt       # Theme parser + ColorMap + master/layout
│   │   ├── ooxml_text.mbt        # Text body/paragraph/run parsing
│   │   └── ooxml_parse.mbt       # Shape/Slide/Fill parsing + rels
│   ├── renderer/
│   │   ├── renderer.mbt          # Shape/Table SVG rendering + public API
│   │   ├── renderer_text.mbt     # Text SVG rendering (bullets, wrapping, tabs)
│   │   └── renderer_fill.mbt     # Gradient/pattern fill SVG rendering
│   ├── svg_parser/svg_parser.mbt # SVG → SlideData (reverse transform)
│   ├── serializer/serializer.mbt # SlideData → OOXML slide XML
│   └── main/
│       ├── main.mbt              # Wasm export API + slide cache
│       └── main_inherit.mbt      # Placeholder inheritance + text defaults
├── lib/                           # TypeScript library source
│   ├── index.ts                   # Public API re-exports
│   ├── pptx-renderer.ts          # PptxRenderer class (core API)
│   ├── wasm-compat.ts            # 3-tier Wasm js-string fallback
│   ├── zip.ts                    # ZIP extraction / building
│   └── utils.ts                  # bytesToBase64, crc32
├── dist/                          # Compiled JS + .d.ts (tsc output)
├── web/
│   ├── host.js                   # Legacy JS host (reference only)
│   └── index.html                # Demo UI (imports from dist/)
└── test_fixtures/
    ├── minimal.pptx              # 2-slide minimal test PPTX
    ├── test_features.pptx        # Feature regression test fixture
    ├── gen_test_features.py      # Python script to regenerate test fixture
    └── test_node.mjs             # Node.js test suite (JS layer)
```

**Module dependencies (no cycles):**
```
main → renderer   → ooxml → xml
     → svg_parser → ooxml → xml
     → serializer → ooxml
     → ffi
```

## API Reference

### Wasm Exports

| Function | Returns | Description |
|----------|---------|-------------|
| `initialize_pptx()` | `"OK:<count>"` or `"ERROR:..."` | Initialize PPTX and get slide count |
| `get_slide_count()` | `Int` | Number of slides |
| `get_slide_xml_raw(idx)` | `String` | Raw slide XML |
| `get_entry_list()` | `String` | ZIP entry list (newline-separated) |
| `render_slide_svg(idx)` | `String` | SVG with `data-ooxml-*` attributes |
| `update_slide_from_svg(idx, svg)` | `"OK"` or `"ERROR:..."` | Update SlideData from SVG |
| `get_slide_ooxml(idx)` | `String` | OOXML slide XML (regenerated if modified) |
| `get_modified_entries()` | `String` | Modified entries (`path\tcontent\n` format) |

### JS API (PptxRenderer class)

```javascript
await renderer.init(wasmUrl)              // Initialize Wasm module
await renderer.loadPptx(arrayBuffer)      // Load PPTX → { slideCount }
renderer.renderSlideSvg(slideIdx)         // Get SVG string
renderer.updateSlideFromSvg(idx, svg)     // Update internal data from SVG
renderer.getSlideOoxml(idx)               // Get OOXML XML
await renderer.exportPptx()               // Export as PPTX ArrayBuffer
```

## Known Limitations

- SmartArt / Charts render as gray fallback
- EMF/WMF images not supported
- Animations and transitions are ignored
- Does not run in Node.js (wasm-gc is browser-only)

## License

MIT
