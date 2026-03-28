# pptx-svg

PPTX and SVG round-trip conversion library. Runs entirely in the browser with zero dependencies.
- [Demo Site for GitHub Pages](https://t-ujiie-g.github.io/pptx-svg/)

[Japanese / 日本語](README.ja.md)

## Features

- **PPTX to SVG**: Convert PowerPoint slides to high-quality SVG
- **SVG to PPTX**: Edit SVG and export back to a valid .pptx file (lossless round-trip)
- **Browser-native**: No server required. ZIP, OOXML parsing, SVG generation all run client-side
- **Zero dependencies**: About 230KB Wasm binary, no npm dependencies
- **Framework-agnostic**: Works with React, Vue, Svelte, vanilla JS, or any framework

## Install

```bash
npm install pptx-svg
```

## Quick Start

```ts
import { PptxRenderer } from 'pptx-svg';

const renderer = new PptxRenderer();
await renderer.init();                        // Wasm loaded automatically

const file = await fetch('presentation.pptx');
await renderer.loadPptx(await file.arrayBuffer());

const svgString = renderer.renderSlideSvg(0); // Slide 1 as SVG
document.getElementById('viewer').innerHTML = svgString;
```

### React

```tsx
import { useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';

function SlideViewer({ pptxBuffer }: { pptxBuffer: ArrayBuffer }) {
  const [svg, setSvg] = useState('');
  const rendererRef = useRef<PptxRenderer | null>(null);

  useEffect(() => {
    const renderer = new PptxRenderer();
    rendererRef.current = renderer;
    renderer.init()
      .then(() => renderer.loadPptx(pptxBuffer))
      .then(() => setSvg(renderer.renderSlideSvg(0)));
  }, [pptxBuffer]);

  return <div dangerouslySetInnerHTML={{ __html: svg }} />;
}
```

### Vanilla JS (no bundler)

```html
<script type="importmap">
{ "imports": { "pptx-svg": "https://cdn.jsdelivr.net/npm/pptx-svg/dist/index.js" } }
</script>
<script type="module">
  import { PptxRenderer } from 'pptx-svg';
  const renderer = new PptxRenderer();
  await renderer.init();
  // ...
</script>
```

See [`examples/`](examples/) for complete working examples.
- [Demo Site for GitHub Pages](https://t-ujiie-g.github.io/pptx-svg/)

## API Reference

### `PptxRenderer`

```ts
import { PptxRenderer } from 'pptx-svg';

const renderer = new PptxRenderer(options?);
```

**Options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `measureText` | `(text, fontFace, fontSizePx) => number` | Canvas 2D | Custom text width measurement. |
| `fontFallbacks` | `Record<string, string[]>` | (built-in) | Custom font fallback mappings. Merged with built-in defaults. |
| `logLevel` | `'silent' \| 'error' \| 'warn' \| 'info' \| 'debug'` | `'error'` | Console output verbosity. |

**Methods:**

| Method | Returns | Description |
|--------|---------|-------------|
| `init(wasmSource?)` | `Promise<void>` | Load the Wasm module. No arguments needed (auto-resolved). Pass URL or ArrayBuffer to override. |
| `loadPptx(buffer)` | `Promise<{ slideCount }>` | Load a PPTX file from ArrayBuffer. |
| `getSlideCount()` | `number` | Number of slides. |
| `isSlideHidden(idx)` | `boolean` | Check if a slide is hidden (`show="0"`). |
| `renderSlideSvg(idx)` | `string` | Render slide as SVG string (0-indexed). |
| `updateSlideFromSvg(idx, svg)` | `string` | Update slide data from edited SVG. Returns `"OK"` or `"ERROR:..."`. |
| `getSlideOoxml(idx)` | `string` | Get OOXML XML for a slide. |
| `exportPptx()` | `Promise<ArrayBuffer>` | Export as .pptx file with modifications applied. |
| `getSlideXmlRaw(idx)` | `string` | Raw slide XML (for debugging). |
| `getEntryList()` | `string[]` | All ZIP entry paths (for debugging). |

**Shape-level Editing Methods:**

| Method | Returns | Description |
|--------|---------|-------------|
| `renderShapeSvg(slideIdx, shapeIdx)` | `string` | Render a single shape as SVG fragment. |
| `updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot)` | `string` | Update position/size/rotation (EMU). Returns re-rendered SVG. |
| `updateShapeText(slideIdx, shapeIdx, paraIdx, runIdx, text)` | `string` | Update text content. Returns re-rendered SVG. |
| `updateShapeFill(slideIdx, shapeIdx, r, g, b)` | `string` | Update solid fill color (0-255). Returns re-rendered SVG. |

All `update*` methods modify the cached SlideData in-place, mark the slide as modified for export, and return the re-rendered shape SVG. See [`docs/editing-guide.md`](docs/editing-guide.md) for usage patterns.

**Notes & Comments:**

| Method | Returns | Description |
|--------|---------|-------------|
| `getSlideNotes(idx)` | `string[]` | Speaker notes as array of paragraph strings. |
| `getSlideComments(idx)` | `SlideComment[]` | Comments with text, author ID, date, and position. |
| `getCommentAuthors()` | `CommentAuthor[]` | All comment authors (id, name, initials). |

Notes and comments are automatically preserved in round-trip export.

**Unit Conversion Helpers:**

```ts
import { pxToEmu, emuToPx, ptToHundredths, hundredthsToPt, degreesToOoxml, ooxmlToDegrees } from 'pptx-svg';

pxToEmu(100)          // 952500 EMU
emuToPx(914400)       // 96 px
ptToHundredths(18)    // 1800
hundredthsToPt(1800)  // 18
degreesToOoxml(90)    // 5400000
ooxmlToDegrees(5400000) // 90
```

**SVG DOM Helpers:**

```ts
import { findShapeElement, getShapeTransform, getAllShapes, getSlideScale } from 'pptx-svg';

const shapes = getAllShapes(svgElement);           // All shape <g> elements
const g = findShapeElement(svgElement, 0);         // Shape by index
const transform = getShapeTransform(g);            // { x, y, cx, cy, rot } in EMU
const scale = getSlideScale(svgElement);           // EMU per SVG pixel
```

## Supported Features

### Fully Supported

- **Shapes**: AutoShape (rect, ellipse, roundRect, line, ~154 preset geometries), custom geometry (`a:custGeom`), connectors (straight/elbow/curved)
- **Text**: Paragraphs, runs, bullets (char/auto/image), fonts (Latin/EA/CS/Symbol), bold/italic/underline/strikethrough, superscript/subscript, character spacing, kerning, capitalization, hyperlinks, tabs, RTL
- **Text body**: Vertical alignment, margins, auto-fit, font scale, rotation, vertical text, multi-column, text warp (prstTxWarp)
- **Fill**: Solid color, gradient (linear/radial with stops), pattern (48 presets), image fill (stretch/tile/crop)
- **Stroke**: 11 dash patterns, 5 arrow types, line cap/join, compound lines, gradient/pattern stroke
- **Effects**: Outer shadow, inner shadow, preset shadow, glow, soft edge, reflection, blur (all via SVG filters)
- **Images**: PNG/JPEG/GIF/SVG, crop, alpha, brightness/contrast, duotone, color change
- **Tables**: Cell merge (grid span, row span), borders (including diagonal), margins, anchoring, table styles, conditional formatting (banded rows/cols, first/last row/col)
- **Charts**: Bar (clustered/stacked/percentStacked), Line, Pie, Doughnut, Scatter, Area, Radar, Bubble, Stock, Surface, OfPie (13 types), data labels, data points, trendlines, error bars, composite charts
- **Group shapes**: Recursive nesting with coordinate transforms
- **Theme**: 12 theme colors, font scheme, all color modifiers (tint, shade, saturation, luminance, etc.)
- **Master/Layout inheritance**: Placeholder inheritance, `p:clrMapOvr`
- **Background**: Solid, gradient, image, pattern backgrounds
- **3D**: Data preservation for round-trip (bevel, extrusion, contour, material, camera, lighting)
- **Placeholder auto content**: Slide number, date, footer
- **Speaker notes**: Read via `getSlideNotes()`, preserved in round-trip export
- **Comments**: Read via `getSlideComments()` / `getCommentAuthors()`, preserved in round-trip export
- **SmartArt**: Fallback shapes from `mc:AlternateContent` rendered; `mc:Choice` (DiagramML) preserved for round-trip
- **OLE / Embedded objects**: Fallback image from `p:oleObj` rendered; original XML preserved for round-trip
- **Media** (video/audio): Poster frame image rendered; original XML preserved for round-trip
- **EMF / WMF images**: Converted to SVG at runtime via built-in converter
- **Math equations** (OMML `m:oMath`): SVG rendering of fractions, radicals, integrals, matrices, accents, and operators; original XML preserved for round-trip

### Supported with Limitations
- **TIFF images** - binary preserved for round-trip; not supported by browser `<img>` in all browsers
- **Embedded fonts** - binary preserved for round-trip; uses system font fallback for rendering

### Data Preservation (no visual rendering)

- **Animations** (`p:timing`) - preserved in round-trip export; static rendering only
- **Transitions** (`p:transition`) - preserved in round-trip export; static rendering only
- **Hidden slides** - detected via `isSlideHidden()` API; `show="0"` preserved in round-trip export

### Out of Scope

- **Macros / VBA** - not supported for security reasons

## SVG Output Format

The generated SVG embeds `data-ooxml-*` attributes that preserve all OOXML metadata for round-trip conversion. See [`docs/svg-specification.md`](docs/svg-specification.md) for the complete attribute reference.

## Architecture

```
[Browser]
  PptxRenderer (TypeScript)
    ├── ZIP extraction (DecompressionStream)
    ├── ZIP building (CompressionStream + CRC-32)
    └── FFI ─── WebAssembly GC (MoonBit)
                  ├── XML parser
                  ├── OOXML parser (types, theme, text, shapes, charts)
                  ├── SVG renderer (shapes, text, fill, geometry, charts)
                  ├── SVG parser (data-ooxml-* → SlideData)
                  └── OOXML serializer (SlideData → XML)
```

## Development

### Prerequisites

- [MoonBit toolchain](https://moonbitlang.com/download/)
- Node.js 18+

### Build

```bash
npm run build          # Wasm + TypeScript + copy wasm to dist/
```

### Test

```bash
npm test               # All tests (MoonBit unit + Node.js integration)
npm run test:moon      # MoonBit unit tests only
npm run test:node      # Node.js integration tests only
```

### Browser Test

```bash
python3 -m http.server 8765 --directory .
# Open http://localhost:8765/web/index.html
```

## Release

Releases are published to npm via GitHub Actions when a version tag is pushed:

```bash
# Update version in package.json, then:
git tag v0.1.0
git push origin v0.1.0
```

Requires `NPM_TOKEN` secret configured in GitHub repository settings.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes following the existing code style
4. Add MoonBit unit tests in `src/*/..._test.mbt` and/or integration tests in `test_fixtures/`
5. Run `npm run build && npm test` to verify
6. Submit a pull request

## License

[MIT](LICENSE)
