# pptx-svg

PPTX and SVG round-trip conversion library. Runs entirely in the browser with zero dependencies.
- [Demo Site for GitHub Pages](https://t-ujiie-g.github.io/pptx-svg/)

[Japanese / 日本語](README.ja.md)

## Features

- **PPTX to SVG**: Convert PowerPoint slides to high-quality SVG
- **SVG to PPTX**: Edit SVG and export back to a valid .pptx file (lossless round-trip)
- **Browser-native**: No server required. ZIP, OOXML parsing, SVG generation all run client-side
- **Zero dependencies**: ~200KB Wasm binary, no npm dependencies
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

| Option | Type | Description |
|--------|------|-------------|
| `measureText` | `(text, fontFace, fontSizePt) => number` | Custom text width measurement. Defaults to Canvas 2D. |

**Methods:**

| Method | Returns | Description |
|--------|---------|-------------|
| `init(wasmSource?)` | `Promise<void>` | Load the Wasm module. No arguments needed (auto-resolved). Pass URL or ArrayBuffer to override. |
| `loadPptx(buffer)` | `Promise<{ slideCount }>` | Load a PPTX file from ArrayBuffer. |
| `getSlideCount()` | `number` | Number of slides. |
| `renderSlideSvg(idx)` | `string` | Render slide as SVG string (0-indexed). |
| `updateSlideFromSvg(idx, svg)` | `string` | Update slide data from edited SVG. Returns `"OK"` or `"ERROR:..."`. |
| `getSlideOoxml(idx)` | `string` | Get OOXML XML for a slide. |
| `exportPptx()` | `Promise<ArrayBuffer>` | Export as .pptx file with modifications applied. |
| `getSlideXmlRaw(idx)` | `string` | Raw slide XML (for debugging). |
| `getEntryList()` | `string[]` | All ZIP entry paths (for debugging). |

## Supported Features

### Fully Supported

- **Shapes**: AutoShape (rect, ellipse, roundRect, line, ~154 preset geometries), custom geometry (`a:custGeom`), connectors (straight/elbow/curved)
- **Text**: Paragraphs, runs, bullets (char/auto/image), fonts (Latin/EA/CS/Symbol), bold/italic/underline/strikethrough, superscript/subscript, character spacing, kerning, capitalization, hyperlinks, tabs, RTL
- **Text body**: Vertical alignment, margins, auto-fit, font scale, rotation, vertical text, multi-column, text warp (prstTxWarp)
- **Fill**: Solid color, gradient (linear/radial with stops), pattern (48 presets), image fill (stretch/tile/crop)
- **Stroke**: 11 dash patterns, 5 arrow types, line cap/join, compound lines, gradient/pattern stroke
- **Effects**: Outer shadow, inner shadow, glow, soft edge, reflection (all via SVG filters)
- **Images**: PNG/JPEG/GIF/SVG, crop, alpha, brightness/contrast, duotone, color change
- **Tables**: Cell merge (grid span, row span), borders (including diagonal), margins, anchoring, table styles, conditional formatting (banded rows/cols, first/last row/col)
- **Charts**: Bar, Line, Pie, Doughnut, Scatter, Area, Radar, Bubble, Stock, Surface, OfPie (13 types), data labels, data points, trendlines, error bars, composite charts
- **Group shapes**: Recursive nesting with coordinate transforms
- **Theme**: 12 theme colors, font scheme, all color modifiers (tint, shade, saturation, luminance, etc.)
- **Master/Layout inheritance**: Placeholder inheritance, `p:clrMapOvr`
- **Background**: Solid, gradient, image, pattern backgrounds
- **3D**: Data preservation for round-trip (bevel, extrusion, contour, material, camera, lighting)
- **Placeholder auto content**: Slide number, date, footer

### Not Yet Supported

- **SmartArt** (`dgm:*` DiagramML) - planned: fallback image display
- **OLE / Embedded objects** (`p:oleObj`) - planned: fallback image display
- **Media** (video/audio) - planned: poster frame display
- **EMF/WMF images** - cannot be decoded in browser
- **TIFF images** - not supported by browser `<img>`
- **Math equations** (OMML `m:oMath`) - planned: plain text fallback
- **Embedded fonts** - uses system font fallback
- **Speaker notes** (`p:notes`) - planned
- **Comments** (`p:cmAuthorLst` / `p:cmLst`) - planned

### Out of Scope

- **Animations** (`p:timing`) - static rendering only
- **Transitions** (`p:transition`) - static rendering only
- **Macros / VBA** - not supported for security reasons

## SVG Output Format

The generated SVG embeds `data-ooxml-*` attributes that preserve all OOXML metadata for round-trip conversion. See [`docs/svg-specification.md`](docs/svg-specification.md) for the complete attribute reference.

## Browser Compatibility

| Browser | Minimum Version | Notes |
|---------|----------------|-------|
| Chrome | 111+ | Full support (Tier 3 Wasm fallback) |
| Firefox | 120+ | Full support |
| Safari | 17+ | Full support |
| Edge | 111+ | Same as Chrome |
| Node.js | Not supported | Wasm-GC requires browser runtime |

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
npm test               # Node.js test suite (ZIP + XML structure)
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
# GitHub Actions builds, tests, and publishes to npm
```

Requires `NPM_TOKEN` secret configured in GitHub repository settings.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes following the existing code style
4. Add tests in `test_fixtures/gen_test_features.py` and `test_fixtures/test_node.mjs`
5. Run `npm run build && npm test` to verify
6. Submit a pull request

## License

[MIT](LICENSE)
